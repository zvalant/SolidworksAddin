using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.LinkLabel;



namespace SolidworksAddTest
{
    [ComVisible(true)]
    [Guid("4099c769-bd7d-49a6-ac97-1ec1e38ddcf9")]
    [ProgId("SolidworksAddTest.EcnProcessService")]



    
        public partial class EcnProcessService : UserControl
        {
            private SWTestRP parentAddin;
        public SolidworksService ThisSolidworksService { get; set; }
        public ReleaseValidationService ThisReleaseValidationService { get; set; }
        private PartValidation ThisPartValidation {  get; set; }
        //private AssemblyValidation ThisAssemblyValidation { get; set; }
        //private DrawingValidation ThisDrawingValidation { get; set; }

        public EcnRelease ThisEcnRelease { get; set; }

        public ReleaseReport ThisReleaseReport { get; set; }

        public Utility ThisUtility { get; set; }    

        public EcnProcessService()
            {
            InitializeComponent();
            this.BackColor = Color.White;
            }
        public void InitalizeRelease()
        {
            //testReleaseNumber will need to be changed to be ui input dependent along with other release options.
            string testReleaseNumber = "52635_test";



            int releaseMode = 0;
            ThisSolidworksService = new SolidworksService(parentAddin.SolidWorksApplication);
            ThisEcnRelease = new EcnRelease(testReleaseNumber, releaseMode);
            ThisReleaseReport = new ReleaseReport(testReleaseNumber, releaseMode);
            ThisReleaseValidationService = new ReleaseValidationService(ThisSolidworksService);
            ThisUtility = new Utility();
        }
        public void SetParentAddin(SWTestRP parent)
        {
            parentAddin = parent;
            //test values to be removed later
            // Add a simple control to verify it's working
        }

        private void GenerateButton_Click(object sender, EventArgs e)
        {
            DateTime startTime = DateTime.Now;
            int release = RunRelease();
            DateTime endTime = DateTime.Now;
            int runtime = (int)(endTime - startTime).TotalMilliseconds;

        }
        private int RunRelease() 
        {
            //var thisRelease = new EcnRelease("50001", 0);
            //modify ecn location after inital test
            InitalizeRelease();

            var testReleaseList = new List<string>();
            ClearEcnLocalFolder(ThisEcnRelease.ReleaseFolderTemp, true);
            CopyEcnFolder(ThisEcnRelease.ReleaseFolderSrc, ThisEcnRelease.ReleaseFolderTemp);

            string folderPath = ThisEcnRelease.ReleaseFolderTemp;
            
            //prevent active open files from being included
            foreach (string file in Directory.GetFiles(folderPath))
            {
                if (ThisUtility.GetFileWithExt(file)[0] == '~') 
                {
                    continue;
                }
                testReleaseList.Add(file);
            }


            //NEED TO DO**provide validation that files in folder are for ecn and all files required for ecn are present
            
            //Update data models with correct files
            for (int i = 0; i < testReleaseList.Count; i++)
            {
                string currentFile = testReleaseList[i];
                var currentFileObj = new EcnFile();
                currentFileObj.FilePath = testReleaseList[i];
                currentFileObj.FileName = ThisUtility.GetFileWithExt(currentFile);
                ThisEcnRelease.FileNames.Add(currentFileObj.FileName);
                ThisEcnRelease.AddFile(currentFileObj, currentFileObj.FileName);
                switch (Path.GetExtension(currentFile))
                {
                    case SolidworksService.PARTFILEEXT:
                        currentFileObj.DocumentType = swDocumentTypes_e.swDocPART;
                        break;
                    case SolidworksService.ASSEMBLYFILEEXT:
                        currentFileObj.DocumentType = swDocumentTypes_e.swDocASSEMBLY;
                        break;
                    case SolidworksService.DRAWINGFILEEXT:
                        currentFileObj.DocumentType = swDocumentTypes_e.swDocDRAWING;
                        break;
                    default:
                        break;
                }

                //SetSearchPaths(currentFile);
               //ReleaseFile(currentFileObj);
            }
            List<string> reportLines = new List<string>();
            int invalidReferences = 0;
            string sectionHeader = "Reference Validation";
            bool validationStatus = true;
            ThisReleaseReport.WriteSectionHeader(sectionHeader);
            foreach ( string fileName in ThisEcnRelease.Files.Keys)

            {
                
                EcnFile currentFile = ThisEcnRelease.Files[fileName];
                SearchAndDependenciesValidationResult DependenciesValidation = SearchPathAndGraphGeneration(ThisEcnRelease.Files[fileName].FilePath, ThisEcnRelease, currentFile);
                if (DependenciesValidation.FoundInvalidDependencies)
                {
                    invalidReferences++;
                    validationStatus = false;
                    foreach (Tuple<string,string> referencePair in DependenciesValidation.InvalidDependencies)
                    {
                        string parentFile = referencePair.Item1;
                        string referenceFile = referencePair.Item2;
                        reportLines.Add($"{referenceFile} is wrong reference for {parentFile}");
                    }
                }
                currentFile.InsertSearchPaths(DependenciesValidation.SearchPaths);

                ThisEcnRelease.LeafFiles.Add(currentFile);
                
            }

            if (validationStatus)
            {
                reportLines.Insert(0,("Validation Status: Passed"));
                reportLines.Add("");
            }
            else
            {
                reportLines.Insert(0, ("Validation Status: Failed"));
                reportLines.Add("");
                ThisReleaseReport.WriteToReport(reportLines);
                goto FinishRelease;
            }
            ThisReleaseReport.WriteToReport(reportLines);

            //clear report lines since reference section is now complete

            /*
             * This is the Release File Validation Section
             */
            sectionHeader = "File Release Validation";
            ThisReleaseReport.WriteSectionHeader(sectionHeader);

            foreach (EcnFile file in ThisEcnRelease.Files.Values)

            {
                foreach (EcnFile parentFile in file.Parents)
                {
                    parentFile.LoadedFilesRemaining++;
                    
                    if (ThisEcnRelease.LeafFiles.Contains(parentFile))
                    {
                        ThisEcnRelease.RemoveLeafFile(parentFile);
                    }
                }
                
            }
            
            foreach (EcnFile file in ThisEcnRelease.LeafFiles)
            {
                ThisEcnRelease.ProcessFilesPush(file);

            }
            
            while (ThisEcnRelease.ProcessingFileQueue.Count > 0)
            {
                int releaseStatus = 0;
                var currentFile = ThisEcnRelease.ProcessFilesPop();
                if (ThisEcnRelease.CompletedFiles.Contains(currentFile) || currentFile.LoadedFilesRemaining>0)
                {
                    continue;
                }
    
                ThisSolidworksService.ApplySearchPaths(currentFile.SearchPaths);
                ThisEcnRelease.PushOpenFileStack(currentFile);
                releaseStatus = ReleaseFile(currentFile);
                if (releaseStatus != 0)
                {
                    goto FinishRelease;
                }
                ThisEcnRelease.AddCompletedFile(currentFile);
                foreach (EcnFile ParentFile in currentFile.Parents)
                {
                    if (ParentFile.DocumentType == swDocumentTypes_e.swDocDRAWING && ParentFile.LoadedFilesRemaining == 1)
                    {
                        releaseStatus = ReleaseFile(ParentFile);
                        if(releaseStatus != 0)
                        {
                            goto FinishRelease;
                        }

                        ThisSolidworksService.CloseFile(ParentFile.FilePath);
                        ThisEcnRelease.AddCompletedFile(ParentFile);

                    }
                    else 
                    {
                        ThisEcnRelease.ProcessFilesPush(ParentFile);
                    }
                    ParentFile.LoadedFilesRemaining--;

                }
                
                while (ThisEcnRelease.OpenFilesStack.Count > 0)
                {
                    int parentsCompleted = 0;
                    EcnFile openFileCurrent = ThisEcnRelease.OpenFilesStack.Peek();
                    foreach (EcnFile parent in openFileCurrent.Parents)
                    {
                        if (ThisEcnRelease.CompletedFiles.Contains(parent))
                        {
                            parentsCompleted++;
                        }
                    }
                    if (parentsCompleted >= openFileCurrent.Parents.Count)
                    {
                        ThisEcnRelease.OpenFilesStack.Pop();
                        ThisSolidworksService.CloseFile(openFileCurrent.FilePath);
                    }
                    else 
                    {
                        break;
                    }
                }
                reportLines.Add("Validation Status: Passed");
                ThisReleaseReport.WriteToReport(reportLines);
                goto FinishRelease;
                
            }

        
        FinishRelease:
            ThisSolidworksService.CloseAllDocuments();
            ThisReleaseReport.FinishReport();
            ThisReleaseReport.OpenReport();
            ClearEcnLocalFolder(folderPath, false);
            
            /*
            foreach (EcnFile file in ThisEcnRelease.LeafFiles)
            {
                FileTraversal(file, file.FilePath);
            }
            */
            return 0;
        }

        private SearchAndDependenciesValidationResult SearchPathAndGraphGeneration(string filepath, EcnRelease thisRelease, EcnFile currentFile)
        {
            List<string> reportLines = new List<string>();
            //This will be replaced with ecn folder 
            if (parentAddin == null)
            {
                MessageBox.Show("Parent add-in is not set.");
                return null;
            }
            HashSet<string> dependencies = new HashSet<string>();
            Dictionary<string, int> folderCount = new Dictionary<string, int>();
            var folderPriority = new List<(int count, string folderPath)>();
            SearchAndDependenciesValidationResult currentDependciesValidationResult = new SearchAndDependenciesValidationResult(thisRelease.ReleaseFolderTemp,thisRelease.ReleaseFolderSrc);
            GetDependenciesAndParents(filepath, dependencies, folderCount , currentFile, currentDependciesValidationResult);
            if (currentDependciesValidationResult.CriticalError)
            {
                reportLines.Add("Critical Errors Found while gathering dependencies");
                return currentDependciesValidationResult;
                
            }
            int dependenciesCount = dependencies.Count;
            if (currentDependciesValidationResult.FoundInvalidDependencies)
            {
                foreach (Tuple<string,string> dependent in currentDependciesValidationResult.InvalidDependencies)
                {
                    string parent = dependent.Item1;
                    string reference = dependent.Item2;

                }
                
                    
                
            }
            foreach (KeyValuePair<string, int> path in folderCount)
            {
                folderPriority.Add((path.Value, path.Key));

            }
            folderPriority.Sort();
            folderPriority.Reverse();
            List<string> searchPathPriority = new List<string>();
            searchPathPriority.Add(thisRelease.ReleaseFolderTemp);
            searchPathPriority.Add(thisRelease.ReleaseFolderSrc);
            for (int i = 0; i < folderPriority.Count; i++)
            {
                searchPathPriority.Add($"M:/{folderPriority[i].folderPath}");
            }

            string[] archiveFolders = Directory.GetDirectories(@"M:/");
            foreach (string archiveFolder in archiveFolders)
            {
                searchPathPriority.Add(archiveFolder);
            }
            currentDependciesValidationResult.SearchPaths = searchPathPriority;


            return currentDependciesValidationResult;
        }

        //may want to split up leaves from this in future
        private void GetDependenciesAndParents(string docName, HashSet<string> dependencies, Dictionary<string, int> folderCount, EcnFile currentFile, SearchAndDependenciesValidationResult currentDependenciesValidationResult)
        {
            try
            {
                SearchAndDependenciesValidation currentValidation = new SearchAndDependenciesValidation(ThisSolidworksService);
                HashSet<Tuple<string, string>> wrongExtRef = new HashSet<Tuple<string,string>>();
                List<string> reportLines = new List<string>();
                string[] DepList = ThisSolidworksService.GetDocumentDependencies(docName);
                if (DepList == null)
                { 
                    return;
                }
      
                for (int i = 0; i < DepList.Length; i += 2)
                {
                    string currentDependent = DepList[i + 1];
                    string currentFileDependent = ThisUtility.GetFileWithExt(currentDependent);
                    if (!dependencies.Contains(currentDependent))
                    {

                        currentValidation.DependentValidation(currentDependent,docName, folderCount, currentDependenciesValidationResult, ThisEcnRelease);
                        dependencies.Add(currentDependent);
                        if (ThisEcnRelease.Files.ContainsKey(currentFileDependent))
                        {
                            var dependentFileObj = ThisEcnRelease.Files[currentFileDependent];
                            dependentFileObj.InsertParent(currentFile);
                            GetDependenciesAndParents(currentDependent, dependencies, folderCount, dependentFileObj, currentDependenciesValidationResult);
                        }
                        else 
                        {
                            GetDependenciesAndParents(currentDependent, dependencies, folderCount, currentFile, currentDependenciesValidationResult);
                        }


                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error Generating Dependencies: {ex.Message}: {docName}");
            }
            return;
        }
   

        private int ReleaseFile(EcnFile file)
        {
            List<string> reportLines = new List<string>();
            int releaseResult = 0;
            swDocumentTypes_e docType = file.DocumentType;
            switch (docType)
            {
                case swDocumentTypes_e.swDocPART:
                    
                    ModelDoc2 activePart = ThisSolidworksService.OpenPart(file.FilePath);
                     PartValidation currentPartValidation = new PartValidation(ThisSolidworksService);

                    PartValidationResult partReleaseResult = currentPartValidation.RunPartValidation(activePart);
                    if (partReleaseResult.CriticalError)
                    {
                        reportLines.Add($"Validation Status: CRITICAL ERROR \n COULD NOT RUN VALIDATION");
                        releaseResult = 1;
                    }
                    else if (partReleaseResult.FoundFeatureErrors)
                    {
                        reportLines.Add($"Validation Status: Failed");
                        reportLines.Add(file.FilePath);

                        foreach (PartFeatureInfo partInfo in partReleaseResult.TotalSketchErrors)
                        {
                            //MessageBox.Show($"{partInfo.FeatureName}, {partInfo.SketchName} is not properly defined");
                            reportLines.Add($" \t{partInfo.ConfigurationName} - {partInfo.FeatureName}, {partInfo.SketchName} NOT PROPERLY DEFINED");
                            releaseResult = 1;

                        }
                    }
                    break;
                case swDocumentTypes_e.swDocASSEMBLY:
                    ModelDoc2 activeAssy = ThisSolidworksService.OpenAssembly(file.FilePath);
                    AssemblyValidationResult assemblyReleaseResult = ThisReleaseValidationService.CheckAssembly(activeAssy);
                    if (assemblyReleaseResult.CriticalError)
                    {
                        reportLines.Add($"Validation Status: CRITICAL ERROR \n COULD NOT RUN VALIDATION");
                        releaseResult = 1;

                    }
                    if (assemblyReleaseResult.FoundComponentErrors)

                    {
                        reportLines.Add($"Validation Status: Failed");
                        foreach (ComponentInfo componentError in assemblyReleaseResult.TotalComponentErrors)
                        {
                            reportLines.Add($"\t-{componentError.Name}({componentError.Configuration}) - {componentError.ConstraintStatus}");
                        }
                        releaseResult = 1;

                    }

                    break;
                case swDocumentTypes_e.swDocDRAWING:
                    docType = swDocumentTypes_e.swDocDRAWING;
                    ModelDoc2 activeDrawing = ThisSolidworksService.OpenDrawing(file.FilePath);
                    DrawingValidationResult drawingReleaseResult = ThisReleaseValidationService.CheckDrawing(activeDrawing, file.FilePath);
                    if (drawingReleaseResult.CriticalError)
                    {
                        reportLines.Add($"Validation Status: CRITICAL ERROR \n COULD NOT RUN VALIDATION");
                        releaseResult = 1;
                    }
                    else if(drawingReleaseResult.FoundDanglingAnnotations)
                    {
                        reportLines.Add($"Validation Status: Failed");
                        releaseResult = 1;
                        foreach (SheetInfo currentSheetInfo in drawingReleaseResult.TotalSheets) 
                        {
                            string currentSheetName = currentSheetInfo.Name;
                            if (!currentSheetInfo.FoundDanglingAnnotations)
                            {
                                continue;
                            }
                            foreach (ViewInfo currentViewInfo in currentSheetInfo.Views)
                            {
                                if (!currentViewInfo.FoundDanglingAnnotations) 
                                {
                                    continue;
                                }
                                string currentViewName = currentViewInfo.Name;
                                foreach (string annotationtype in currentViewInfo.AnnotationCount.Keys)
                                {
                                    reportLines.Add($"\t{currentSheetName}, {currentViewName} - {currentViewInfo.AnnotationCount[annotationtype]} dangling {annotationtype}(s)");
                                }
                            }
                        }


                    }
                    
                    if (drawingReleaseResult.FoundDanglingAnnotations)
                    {
                        ThisSolidworksService.DeleteAllSelections(activeDrawing);
                        ThisSolidworksService.MoveSWFile(file.FilePath, file.FileName, ThisEcnRelease.ReleaseFolderSrc);


                    }

                    break;
                default:
                    break;
            }
            //ThisSolidworksService.CloseFile(file.FilePath);
            if (releaseResult != 0)
            {
                ThisReleaseReport.WriteToReport(reportLines);
                return 1;
            }
            return 0;
        }
        // this would be used for DFS traversal if used in future
        private void FileTraversal(EcnFile currentFile, string filePath)
        {
            bool canClose = true;
            if (ThisEcnRelease.CompletedFiles.Contains(currentFile))
            {
                return;
            }
            if (currentFile.LoadedFilesRemaining > 0)
            {
                return;
            }
            ThisEcnRelease.AddCompletedFile(currentFile);
            ThisSolidworksService.ApplySearchPaths(currentFile.SearchPaths);
            ReleaseFile(currentFile);
            
            foreach (EcnFile parentFile in currentFile.Parents)
            {
                parentFile.LoadedFilesRemaining--;
                if (parentFile.LoadedFilesRemaining > 0) { 
                    canClose = false;
                    continue;
                }
                FileTraversal(parentFile, parentFile.FilePath);
                
                
                
            }
            
            if (canClose)
            {
                ThisSolidworksService.CloseFile(filePath);


            }
            
            //ThisSolidworksService.CloseFile(filePath);

        }
        public void CopyEcnFolder(string sourceFolder, string destFolder)
        {
            try
            {
                if (!Directory.Exists(sourceFolder))
                {
                    MessageBox.Show($"Source folder does not exist {sourceFolder}");
                    return;
                }

                Directory.CreateDirectory(destFolder);

                foreach (string file in Directory.GetFiles(sourceFolder))
                {
                    string fileName = Path.GetFileName(file);
                    string destFile = Path.Combine(destFolder, fileName);
                    File.Copy(file, destFile, true);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying folder: {ex.Message}");
            }
        }
        public void ClearEcnLocalFolder(string tempFolder, bool startup)
        {
            try
            {
                if (Directory.Exists(tempFolder))
                {
                    Directory.Delete(tempFolder, true); 
                }

                if (startup)
                {
                    Directory.CreateDirectory(tempFolder);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error clearing temp folder: {ex.Message}");
            }
        }




    }
}

