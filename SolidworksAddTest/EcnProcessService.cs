using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Data;
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
        public EcnRelease ThisEcnRelease { get; set; }

        public ReleaseReport ThisReleaseReport { get; set; }



        public EcnProcessService()
            {
            InitializeComponent();
            this.BackColor = Color.White;
            }
        public void InitalizeRelease()
        {
            //testReleaseNumber will need to be changed to be ui input dependent along with other release options.
            string testReleaseNumber = "52585";
            int releaseMode = 0;
            ThisSolidworksService = new SolidworksService(parentAddin.SolidWorksApplication);
            ThisEcnRelease = new EcnRelease(testReleaseNumber, releaseMode);
            ThisReleaseReport = new ReleaseReport(testReleaseNumber, releaseMode);
            ThisReleaseValidationService = new ReleaseValidationService(ThisSolidworksService);

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
                if (GetFileWithExt(file)[0] == '~') 
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
                currentFileObj.FileName = GetFileWithExt(currentFile);
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
            foreach ( string fileName in ThisEcnRelease.Files.Keys)

            {
                EcnFile currentFile = ThisEcnRelease.Files[fileName];
                List<string> fileSearchPaths = GetSearchPath(ThisEcnRelease.Files[fileName].FilePath, ThisEcnRelease, currentFile);
                currentFile.InsertSearchPaths(fileSearchPaths);
                
            }

            
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

        private List<string> GetSearchPath(string filepath, EcnRelease thisRelease, EcnFile currentFile)
        {
            //This will be replaced with ecn folder 
            if (parentAddin == null)
            {
                MessageBox.Show("Parent add-in is not set.");
                return null;
            }
            HashSet<string> dependencies = new HashSet<string>();
            Dictionary<string, int> folderCount = new Dictionary<string, int>();
            var folderPriority = new List<(int count, string folderPath)>();
            GetDependenciesAndLeaves(filepath, dependencies, folderCount , currentFile);
            int dependenciesCount = dependencies.Count;
            foreach (KeyValuePair<string, int> path in folderCount)
            {
                folderPriority.Add((path.Value, path.Key));

            }
            folderPriority.Sort();
            folderPriority.Reverse();
            List<string> searchPathPriority = new List<string>();
            searchPathPriority.Add(thisRelease.ReleaseFolderTemp);
            for (int i = 0; i < folderPriority.Count; i++)
            {
                searchPathPriority.Add(folderPriority[i].folderPath);

            }

            string[] archiveFolders = Directory.GetDirectories(@"M:/");
            foreach (string archiveFolder in archiveFolders)
            {
                searchPathPriority.Add(archiveFolder);
            }



            return searchPathPriority;
        }

        //may want to split up leaves from this in future
        private void GetDependenciesAndLeaves(string DocName, HashSet<string> dependencies, Dictionary<string, int> folderCount, EcnFile currentFile)
        {
            try
            {

                string[] DepList = ThisSolidworksService.GetDocumentDependencies(DocName);

                if (DepList == null || DepList.Length == 0)
                {
                        ThisEcnRelease.AddLeafFile(currentFile);
                    return;
                }
                for (int i = 0; i < DepList.Length; i += 2)
                {
                    string currentDependent = DepList[i + 1];
                    string currentFileDependent = GetFileWithExt(currentDependent);

                    if (!dependencies.Contains(currentDependent))
                    {
                
                        ParsePath(currentDependent, folderCount);
                        dependencies.Add(currentDependent);
                        if (ThisEcnRelease.Files.ContainsKey(currentFileDependent))
                        {
                            var dependentFileObj = ThisEcnRelease.Files[currentFileDependent];
                            dependentFileObj.InsertParent(currentFile);
                            if (ThisEcnRelease.LeafFiles.Contains(currentFile))
                            {
                                ThisEcnRelease.RemoveLeafFile(currentFile);
                            }
                            GetDependenciesAndLeaves(currentDependent, dependencies, folderCount, dependentFileObj);
                        }
                        else 
                        {
                            GetDependenciesAndLeaves(currentDependent, dependencies, folderCount, currentFile);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error Generating Dependencies: {ex.Message}");
            }
            return;
        }
        private string GetFileWithExt(string docName)
        {
            string[] path = docName.Split(new char[] { '\\' });
            string fileName = path[path.Length - 1];
            return fileName;

        }

        private void ParsePath(string docName, Dictionary<string, int> folderCount)

        {
            string[] path = docName.Split(new char[] {'\\' });

            string archiveFolder = path[0]+"\\"+path[1];
            
            if (path[0] == "M:")
            {
                if (folderCount.ContainsKey(archiveFolder))
                {
                    folderCount[archiveFolder] += 1;
                }
                else 
                {
                    folderCount[archiveFolder] = 1;
                }
                
            }

        }
        

        private int ReleaseFile(EcnFile file)
        {
            List<string> reportLines = new List<string>();
            reportLines.Add($"{file.FilePath}");
            int releaseResult = 0;
            swDocumentTypes_e docType = file.DocumentType;
            switch (docType)
            {
                case swDocumentTypes_e.swDocPART:
                    ModelDoc2 activePart = ThisSolidworksService.OpenPart(file.FilePath);
                    PartValidationResult partReleaseResult = ThisReleaseValidationService.CheckPart(activePart, file.FilePath);
                    if (partReleaseResult.CriticalError)
                    {
                        reportLines.Add($"Validation Status: CRITICAL ERROR \n COULD NOT RUN VALIDATION");
                        releaseResult = 1;
                    }
                    else if (partReleaseResult.FoundFeatureErrors)
                    {
                        reportLines.Add($"Validation Status: Failed");

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
                    if (assemblyReleaseResult.CritalError)
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
                        ThisSolidworksService.MoveSWFile(ThisEcnRelease.ReleaseFolderTemp, ThisEcnRelease.ReleaseFolderSrc);
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
                    /*
                    if (drawingReleaseResult.FoundDanglingAnnotations)
                    {
                        ThisSolidworksService.DeleteAllSelections(activeDrawing);

                    }
                    */
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
            
            ThisSolidworksService.CloseFile(filePath);

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

