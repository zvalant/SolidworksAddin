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
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
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
        // Generate all required services/classes for release process
        public void InitalizeRelease(string ecnReleaseNumber)
        {

            bool isReadiness = true;
            ThisSolidworksService = new SolidworksService(parentAddin.SolidWorksApplication);
            ThisEcnRelease = new EcnRelease(ecnReleaseNumber, isReadiness);
            ThisReleaseReport = new ReleaseReport(ecnReleaseNumber, isReadiness);
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

            string ecnNumber = textBox1.Text.Trim();
            DateTime startTime = DateTime.Now;
            int release = RunRelease(ecnNumber);
            DateTime endTime = DateTime.Now;
            int runtime = (int)(endTime - startTime).TotalMilliseconds;

        }

        private int RunRelease(string ecnNumber) 
        {

            InitalizeRelease(ecnNumber);
            string folderPath = ThisEcnRelease.ReleaseFolderTemp;
            string sectionHeader;
            var FilesInECNFolder = new List<string>();
            ClearEcnLocalFolder(ThisEcnRelease.ReleaseFolderTemp, true);
            bool canCopyFolderOver = CopyEcnFolder(ThisEcnRelease.ReleaseFolderSrc, ThisEcnRelease.ReleaseFolderTemp);
            if (!canCopyFolderOver)
            {
                ThisReleaseReport.WriteToReportSingleline($"Cant Find Folder {ThisEcnRelease.ReleaseFolderSrc}");
                FinishRelease(folderPath);
                return 1;
            }
            // collect all files that are not actively open temps '~' is used to exclude active files
            foreach (string file in Directory.GetFiles(folderPath))
            {
                string fileName = ThisUtility.GetFileWithExt(file);
                string fileExtension = ThisUtility.GetFileExt(file);
                // will only add files if it contains an extension and is not an actively opened file
                if (fileName[0] != '~' && ThisEcnRelease.validReleaseExtensions.Contains(fileExtension))
                {
                    FilesInECNFolder.Add(file);
                }
            }
            sectionHeader = "ECN & Folder Validation";
            ThisReleaseReport.WriteSectionHeader(sectionHeader);

            // check ecn txt file and run comparative check to make sure appropriate files are in folder
            List<List<string>> ecnData;
            ecnData = ParseECNFile(ThisEcnRelease);
            Dictionary<string,string> ecnDataCurrentRevs = new Dictionary<string, string>();
            int partIdx = 0;
            int currentRevIdx = 2;

            foreach (List<string> partLine in ecnData)
            {
                ecnDataCurrentRevs[partLine[partIdx]] = partLine[currentRevIdx];
            }

            ECNFileValidation currentEcnFileValidation = new ECNFileValidation(ThisSolidworksService);
            EcnFileValidationResult fileValidationResult = new EcnFileValidationResult();
            try
            {
                fileValidationResult = currentEcnFileValidation.RunEcnFileValidation(ecnData, FilesInECNFolder, ThisEcnRelease.ReadinessForRelease);
            }
            catch (Exception e)
            {
                ThisReleaseReport.WriteToReportSingleline($"{e.StackTrace}");
            }
            if (fileValidationResult.CriticalError)
            {
                ThisReleaseReport.WriteToReportSingleline("CRITAL ERROR AT FILE VALIDATION");
                FinishRelease(folderPath);
                return 1;
            }
            if (fileValidationResult.ValidataionError)
            {
                ThisReleaseReport.WriteToReportSingleline($"Validation Status: Failed");

                if (fileValidationResult.FoundDuplicateECNFile)
                {

                    ThisReleaseReport.WriteToReportSingleline($"Found Duplicate numbers in txt file: {ThisEcnRelease.ReleaseTxtFile}");
                    foreach (string file in fileValidationResult.DuplicateEcnFile)
                    {
                        ThisReleaseReport.WriteToReportSingleline($"{file}");
                    }
                }
                if (fileValidationResult.FoundMissingDrawingFile)
                {
                    ThisReleaseReport.WriteToReportSingleline($"Release is missing following Drawing Documents:");
                    foreach (string file in fileValidationResult.MissingDrawingFile)
                    {
                        ThisReleaseReport.WriteToReportSingleline($"{file}");
                    }

                }
                if (fileValidationResult.FoundFailedApproval && !ThisEcnRelease.ReadinessForRelease)
                    ThisReleaseReport.WriteToReportSingleline($"Found Part(s) Missing QAD Approval");
                {
                    foreach (string file in fileValidationResult.FailedApproval)
                    {
                        ThisReleaseReport.WriteToReportSingleline($"{file}");
                    }
                    
                }
                FinishRelease(folderPath);
                return 1;
            }
            
            ThisReleaseReport.WriteToReportSingleline($"Validation Status: Success");
            if (ThisEcnRelease.ReadinessForRelease && fileValidationResult.FoundFailedApproval)
            {
                ThisReleaseReport.WriteToReportSingleline($"=".PadRight(90, '='));
                ThisReleaseReport.WriteToReportSingleline($"WARNING: FOLLOWING ISSSUES WILL NOT PASS OFFICIAL RELEASE:");
                ThisReleaseReport.WriteToReportSingleline($"=".PadRight(90, '='));

                if (fileValidationResult.FoundFailedApproval)
                {
                    ThisReleaseReport.WriteToReportSingleline($"Found Part(s) Missing QAD Approval");
                    foreach (string file in fileValidationResult.FailedApproval)
                    {
                        ThisReleaseReport.WriteToReportSingleline($"{file}");
                    }

                }
            }
            ThisSolidworksService.CloseAllDocuments();
            //prevent active open files from being included



            //NEED TO DO**provide validation that files in folder are for ecn and all files required for ecn are present
            
            //Update data models with correct files
            for (int i = 0; i < FilesInECNFolder.Count; i++)
            {
                string currentFile = FilesInECNFolder[i];
                string currentFileWOExt = Path.GetFileNameWithoutExtension( currentFile );
                var currentFileObj = new EcnFile();
 
                currentFileObj.FileName = ThisUtility.GetFileWithExt(currentFile);
                currentFileObj.FilePath = FilesInECNFolder[i];
                currentFileObj.FilePathSrc = ThisEcnRelease.ReleaseFolderSrc + "\\" + currentFileObj.FileName;
                if (ecnDataCurrentRevs.ContainsKey(currentFileWOExt))
                {
                    currentFileObj.Revision = ecnDataCurrentRevs[currentFileWOExt];
                }



                ThisEcnRelease.FileNames.Add(currentFileObj.FileName);
                ThisEcnRelease.AddFile(currentFileObj, currentFileObj.FileName);
                switch (Path.GetExtension(currentFile))
                {
                    case SolidworksService.PARTFILEEXT:
                        currentFileObj.SWDocumentType = swDocumentTypes_e.swDocPART;
                        break;
                    case SolidworksService.ASSEMBLYFILEEXT:
                        currentFileObj.SWDocumentType = swDocumentTypes_e.swDocASSEMBLY;
                        break;
                    case SolidworksService.DRAWINGFILEEXT:
                        currentFileObj.SWDocumentType = swDocumentTypes_e.swDocDRAWING;
                        break;
                    case SolidworksService.EXCELFILEEXT:
                        currentFileObj.IsExcelDocument = true;
                        break;

                    default:
                        MessageBox.Show($"{currentFile} is not Valid Release Document Type");
                        break;
                }

            }

            List<string> reportLines = new List<string>();
            Dictionary<string, string> wrongReferencePairs = new Dictionary<string, string>();
            int invalidReferences = 0;
            sectionHeader = "Reference Validation";
            bool validationStatus = true;
            ThisReleaseReport.WriteSectionHeader(sectionHeader);
            foreach ( string fileName in ThisEcnRelease.Files.Keys)
            {
                EcnFile currentFile = ThisEcnRelease.Files[fileName];
                SearchAndDependenciesValidationResult DependenciesValidation = SearchPathAndGraphGeneration(ThisEcnRelease.Files[fileName].FilePath, ThisEcnRelease, currentFile);
                string srcParentFilePath;
                string srcReferenceFilePath;
                if (DependenciesValidation.FoundInvalidDependencies)
                {
                    invalidReferences++;
                    validationStatus = false;
                    foreach (string parentFilePath in DependenciesValidation.InvalidDependencies.Keys)
                    {
                        srcParentFilePath = parentFilePath;
                 
                        string parentFileName = ThisUtility.GetFileWithExt(parentFilePath);
                        foreach (string referenceFilePath in DependenciesValidation.InvalidDependencies[parentFilePath])
                        {
                            srcReferenceFilePath = referenceFilePath;
                            string referenceFileName = ThisUtility.GetFileExt(referenceFilePath);
                            if (ThisEcnRelease.Files.ContainsKey(referenceFileName))
                            {
                                srcReferenceFilePath = ThisEcnRelease.Files[referenceFileName].FilePathSrc;
                            }
                            if (ThisEcnRelease.Files.ContainsKey(parentFileName))
                            {
                                srcParentFilePath = ThisEcnRelease.Files[parentFileName].FilePathSrc;
                            }
                            wrongReferencePairs[parentFilePath] = srcReferenceFilePath;
                        }
            
                    }
                }
                currentFile.InsertSearchPaths(DependenciesValidation.SearchPaths);
                ThisEcnRelease.LeafFiles.Add(currentFile);


            }
            foreach (string parentFilePath in wrongReferencePairs.Keys)
            {
                string srcReferenceFilePath = wrongReferencePairs[parentFilePath];
                reportLines.Add($"{srcReferenceFilePath} is wrong reference for {parentFilePath}");

            }
    


            ThisReleaseReport.WriteValidationStatus(validationStatus, reportLines);
            ThisReleaseReport.WriteToReportMultiline(reportLines);
            reportLines.Clear();
            if (!validationStatus)
            {
                FinishRelease(folderPath);
                return 1;
            }

            //clear report lines since reference section is now complete

            /*
             * This is the Release File Validation Section
             */
            reportLines.Clear();
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
            ThisReleaseReport.WriteToReportMultiline(reportLines);
            reportLines.Clear();
            foreach (EcnFile file in ThisEcnRelease.LeafFiles)
            {
                ThisEcnRelease.ProcessFilesPush(file);
            }
            
            while (ThisEcnRelease.ProcessingFileQueue.Count > 0)
            {
                foreach (EcnFile ecnFile in ThisEcnRelease.ProcessingFileQueue)
                {
                }
                int releaseStatus = 0;
                var currentFile = ThisEcnRelease.ProcessFilesPop();
                if (ThisEcnRelease.ReleasedFiles.Contains(currentFile) || currentFile.LoadedFilesRemaining>0)
                {
                    continue;
                }
                
                SolidworksServiceResult<bool> searchPathsResult = ThisSolidworksService.ApplySearchPaths(currentFile.SearchPaths);
                if (!searchPathsResult.Success)
                {
                    reportLines.Add($"Error applying search paths for {currentFile.FilePath}");
                    FinishRelease(folderPath);
                }
                ThisEcnRelease.PushOpenFileStack(currentFile);
                releaseStatus = ReleaseFile(currentFile);

                if (releaseStatus != 0)
                {
                    FinishRelease(folderPath);
                    return 1;
                }
                ThisEcnRelease.AddReleasedFile(currentFile);
                foreach (EcnFile parentFile in currentFile.Parents)
                {
                    if (parentFile.SWDocumentType == swDocumentTypes_e.swDocDRAWING && parentFile.LoadedFilesRemaining == 1)
                    {
                        releaseStatus = ReleaseFile(parentFile);
                        if (releaseStatus != 0)
                        {
                            FinishRelease(folderPath);
                            return 1;
                        }
                        
                        SolidworksServiceResult<bool> closeFileResult = ThisSolidworksService.CloseFile(parentFile.FilePath);
                        ThisEcnRelease.AddReleasedFile(parentFile);

                    }
                    else 
                    {
                        ThisEcnRelease.ProcessFilesPush(parentFile);
                    }
                    parentFile.LoadedFilesRemaining--;


                }

                foreach (EcnFile ecnFile in ThisEcnRelease.ProcessingFileQueue)
                {
                }
                while (ThisEcnRelease.OpenFilesStack.Count > 0)
                {
                    int parentsCompleted = 0;
                    EcnFile openFileCurrent = ThisEcnRelease.OpenFilesStack.Peek();
                    foreach (EcnFile parent in openFileCurrent.Parents)
                    {
                        if (ThisEcnRelease.ReleasedFiles.Contains(parent))
                        {
                            parentsCompleted++;
                        }
                    }
                    if (parentsCompleted >= openFileCurrent.Parents.Count)
                    {
                        ThisEcnRelease.OpenFilesStack.Pop();

                        SolidworksServiceResult<bool> closeFileResult = ThisSolidworksService.CloseFile(openFileCurrent.FilePath);

                    }
                    else 
                    {
                        break;
                    }
                }
      
                
            }
            reportLines.Add("Validation Status: Passed");
            ThisReleaseReport.WriteToReportMultiline(reportLines);
            FinishRelease(folderPath);




            /*
            foreach (EcnFile file in ThisEcnRelease.LeafFiles)
            {
                FileTraversal(file, file.FilePath);
            }
            */
            return 0;
        }
        private void FinishRelease(string folderPath)
        {
            ThisSolidworksService.CloseAllDocuments();
            ThisReleaseReport.FinishReport();
            ThisReleaseReport.OpenReport();
            ClearEcnLocalFolder(folderPath, false);
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
            List<string> reportLines = new List<string>();
            SearchAndDependenciesValidation currentValidation = new SearchAndDependenciesValidation(ThisSolidworksService);
            HashSet<Tuple<string, string>> wrongExtRef = new HashSet<Tuple<string, string>>();
            SolidworksServiceResult<string[]> getDependenciesResult = ThisSolidworksService.GetDocumentDependencies(docName);
            try
            {

                if (!getDependenciesResult.Success)
                {
                    reportLines.Add(getDependenciesResult.ErrorMessage);
                    ThisReleaseReport.WriteToReportMultiline(reportLines);
                    
                }
                string[] DepList = getDependenciesResult.response;
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
                reportLines.Add($"Error Generating Dependencies: {ex.Message}: {docName} {getDependenciesResult.response} stackTrace: {ex.StackTrace}");
            }
            ThisReleaseReport.WriteToReportMultiline( reportLines );    
            return;
        }
   

        private int ReleaseFile(EcnFile currentFileObj)
        {
            List<string> reportLines = new List<string>();
            int releaseResult = 0;
            swDocumentTypes_e docType = currentFileObj.SWDocumentType;
            switch (docType)
            {
                case swDocumentTypes_e.swDocPART:
                    
                    SolidworksServiceResult<ModelDoc2> openPartResult = ThisSolidworksService.OpenPart(currentFileObj.FilePath);
                    if (!openPartResult.Success)
                    {
                        reportLines.Add($" Error Opening Part {currentFileObj.FilePath}");
                        return 1;

                    }
                    ModelDoc2 activePart = openPartResult.response;
                    
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
                        reportLines.Add(currentFileObj.FilePath);
                        releaseResult = 1;


                        foreach (PartFeatureInfo partInfo in partReleaseResult.TotalSketchErrors)
                        {
                            //MessageBox.Show($"{partInfo.FeatureName}, {partInfo.SketchName} is not properly defined");
                            reportLines.Add($" \t{partInfo.ConfigurationName} - {partInfo.FeatureName}, {partInfo.SketchName} NOT PROPERLY DEFINED");
                            releaseResult = 1;

                        }
                    }
                    break;
                case swDocumentTypes_e.swDocASSEMBLY:

                    SolidworksServiceResult<ModelDoc2> openAssemblyResult = ThisSolidworksService.OpenAssembly(currentFileObj.FilePath);
                    if (!openAssemblyResult.Success)
                    {
                        reportLines.Add($"Error Opening {currentFileObj.FilePath}");
                        return 1;
                    }

                    ModelDoc2 activeAssy = openAssemblyResult.response;
                    AssemblyValidationResult assemblyReleaseResult = ThisReleaseValidationService.CheckAssembly(activeAssy);
                    if (assemblyReleaseResult.CriticalError)
                    {
                        reportLines.Add($"Validation Status: CRITICAL ERROR \n COULD NOT RUN VALIDATION");
                        releaseResult = 1;

                    }
                    if (assemblyReleaseResult.FoundComponentErrors)

                    {
                        reportLines.Add($"Validation Status: Failed");
                        reportLines.Add (currentFileObj.FilePath);
                        foreach (ComponentInfo componentError in assemblyReleaseResult.TotalComponentErrors)
                        {
                            reportLines.Add($"\t-{componentError.Name}({componentError.Configuration}) - {componentError.ConstraintStatus}");
                        }
                        releaseResult = 1;

                    }

                    break;
                case swDocumentTypes_e.swDocDRAWING:
                    bool deleteAnnotations = false;
                    docType = swDocumentTypes_e.swDocDRAWING;
                    SolidworksServiceResult<ModelDoc2> openDrawingResult = ThisSolidworksService.OpenDrawing(currentFileObj.FilePath);
                    if (!openDrawingResult.Success)
                    {
                        return 1;
                    }
                    ModelDoc2 activeDrawing = openDrawingResult.response;
                    DrawingValidationResult drawingReleaseResult = ThisReleaseValidationService.CheckDrawing(activeDrawing, currentFileObj.FilePath,currentFileObj);
                    if (drawingReleaseResult.CriticalError)
                    {
                        reportLines.Add($"Validation Status: CRITICAL ERROR \n COULD NOT RUN VALIDATION");
                        releaseResult = 1;
                    }
                    else if(drawingReleaseResult.FoundDanglingAnnotations)
                    {
                        reportLines.Add($"Validation Status: Failed");
                        reportLines.Add(currentFileObj.FilePath);
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
                    
                    if (drawingReleaseResult.FoundDanglingAnnotations && deleteAnnotations)
                    {
                        ThisSolidworksService.DeleteAllSelections(activeDrawing);
                        ThisSolidworksService.MoveSWFile(currentFileObj.FilePath, currentFileObj.FileName, ThisEcnRelease.ReleaseFolderSrc);


                    }

                    break;
                default:
                    break;
            }
            //ThisSolidworksService.CloseFile(file.FilePath);
            if (releaseResult != 0)
            {
                ThisReleaseReport.WriteToReportMultiline(reportLines);
                return 1;
            }
            return 0;
        }
        // this would be used for DFS traversal if used in future

        public bool CopyEcnFolder(string sourceFolder, string destFolder)
        {
            try
            {
                if (!Directory.Exists(sourceFolder))
                {
                    MessageBox.Show($"Source folder does not exist {sourceFolder}");
                    return false;
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
                string msg = $"Error copying folder: {ex.Message}";
                List<string> list = new List<string>();
                list.Add(msg);
                ThisReleaseReport.WriteToReportMultiline(list);
                return false;
            }
            return true;
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
                string msg = $"Error clearing temp folder: {ex.Message}";
                List<string> list = new List<string>();
                list.Add(msg);
                ThisReleaseReport.WriteToReportMultiline(list);
            }
        }
        private List<List<string>> ParseECNFile(EcnRelease thisRelease)
        {
            List<List<string>> EcnFileData = new List<List<string>>();
            
            string ecnTextFile = thisRelease.ReleaseTxtFile;
            try
            {
                int txtFileIndex = 0;

                // Read lines from the file one by one as you loop
                foreach (string line in File.ReadLines(ecnTextFile))
                {
                    //current format of txt file has empty first line
                    if (txtFileIndex == 0)
                    {
                        txtFileIndex++;
                        continue;
                    }

                    // Process each line here
                    Console.WriteLine(line);
                    string[] parsedLine = line.Split(new char[] { ' ' });
                    List<string> parsedData = new List<string>(parsedLine);
                    EcnFileData.Add(parsedData);
                    txtFileIndex++;
                }
            }
            catch (IOException e)
            {
                Console.WriteLine($"An I/O error occurred: {e.Message}");
            }
            catch (Exception e)
            {
                Console.WriteLine($"An unexpected error occurred: {e.Message}");
            }

            return EcnFileData;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}

