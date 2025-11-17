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
            string testReleaseNumber = "50001";
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

            MessageBox.Show($"Runtime: {runtime/1000}.{runtime%1000} s");
        }
        private int RunRelease() 
        {
            //var thisRelease = new EcnRelease("50001", 0);
            //modify ecn location after inital test
            InitalizeRelease();

            var testReleaseList = new List<string>();
            string folderPath = $@"C:\Users\zacv\Documents\releaseTest\{ThisEcnRelease.ReleaseNumber}";

            foreach (string file in Directory.GetFiles(folderPath))
            {
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
                    case ".SLDPRT":
                        currentFileObj.DocumentType = swDocumentTypes_e.swDocPART;
                        break;
                    case ".SLDASM":
                        currentFileObj.DocumentType = swDocumentTypes_e.swDocASSEMBLY;
                        break;
                    case ".SLDDRW":
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
                var currentFile = ThisEcnRelease.ProcessFilesPop();
                if (ThisEcnRelease.CompletedFiles.Contains(currentFile))
                {
                    continue;
                }
                if (currentFile.LoadedFilesRemaining > 0)
                {
                    continue;
                }
                ThisSolidworksService.ApplySearchPaths(currentFile.SearchPaths);
                //ApplySWSearchPaths(currentFile.SearchPaths);
                int releaseStatus = ReleaseFile(currentFile);
                if (releaseStatus != 0)
                {
                    return 1;
                }
                ThisEcnRelease.AddCompletedFile(currentFile);
                if (currentFile.Parents.Count < 1)
                {
                    ThisSolidworksService.CloseFile(currentFile.FilePath);
                }
                
                foreach (EcnFile ParentFile in currentFile.Parents)
                {
                    if (ParentFile.DocumentType == swDocumentTypes_e.swDocDRAWING && ParentFile.LoadedFilesRemaining == 1)
                    {
                        ReleaseFile(ParentFile);
                        ThisSolidworksService.CloseFile(ParentFile.FilePath);
                        ThisEcnRelease.AddCompletedFile(ParentFile);
                        //thisRelease.ProcessFilesPush(file);

                    }
                    else 
                    {
                        ThisEcnRelease.ProcessFilesPush(ParentFile);
                    }
                    ParentFile.LoadedFilesRemaining--;

                }
                ThisSolidworksService.CloseFile(currentFile.FilePath);


            }


            ThisSolidworksService.CloseAllDocuments();



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
            GetDependenciesAndLeafs(filepath, dependencies, folderCount , currentFile);
            int dependenciesCount = dependencies.Count;
            foreach (KeyValuePair<string, int> path in folderCount)
            {
                folderPriority.Add((path.Value, path.Key));

            }
            folderPriority.Sort();
            folderPriority.Reverse();
            List<string> searchPathPriority = new List<string>();
            searchPathPriority.Add(thisRelease.ReleaseFolder);
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
        private void GetDependenciesAndLeafs(string DocName, HashSet<string> dependencies, Dictionary<string, int> folderCount, EcnFile currentFile)
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
                            GetDependenciesAndLeafs(currentDependent, dependencies, folderCount, dependentFileObj);
                        }
                        else 
                        {
                            GetDependenciesAndLeafs(currentDependent, dependencies, folderCount, currentFile);
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
            int releaseResult = 0;
            swDocumentTypes_e docType = file.DocumentType;
            switch (docType)
            {
                case swDocumentTypes_e.swDocPART:
                    ModelDoc2 activePart = ThisSolidworksService.OpenPart(file.FilePath);
                    PartValidationResult partReleaseResult = ThisReleaseValidationService.CheckPart(activePart, file.FilePath);


                    break;
                case swDocumentTypes_e.swDocASSEMBLY:
                    ModelDoc2 activeAssy = ThisSolidworksService.OpenAssembly(file.FilePath);
                    releaseResult = ThisReleaseValidationService.CheckAssembly(activeAssy);

                    break;
                case swDocumentTypes_e.swDocDRAWING:
                    docType = swDocumentTypes_e.swDocDRAWING;
                    ModelDoc2 activeDrawing = ThisSolidworksService.OpenDrawing(file.FilePath);
                    DrawingValidationResult drawingReleaseResult = ThisReleaseValidationService.CheckDrawing(activeDrawing, file.FilePath);
                    if (drawingReleaseResult.CriticalError)
                    {
                        MessageBox.Show("Crit ERROR!");
                    }
                    else if(drawingReleaseResult.FoundDanglingAnnotations)
                    {
                        List<string> reportLines = new List<string>();
                        reportLines.Add($"{file.FilePath}\rValidation Status: Failed\n");
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

                        releaseResult = 0;
                        ThisReleaseReport.WriteToReport(reportLines);

                    }
                    break;
                default:
                    break;
            }
            if (releaseResult != 0)
            {
                ThisSolidworksService.CloseFile(file.FilePath);
                ThisReleaseReport.FinishReport();
                ThisReleaseReport.OpenReport();
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
            /*
            if (canClose)
            {
                CloseSWFile(filePath);


            }
            */
            ThisSolidworksService.CloseFile(filePath);

        }


        // Helper method to interpret error codes
       

    }
}

