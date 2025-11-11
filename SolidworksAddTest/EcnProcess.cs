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


    public class EcnFile
    { 
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public int LoadedFilesRemaining { get; set; }
        public List<string> SearchPaths { get; set; }
        public HashSet<EcnFile> Parents { get; set; }
        public swDocumentTypes_e DocumentType { get; set; }
        public EcnFile()
        {
            SearchPaths = new List<string>();
            Parents = new HashSet<EcnFile>();
            LoadedFilesRemaining = 0;
        }
        public void InsertSearchPaths(List<string> searchPaths) { 
            SearchPaths = searchPaths;
        }
        public void InsertParent(EcnFile parent) 
        { 
            Parents.Add(parent);
        }

    
    }
    public class SolidworksService
    {
        SldWorks SolidWorksApp { get; set; }

        public SolidworksService(SldWorks solidworksApp) 
        {

            SolidWorksApp = solidworksApp;
        }
        
        public void OpenFile(string filename)
        {   

        }
        public void ApplySearchPaths(List<string> searchPathPriority)
        {
            try
            {

                SolidWorksApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swUseFolderSearchRules, true);

                SolidWorksApp.SetSearchFolders((int)swSearchFolderTypes_e.swDocumentType, null);
                string searchFolders = SolidWorksApp.GetSearchFolders((int)swUserPreferenceToggle_e.swUseFolderSearchRules);




                // Clear all search folders (optional - be careful!)
                // swApp.SetSearchFolders(null);

                // Set new search folders

                string foldersString = "";
                foreach (string folder in searchPathPriority)
                {
                    foldersString += folder + ";";
                }
                // Validate paths exist before adding
                var validPaths = searchPathPriority.Where(path => Directory.Exists(path)).ToArray();

                if (validPaths.Length > 0)
                {
                    SolidWorksApp.SetSearchFolders((int)swSearchFolderTypes_e.swDocumentType, foldersString);
                }
                string newSearchFolders = SolidWorksApp.GetSearchFolders((int)swUserPreferenceToggle_e.swUseFolderSearchRules);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error managing search paths class: {ex.Message}");
            }
        }
        public void CloseAllDocuments()
        {
            SolidWorksApp.CloseAllDocuments(true);
        }
    }
    public class EcnRelease
    {
        public string ReleaseNumber { get; set; }
        public Dictionary<string, EcnFile> Files { get; set; }
        public HashSet<EcnFile> LeafFiles { get; set; }
        public HashSet<EcnFile> CompletedFiles { get; set; } 
        public Queue<EcnFile> ProcessingFileQueue { get; set; }
        public Stack<EcnFile> OpenFilesStack {  get; set; }
        public EcnRelease(string releaseNumber, int releaseType)
        {
            ReleaseNumber = releaseNumber;
            Files = new Dictionary<string, EcnFile>();
            LeafFiles = new HashSet<EcnFile>();
            CompletedFiles = new HashSet<EcnFile>(); 
            ProcessingFileQueue = new Queue<EcnFile>();
            OpenFilesStack = new Stack<EcnFile>();
        }
   
        public void AddFile(EcnFile file, string fileName)
        {
            Files[fileName] = file;
        }
        public void AddLeafFile(EcnFile file)
        {
            LeafFiles.Add(file);
        }
        public void RemoveLeafFile(EcnFile file)
        {
            LeafFiles.Remove(file);
        }
        public void AddCompletedFile(EcnFile file)
        {
            CompletedFiles.Add(file);
        }
        public void ProcessFilesPush(EcnFile file)
        {
            ProcessingFileQueue.Enqueue(file);
        }
        public EcnFile ProcessFilesPop()
        {
            return ProcessingFileQueue.Dequeue();
        }
        public void PushOpenFileStack(EcnFile file)
        {
            OpenFilesStack.Push(file);
        }
        public EcnFile PopOpenFileStack()
        {
           return OpenFilesStack.Pop();
        }

    }
    public class ReleaseReport
    {
        public string EcnNumber { get; set; }
        public string ReleaseType { get; set; } 
        private string reportFilePath { get; set; }
        public DateTime runTime {  get; set; }
        public DateTime startTime { get; set; }   
        public Dictionary<EcnFile, List<string>> Files { get; set; }
        public ReleaseReport(string ecnNumber, int releaseType)
        {
            reportFilePath = @"C:\Users\zacv\Documents\releaseTest";
            EcnNumber = ecnNumber;
            Files = new Dictionary<EcnFile, List<string>>();
            if (releaseType == 0)
            {
                ReleaseType = "RELEASE";
            }
            else 
            {
                ReleaseType = "READINESS FOR RELEASE";
            }
            string reportFolder = @"C:\Users\zacv\Documents\releaseTest";
            if (!System.IO.Directory.Exists(reportFolder))
            {
                System.IO.Directory.CreateDirectory(reportFolder);
            }
            startTime = DateTime.Now;
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            reportFilePath = System.IO.Path.Combine(reportFolder, $"{EcnNumber}_{timestamp}.txt");

            // Create initial report file
            CreateReportFile();

        }
        private void CreateReportFile()
        {
            try
            {
                using (System.IO.StreamWriter writer = new System.IO.StreamWriter(reportFilePath))
                {
                    writer.WriteLine("=".PadRight(60, '='));
                    writer.WriteLine($"ECN {ReleaseType}: {EcnNumber}");
                    writer.WriteLine("=".PadRight(60, '='));
                    writer.WriteLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    writer.WriteLine();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating report: {ex.Message}");
            }
        }
        public void FinishReport() 
        {
            int DateTimeRunTime = (int)(DateTime.Now - startTime).TotalMilliseconds;

            string runtimeString = $"Total Runtime: {DateTimeRunTime/1000}.{DateTimeRunTime%1000} S";
          

            List<string> FinalRuntime = new List<string>();
            FinalRuntime.Add( "" );
            FinalRuntime.Add("");
            FinalRuntime.Add(runtimeString);
            WriteToReport(FinalRuntime);
        }

        public void AddFile(EcnFile file)
        {
            Files[file] = new List<string>();
        }
        public void WriteToReport(List<string> lines)
        {
            try
            {
                System.IO.File.AppendAllLines(reportFilePath, lines);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error writing to report: {ex.Message}");
            }
        }
        public void OpenReport()
        {
            try
            {
                if (System.IO.File.Exists(reportFilePath))
                {
                    System.Diagnostics.Process.Start("notepad.exe", reportFilePath);
                }
                else
                {
                    MessageBox.Show("Report file not found");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening report: {ex.Message}");
            }
        }
    }

        public partial class EcnProcessService : UserControl
        {
            private SWTestRP parentAddin;
        public SolidworksService ThisSolidworksService { get; set; }
        public EcnRelease ThisEcnRelease { get; set; }

        public ReleaseReport ThisReleaseReport { get; set; }



        public EcnProcessService()
            {
            InitializeComponent();
            this.BackColor = Color.White;
            }
        public void InitalizeRelease()
        {
            string testReleaseNumber = "50001";
            int releaseMode = 0;
            ThisSolidworksService = new SolidworksService(parentAddin.SolidWorksApplication);
            ThisEcnRelease = new EcnRelease(testReleaseNumber, releaseMode);
            ThisReleaseReport = new ReleaseReport(testReleaseNumber, releaseMode);

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

            var testReleaseList = new List<string>();
            testReleaseList.Add("R:\\52576\\9092551121.SLDDRW");

            //testReleaseList.Add("M:\\181\\1810203000.SLDDRW");
            //testReleaseList.Add("M:\\181\\1810203000.SLDASM");
            testReleaseList.Add("C:\\Users\\zacv\\Documents\\releaseTest\\1810203000.SLDASM");
            testReleaseList.Add("C:\\Users\\zacv\\Documents\\releaseTest\\1810203000.SLDDRW");
            //testReleaseList.Add("M:\\181\\1810214000.SLDDRW");
            //testReleaseList.Add("M:\\181\\1810214000.SLDASM");
            //testReleaseList.Add("M:\\181\\1810214200.SLDASM");
            //testReleaseList.Add("M:\\181\\1810214200.SLDDRW");
            //testReleaseList.Add("M:\\181\\1810214215.SLDPRT");
            //testReleaseList.Add("M:\\181\\1810214215.SLDDRW");
            //testReleaseList.Add("M:\\181\\1810905000.SLDASM");
            //testReleaseList.Add("M:\\181\\1810905000.SLDDRW");
            //testReleaseList.Add("M:\\181\\1810212047.SLDPRT");
            //testReleaseList.Add("M:\\181\\1810212047.SLDDRW");

            //testReleaseList.Add("M:\\196\\1960000000.SLDASM");
            //testReleaseList.Add("M:\\181\\1810200000.SLDDRW");
            //testReleaseList.Add("M:\\181\\1810200003.SLDDRW");

            //testReleaseList.Add("R:\\52552\\9024637660.SLDASM");
            //testReleaseList.Add("R:\\52552\\9024637660.SLDDRW");
            //testReleaseList.Add("R:\\52552\\9024637660.SLDASM");
            //var thisReleaseReport = new ReleaseReport(thisRelease.ReleaseNumber, 1);
            
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
                var activeObj = ThisEcnRelease.Files[fileName];
                SetSearchPaths1(ThisEcnRelease.Files[fileName].FilePath, ThisEcnRelease, activeObj);
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
                ThisEcnRelease.PushOpenFileStack(currentFile);
                ThisSolidworksService.ApplySearchPaths(currentFile.SearchPaths);
                //ApplySWSearchPaths(currentFile.SearchPaths);
                int releaseStatus = ReleaseFile(currentFile, ThisReleaseReport);
                if (releaseStatus != 0)
                {
                    return 1;
                }
                ThisEcnRelease.AddCompletedFile(currentFile);
                if (currentFile.Parents.Count < 1)
                {
                    CloseSWFile(currentFile.FilePath);
                    ThisEcnRelease.PopOpenFileStack();
                }
                
                foreach (EcnFile file in currentFile.Parents)
                {
                    if (file.DocumentType == swDocumentTypes_e.swDocDRAWING && file.LoadedFilesRemaining == 1)
                    {
                        ReleaseFile(file, ThisReleaseReport);
                        CloseSWFile(file.FilePath);
                        ThisEcnRelease.AddCompletedFile(file);
                        //thisRelease.ProcessFilesPush(file);

                    }
                    else 
                    {
                        ThisEcnRelease.ProcessFilesPush(file);
                    }
                    file.LoadedFilesRemaining--;

                }
                CloseSWFile(currentFile.FilePath);


            }



            /*
            
            foreach (EcnFile file in thisRelease.LeafFiles) 

            {
                FileTraversal(file, file.FilePath, thisRelease, thisReleaseReport);

            }
            */

            ThisSolidworksService.CloseAllDocuments();



            return 0;
        }
        private int SetSearchPaths(string filepath)
        {

            if (parentAddin == null)
            {
                MessageBox.Show("Parent add-in is not set.");
                return 0;
            }

            SldWorks swApp = parentAddin.SolidWorksApplication;
            HashSet<string> dependencies = new HashSet<string>();
            Dictionary<string, int> folderCount = new Dictionary<string, int>();
            var folderPriority = new List<(float percent, string folderPath)>();
            GET_DEPENDENCIES(filepath, swApp, dependencies, folderCount);
            int dependenciesCount = dependencies.Count;
            foreach (KeyValuePair<string, int> path in folderCount)
            {   
                folderPriority.Add(((float)path.Value/dependenciesCount, path.Key));
                
            }
            folderPriority.Sort();
            folderPriority.Reverse();
            List<string> searchPathPriority = new List<string>();
            for (int i = 0; i < folderPriority.Count; i++)
            {
                searchPathPriority.Add(folderPriority[i].folderPath);

            }

            string[] archiveFolders = Directory.GetDirectories(@"M:/");
            foreach (string archiveFolder in archiveFolders)
            {
                searchPathPriority.Add(archiveFolder);
            }
            ApplySWSearchPaths(searchPathPriority);
      

            return dependenciesCount;
        }
        private int SetSearchPaths1(string filepath, EcnRelease thisRelease, EcnFile currentFile)
        {
            string testFolder = "C:\\Users\\zacv\\Documents\\releaseTest";
            if (parentAddin == null)
            {
                MessageBox.Show("Parent add-in is not set.");
                return 0;
            }

            SldWorks swApp = parentAddin.SolidWorksApplication;
            HashSet<string> dependencies = new HashSet<string>();
            Dictionary<string, int> folderCount = new Dictionary<string, int>();
            var folderPriority = new List<(int count, string folderPath)>();
            GET_DEPENDENCIES1(filepath, swApp, dependencies, folderCount ,thisRelease, currentFile);
            int dependenciesCount = dependencies.Count;
            foreach (KeyValuePair<string, int> path in folderCount)
            {
                folderPriority.Add((path.Value, path.Key));

            }
            folderPriority.Sort();
            folderPriority.Reverse();
            List<string> searchPathPriority = new List<string>();
            searchPathPriority.Add(testFolder);
            for (int i = 0; i < folderPriority.Count; i++)
            {
                searchPathPriority.Add(folderPriority[i].folderPath);

            }

            string[] archiveFolders = Directory.GetDirectories(@"M:/");
            foreach (string archiveFolder in archiveFolders)
            {
                searchPathPriority.Add(archiveFolder);
            }

            currentFile.InsertSearchPaths(searchPathPriority);


            return dependenciesCount;
        }
        private void GET_DEPENDENCIES(string DocName, SldWorks swApp, HashSet<string> dependencies, Dictionary<string, int> folderCount)
        {
            try
            {
                string[] DepList = swApp.GetDocumentDependencies2(DocName, false, false, false);

                if (DepList == null || DepList.Length == 0)
                {
                    return;
                }
                for (int i = 0; i < DepList.Length; i += 2)
                {
                    string currentDependent = DepList[i+1];

                    if (!dependencies.Contains(currentDependent))
                    {
                        ParsePath(currentDependent, folderCount);
                        dependencies.Add(currentDependent);
                        GET_DEPENDENCIES(currentDependent, swApp, dependencies, folderCount);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error Generating Dependencies: {ex.Message}");
            }
        return;
        }
        private void GET_DEPENDENCIES1(string DocName, SldWorks swApp, HashSet<string> dependencies, Dictionary<string, int> folderCount, EcnRelease thisRelease, EcnFile currentFile)
        {
            try
            {
                string[] DepList = swApp.GetDocumentDependencies2(DocName, false, false, false);

                if (DepList == null || DepList.Length == 0)
                {
                        thisRelease.AddLeafFile(currentFile);
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
                        if (thisRelease.Files.ContainsKey(currentFileDependent))
                        {
                            var dependentFileObj = thisRelease.Files[currentFileDependent];
                            dependentFileObj.InsertParent(currentFile);
                            if (thisRelease.LeafFiles.Contains(currentFile))
                            {
                                thisRelease.RemoveLeafFile(currentFile);
                            }
                            GET_DEPENDENCIES1(currentDependent, swApp, dependencies, folderCount, thisRelease, dependentFileObj);
                        }
                        else 
                        {
                            GET_DEPENDENCIES1(currentDependent, swApp, dependencies, folderCount, thisRelease, currentFile);
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
        private ModelDoc2 OpenAssembly(string filepath)
        {
            try
            {
                if (parentAddin == null)
                {
                    MessageBox.Show("Parent add-in is not set.");
                    return null;
                }

                SldWorks swApp = parentAddin.SolidWorksApplication;

                // Check if file exists
                if (!System.IO.File.Exists(filepath))
                {
                    MessageBox.Show($"File not found: {filepath}");
                    return null;
                }
                
                // Define document type and options
                int docType = (int)swDocumentTypes_e.swDocASSEMBLY;
                int options = (int)swOpenDocOptions_e.swOpenDocOptions_Silent;
                int configuration = 0;
                string configName = "";
                int errors = 0;
                int warnings = 0;

                // Open the document
                ModelDoc2 doc = swApp.OpenDoc6(
                    filepath,
                    docType,
                    options,
                    configName,
                    ref errors,
                    ref warnings
                );
                ModelDocExtension swDocExt = doc.Extension;

                if (doc != null)
                {

                    // Optional: Get some basic info about the assembly
                    string title = doc.GetTitle();
                    string pathName = doc.GetPathName();

                }
                else
                {
                    string errorMsg = GetOpenDocumentError(errors);
                    MessageBox.Show($"Failed to open document.\nError: {errorMsg}\nWarnings: {warnings}");
                }
                return doc;

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception opening assembly: {ex.Message}");
                return null;
            }
        }
        private HashSet<string> getSuppressedMates(ModelDoc2 doc)
        {
            AssemblyDoc currentAssembly = (AssemblyDoc)doc;
            object[] Mates = null;

            object[] components = currentAssembly.GetComponents(true);

            HashSet<string> suppressedMatesResult = new HashSet<string>();
            foreach (object component in components)
            {
                Component2 swComponent = (Component2)component;
                if (swComponent.IsSuppressed()) continue;
                Mates = (Object[])swComponent.GetMates();
                if (Mates == null) continue;
                foreach (object mate in Mates)
                {
                    Feature mateFeat = (Feature)mate;
                    suppressedMatesResult.Add(mateFeat.Name);
                    
                }

            }
            return suppressedMatesResult;
        }
        private int CheckAssembly(ModelDoc2 doc, ReleaseReport releaseReport, string filepath)
        {
            AssemblyDoc currentAssembly = (AssemblyDoc)doc;
            Feature currentFeature = doc.FirstFeature();
            object[] components = currentAssembly.GetComponents(true);
            object[] Mates = null;
            List<string> mateSpace = new List<string>();
            List<string> assyName = new List<string>();
            assyName.Add($"File: {filepath}");
            releaseReport.WriteToReport(assyName);
            int validRelease = 0;
            int componenetError = 0;
            int mateError = 0;
            string previousMate = null;
            string currentMateError = null;
            List<string> compResult = new List<string>();
            HashSet<string> suppressedMatesSet = new HashSet<string>();
            suppressedMatesSet = getSuppressedMates(doc);

            foreach (object component in components)

            { 
                componenetError = 0;
                mateError = 0;
                List<string> mateErrors = new List<string>();
                List<string> componentErrors = new List<string>();
                List<string> MateSuppress = new List<string>();
                Component2 swComponent = (Component2)component;
                string[] swComponentSplitName = swComponent.Name.Split('-');
                string formattedSwComponentName = "";
                for (int i = 0; i < swComponentSplitName.Length - 1; i++)
                { 
                    formattedSwComponentName+= swComponentSplitName[i];
                }
                formattedSwComponentName += '<' + swComponentSplitName[1] + '>';
                

                Mates = (Object[])swComponent.GetMates();
                int solveResult = swComponent.GetConstrainedStatus();
                string partMessage = $"Part: {swComponent.Name2} resolve: {solveResult}" ;
                compResult.Add(partMessage);
                bool isSWComponenetSupressed = false;
                if (swComponent.IsSuppressed())
                {
                    isSWComponenetSupressed = true;
                }
        
                if (swComponent.IsPatternInstance() || isSWComponenetSupressed)
                {
                    continue;
                }
                if (solveResult == (int)swConstrainedStatus_e.swUnderConstrained)
                {
                    componentErrors.Add($"  {formattedSwComponentName} UNDERDEFINED");
                    componenetError = 1;
                }
                else if (solveResult == (int)swConstrainedStatus_e.swOverConstrained)
                {
                    componentErrors.Add($"  {formattedSwComponentName} OVERDEFINED");
                    componenetError = 1;
                }
                else if (solveResult != (int)swConstrainedStatus_e.swFullyConstrained)
                {
                    componentErrors.Add($"  {formattedSwComponentName} NOT PROPERLY DEFINED");
                    componenetError = 1;
                }
            
                 if (Mates != null) 
                {
                    foreach (Object SingleMate in Mates)

                    {
                  
                        if (!(SingleMate is Mate2) || isSWComponenetSupressed)
                        {
                            continue;
                        }

                        Feature mateFeat = (Feature)SingleMate;
                        int errorCodes = mateFeat.GetErrorCode();
                        bool[] isSuppressed = mateFeat.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration,null); 
                        string mateName = mateFeat.Name;
                 
                        if (suppressedMatesSet.Contains(mateName))

                        {
                            continue;
                        }
                        foreach (bool supressed in isSuppressed)
                        {

                            if (supressed)
                            {
                                MateSuppress.Add($"Mate '{mateName}' is SUPPRESSED");
                            }
                            else
                            {
                                MateSuppress.Add($"Mate '{mateName}' is NOT suppressed");
                            }
                        }

                        if (errorCodes != 0 && componenetError==0)
                        {
                            componenetError = 1;
                            componentErrors.Add($"  {formattedSwComponentName} PROPERLY DEFINED BUT HAS MATE ERRORS ");
                            componentErrors.Add($"Mate: {mateFeat.Name} errorNUM: {componenetError} comp: {swComponent.Name2}");
                            continue;
                        }

                        previousMate = mateFeat.Name;
                    }


                 }
                releaseReport.WriteToReport(MateSuppress);
               
                if (componenetError != 0) 
                {
                    releaseReport.WriteToReport(componentErrors);
                    if (mateErrors.Count > 0)
                    {
                        //releaseReport.WriteToReport(mateSpace);
                        //releaseReport.WriteToReport(mateErrors);

                    }
                if(componenetError!=0)
                    {
                        validRelease = 1;
                    }
                }
            }

           
            return validRelease;
            }




        private void OpenPart(string filepath)
        {
            try
            {
                if (parentAddin == null)
                {
                    MessageBox.Show("Parent add-in is not set.");
                    return;
                }

                SldWorks swApp = parentAddin.SolidWorksApplication;

                // Check if file exists
                if (!System.IO.File.Exists(filepath))
                {
                    MessageBox.Show($"File not found: {filepath}");
                    return;
                }

                // Define document type and options
                
                int options = (int)swOpenDocOptions_e.swOpenDocOptions_Silent | (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel;
                int configuration = 0;
                string configName = "";
                int errors = 0;
                int warnings = 0;

                // Open the document
                ModelDoc2 doc = swApp.OpenDoc6(
                    filepath,
                    (int)swDocumentTypes_e.swDocPART,
                    options,
                    configName,
                    ref errors,
                    ref warnings
                );

                if (doc != null)
                {

                    // Optional: Get some basic info about the assembly
                    string title = doc.GetTitle();
                    string pathName = doc.GetPathName();
                }
                else
                {
                    string errorMsg = GetOpenDocumentError(errors);
                    MessageBox.Show($"Failed to open document.\nError: {errorMsg}\nWarnings: {warnings}");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception opening assembly: {ex.Message}");
            }
        }
        private ModelDoc2 OpenDrawing(string filepath)
        {
            try
            {
                if (parentAddin == null)
                {
                    MessageBox.Show("Parent add-in is not set.");
                    return null;
                }

                SldWorks swApp = parentAddin.SolidWorksApplication;
                // Check if file exists
                if (!System.IO.File.Exists(filepath))
                {
                    MessageBox.Show($"File not found: {filepath}");
                    return null;
                }
                swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swAutomaticDrawingViewUpdate,false);
                // Define document type and options

                int options = (int)swOpenDocOptions_e.swOpenDocOptions_Silent | (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel;
                int configuration = 0;
                string configName = "";
                int errors = 0;
                int warnings = 0;

                // Open the document
                ModelDoc2 doc = swApp.OpenDoc6(
                    filepath,
                    (int)swDocumentTypes_e.swDocDRAWING,
                    options,
                    configName,
                    ref errors,
                    ref warnings
                );

                if (doc != null)
                {

                    // Optional: Get some basic info about the assembly
                    string title = doc.GetTitle();
                    string pathName = doc.GetPathName();
                }
                else
                {
                    string errorMsg = GetOpenDocumentError(errors);
                    MessageBox.Show($"Failed to open document.\nError: {errorMsg}\nWarnings: {warnings}");
                    return doc;
                }
                return doc;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception opening assembly: {ex.Message}");
                return null;
            }
        }

        private int CheckDrawing(ModelDoc2 doc, ReleaseReport releaseReport, string filePath, bool deleteAnnotations=false)
        {
            SldWorks swApp = parentAddin.SolidWorksApplication;
            ModelDocExtension swModExt = default(ModelDocExtension);
            int danglingCount = 0;
            int DanglingColorRef = 16777215;
            bool annotationsSeleted = false;
            SelectionMgr swSelmgr;
            SelectData swSelData;
            swSelmgr = (SelectionMgr)doc.SelectionManager;
            swModExt = (ModelDocExtension)doc.Extension;
            //deleteAnnotations = true;
            int sheetIdx = 0;

            swSelData = swSelmgr.CreateSelectData();


            if (doc == null) 
            {
                return 1;
            }
            int totalAnnotations = 0;
            DrawingDoc swDrawingDoc = (DrawingDoc)doc;
            object[] sheets = swDrawingDoc.GetViews();
            string[] sheetNames = swDrawingDoc.GetSheetNames();

            if (sheets == null || sheets.Length == 0)
            {
                MessageBox.Show("No sheets found in drawing");
                return 1;
            }
             foreach (object[] sheetObj in sheets)
             {
                    if (sheetObj == null) continue;
                    string currentSheetName = sheetNames[sheetIdx];
                    foreach (object view in sheetObj)
                    {
                        if (view == null) continue;
                        SolidWorks.Interop.sldworks.View currentView = (SolidWorks.Interop.sldworks.View)view;
                        string currentViewName = currentView.Name;
                        int annotationsCount = currentView.GetAnnotationCount();
                        object[] viewAnnotations = currentView.GetAnnotations();
                        if (viewAnnotations == null) continue;
                        foreach (object annotation in viewAnnotations)
                        {
                            if (annotation == null) continue;
                            Annotation currentAnnotation = (Annotation)annotation;
                            if (!currentAnnotation.IsDangling()) continue;

                            bool validAnnotationCheck = DanglingValidation(currentAnnotation, swSelmgr, swSelData, 
                                currentSheetName, currentViewName);
                            
                           
                            
                        }

                    }
                    sheetIdx++;

             }
            if (deleteAnnotations)
            {
                doc.EditDelete();
            }

             return 0;
        }
        private bool DanglingValidation(Annotation currentAnnotation, SelectionMgr swSelmgr, SelectData swSelData, 
            string currentSheetName, string currentViewName)
        {
            bool validAnnotationCheck = true;
            switch ((int)currentAnnotation.GetType())
            {
                case (int)swAnnotationType_e.swNote:
                    Note currentNote = (Note)currentAnnotation.GetSpecificAnnotation();
                    if (currentNote.IsBomBalloon() || currentNote.IsStackedBalloon())
                    {
                        MessageBox.Show("Dangling BOM Balloon");
                        validAnnotationCheck = false;
                    }
                    break;
                case (int)swAnnotationType_e.swDisplayDimension:
                    validAnnotationCheck = false;
                    break;
                case (int)swAnnotationType_e.swDatumOrigin:
                    validAnnotationCheck = false;
                    break;
                case (int)swAnnotationType_e.swDatumTag:
                    validAnnotationCheck = false;
                    break;
                case (int)swAnnotationType_e.swDatumTargetSym:
                    validAnnotationCheck = false;
                    break;
                case (int)swAnnotationType_e.swGTol:
                    validAnnotationCheck = false;
                    break;
                case (int)swAnnotationType_e.swWeldSymbol:
                    validAnnotationCheck = false;
                    break;
                case (int)swAnnotationType_e.swSFSymbol:
                    validAnnotationCheck = false;
                    break;
                case (int)swAnnotationType_e.swDowelSym:
                    validAnnotationCheck = false;
                    break;
                case (int)swAnnotationType_e.swCenterMarkSym:
                    validAnnotationCheck = false;
                    break;
                case (int)swAnnotationType_e.swCenterLine:
                    validAnnotationCheck = false;
                    break;
                case (int)swAnnotationType_e.swLeader:
                    validAnnotationCheck = false;
                    break;
                case (int)swAnnotationType_e.swCustomSymbol:
                    validAnnotationCheck = false;
                    break;

                default:
                    validAnnotationCheck = true;
                    break;
            }
            if (!validAnnotationCheck)
            {
                currentAnnotation.Select3(true, swSelData);
                MessageBox.Show($"-{currentSheetName}, {currentViewName} - {currentAnnotation.GetName()} is dangling");
            }

            return validAnnotationCheck;
        }
        private void CloseSWFile(string filepath)
        {
            try
            {
                if (parentAddin == null)
                {
                    MessageBox.Show("Parent add-in is not set.");
                    return;
                }

                SldWorks swApp = parentAddin.SolidWorksApplication;
                swApp.CloseDoc(filepath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error Closing SW File: {ex.Message}");
            }
        }

        private int ReleaseFile(EcnFile file, ReleaseReport releaseReport)
        {
            swDocumentTypes_e docType= file.DocumentType;
            int releaseResult = 0;
            switch (docType)
            {
                case swDocumentTypes_e.swDocPART:
                    OpenPart(file.FilePath);
                    break;
                case swDocumentTypes_e.swDocASSEMBLY:
                    ModelDoc2 activeAssy = OpenAssembly(file.FilePath);
                    releaseResult = CheckAssembly(activeAssy, releaseReport, file.FilePath);
                    break;
                case swDocumentTypes_e.swDocDRAWING:
                    docType = swDocumentTypes_e.swDocDRAWING;
                    ModelDoc2 activeSWDrawing = OpenDrawing(file.FilePath);
                    releaseResult = CheckDrawing(activeSWDrawing, releaseReport, file.FilePath);
                    break;
                default:
                    break;
            }
            if (releaseResult != 0)
            {
                CloseSWFile(file.FilePath);
                releaseReport.FinishReport();
                releaseReport.OpenReport();
                return 1;
            }
            return 0;
        }
        private void FileTraversal(EcnFile currentFile, string filePath, EcnRelease thisRelease, ReleaseReport thisReport)
        {
            bool canClose = true;
            if (thisRelease.CompletedFiles.Contains(currentFile))
            {
                return;
            }
            if (currentFile.LoadedFilesRemaining > 0)
            {
                return;
            }
            thisRelease.AddCompletedFile(currentFile);
            ApplySWSearchPaths(currentFile.SearchPaths);
            ReleaseFile(currentFile, thisReport);
            
            foreach (EcnFile parentFile in currentFile.Parents)
            {
                parentFile.LoadedFilesRemaining--;
                if (parentFile.LoadedFilesRemaining > 0) { 
                    canClose = false;
                    continue;
                }
                FileTraversal(parentFile, parentFile.FilePath, thisRelease, thisReport);
                
                
                
            }
            /*
            if (canClose)
            {
                CloseSWFile(filePath);


            }
            */
            CloseSWFile(filePath);

        }


        // Helper method to interpret error codes
        private string GetOpenDocumentError(int errorCode)
        {
            swFileLoadError_e error = (swFileLoadError_e)errorCode;
            return error.ToString();
        }
        private void ApplySWSearchPaths(List<string> searchPathPriority)
        {
            try
            {
                SldWorks swApp = parentAddin.SolidWorksApplication;

                // Get all current search folders
                swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swUseFolderSearchRules, true);

                swApp.SetSearchFolders((int)swSearchFolderTypes_e.swDocumentType, null);
                string searchFolders = swApp.GetSearchFolders((int)swUserPreferenceToggle_e.swUseFolderSearchRules);




                // Clear all search folders (optional - be careful!)
                // swApp.SetSearchFolders(null);

                // Set new search folders

                string foldersString = "";
                foreach (string folder in searchPathPriority)
                {
                    foldersString += folder+";";
                }
                // Validate paths exist before adding
                var validPaths = searchPathPriority.Where(path => Directory.Exists(path)).ToArray();

                if (validPaths.Length > 0)
                {
                    swApp.SetSearchFolders((int)swSearchFolderTypes_e.swDocumentType, foldersString);
                }
                string newSearchFolders = swApp.GetSearchFolders((int)swUserPreferenceToggle_e.swUseFolderSearchRules);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error managing search paths: {ex.Message}");
            }
        }
      

    }
}

