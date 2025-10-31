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
using static System.Windows.Forms.LinkLabel;



namespace SolidworksAddTest
{
    [ComVisible(true)]
    [Guid("4099c769-bd7d-49a6-ac97-1ec1e38ddcf9")]
    [ProgId("SolidworksAddTest.DependenciesResult")]


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

        public partial class DependenciesResult : UserControl
        {
            private SWTestRP parentAddin;


            public DependenciesResult()
            {
                InitializeComponent();


                // Add a simple control to verify it's working
                this.BackColor = Color.White;



            }
        

        public void SetParentAddin(SWTestRP parent)
        {
            parentAddin = parent;
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
            var thisRelease = new EcnRelease("50001", 0);

            var testReleaseList = new List<string>();

            testReleaseList.Add("M:\\181\\1810203000.SLDDRW");
            //testReleaseList.Add("M:\\181\\1810203000.SLDASM");
            testReleaseList.Add("C:\\Users\\zacv\\Documents\\releaseTest\\1810203000.SLDASM");
           // testReleaseList.Add("M:\\181\\1810214000.SLDDRW");
           // testReleaseList.Add("M:\\181\\1810214000.SLDASM");
           //testReleaseList.Add("M:\\181\\1810214200.SLDASM");
           //testReleaseList.Add("M:\\181\\1810214200.SLDDRW");
           //testReleaseList.Add("M:\\181\\1810214215.SLDPRT");
           //testReleaseList.Add("M:\\181\\1810214215.SLDDRW");
           // testReleaseList.Add("M:\\181\\1810905000.SLDASM");
           //testReleaseList.Add("M:\\181\\1810905000.SLDDRW");
            //testReleaseList.Add("M:\\181\\1810212047.SLDPRT");
            //testReleaseList.Add("M:\\181\\1810212047.SLDDRW");
            //testReleaseList.Add("M:\\196\\1960000000.SLDASM");
            //testReleaseList.Add("M:\\181\\1810200000.SLDDRW");
            //testReleaseList.Add("M:\\181\\1810200003.SLDDRW");
            SldWorks swApp = parentAddin.SolidWorksApplication;

            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swLargeAsmModeSuspendAutoRebuild,true);
            var thisReleaseReport = new ReleaseReport(thisRelease.ReleaseNumber, 1);
            
            for (int i = 0; i < testReleaseList.Count; i++)
            {
                string currentFile = testReleaseList[i];
                var currentFileObj = new EcnFile();
                currentFileObj.FilePath = testReleaseList[i];
                currentFileObj.FileName = GetFileWithExt(currentFile);
                thisRelease.AddFile(currentFileObj, currentFileObj.FileName);
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
            foreach ( string fileName in thisRelease.Files.Keys)

            {
                var activeObj = thisRelease.Files[fileName];
                SetSearchPaths1(thisRelease.Files[fileName].FilePath, thisRelease, activeObj);
            }

            
            foreach (EcnFile file in thisRelease.Files.Values)

            {
                foreach (EcnFile parentFile in file.Parents)
                {
                    parentFile.LoadedFilesRemaining++;
                    
                    if (thisRelease.LeafFiles.Contains(parentFile))
                    {
                        thisRelease.RemoveLeafFile(parentFile);
                    }
                }
                
            }
            
            foreach (EcnFile file in thisRelease.LeafFiles)
            {
                thisRelease.ProcessFilesPush(file);

            }
            
            while (thisRelease.ProcessingFileQueue.Count > 0)
            {
                var currentFile = thisRelease.ProcessFilesPop();
                if (thisRelease.CompletedFiles.Contains(currentFile))
                {
                    continue;
                }
                if (currentFile.LoadedFilesRemaining > 0)
                {
                    continue;
                }
                thisRelease.PushOpenFileStack(currentFile);
                ApplySWSearchPaths(currentFile.SearchPaths);
                ReleaseFile(currentFile, thisReleaseReport);
                thisRelease.AddCompletedFile(currentFile);
                if (currentFile.Parents.Count < 1)
                {
                    CloseSWFile(currentFile.FilePath);
                    thisRelease.PopOpenFileStack();
                }
                
                foreach (EcnFile file in currentFile.Parents)
                {
                    if (file.DocumentType == swDocumentTypes_e.swDocDRAWING && file.LoadedFilesRemaining == 1)
                    {
                        ReleaseFile(file, thisReleaseReport);
                        CloseSWFile(file.FilePath);
                        thisRelease.AddCompletedFile(file);
                        thisRelease.ProcessFilesPush(file);

                    }
                    else 
                    {
                        thisRelease.ProcessFilesPush(file);
                    }
                    file.LoadedFilesRemaining--;
                }
                CloseSWFile(currentFile.FilePath);


            }
            swApp.CloseAllDocuments(true);
            
            
            /*
            
            foreach (EcnFile file in thisRelease.LeafFiles) 

            {
                FileTraversal(file, file.FilePath, thisRelease);

            }
            swApp.CloseAllDocuments(true);
            */


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
        private int CheckAssembly(ModelDoc2 doc, ReleaseReport releaseReport, string filepath)
        {
            AssemblyDoc currentAssembly = (AssemblyDoc)doc;
            Feature currentFeature = doc.FirstFeature();
            /*
            while (currentFeature != null)

            {
                object currentFeatureType = currentFeature.GetType();
                Feature currentSubFeature = currentFeature.GetFirstSubFeature();
                string currentFeatureName = currentFeature.Name;
                int currentFeatureError = currentFeature.GetErrorCode();
                if (currentFeatureError != 0)
                {
                    MessageBox.Show($"Feature: {currentFeatureName} Error Code: {currentFeatureError}");

                }
                while (currentSubFeature != null)
                {
                    string currentSubFeatureName = currentSubFeature.Name;
                    int currentSubFeatureError = currentSubFeature.GetErrorCode();
                    if (currentSubFeatureError != 0)
                    {
                        MessageBox.Show($"Subfeature Error! Feature: {currentFeatureName}, Error Code: {currentSubFeatureError}");

                    }
                    currentSubFeature = currentSubFeature.GetNextSubFeature();
                }
                currentFeature = currentFeature.GetNextFeature();

            }
            */
            object[] components = currentAssembly.GetComponents(true);
            object[] Mates = null;
            List<string> mateSpace = new List<string>();
            mateSpace.Add("     CHECK MATES BELOW");
            List<string> assyName = new List<string>();
            assyName.Add($"File: {filepath}");
            releaseReport.WriteToReport(assyName);
            int validRelease = 0;
            int componenetError = 0;
            int mateError = 0;
            string previousMate = null;
            string currentMateError = null;
            foreach (object component in components)

            { 
                componenetError = 0;
                mateError = 0;
                List<string> mateErrors = new List<string>();
                List<string> componentErrors = new List<string>();
                Component2 swComponent = (Component2)component;
                Mates = (Object[])swComponent.GetMates();
                int solveResult = swComponent.GetConstrainedStatus();
                if (Mates == null && swComponent.IsPatternInstance())
                {
                    continue;
                }
                if (solveResult == (int)swConstrainedStatus_e.swUnderConstrained)
                {
                    componentErrors.Add($"  {swComponent.Name} UNDERDEFINED");
                    componenetError = 1;
                }
                else if (solveResult == (int)swConstrainedStatus_e.swOverConstrained)
                {
                    componentErrors.Add($"  {swComponent.Name} OVERDEFINED");
                    componenetError = 1;
                }
                else if (solveResult != (int)swConstrainedStatus_e.swFullyConstrained)
                {
                    componentErrors.Add($"   {swComponent.Name} NOT PROPERLY DEFINED");
                    componenetError = 1;
                }
            
                 if (Mates == null) 
                {
                    continue;
                    
                }

                    foreach (Object SingleMate in Mates)

                    {
                        if (!(SingleMate is Mate2))
                        {
                            continue;
                        }


                        Feature mateFeat = (Feature)SingleMate;
                        int errorCodes = mateFeat.GetErrorCode();
                        if (errorCodes != 0)
                        {
                            mateErrors.Add($"       {mateFeat.Name}");
                        }

                        previousMate = mateFeat.Name;
                    }
    
                if (componenetError != 0) 
                {
                    releaseReport.WriteToReport(componentErrors);
                    if (mateErrors.Count > 0)
                    {
                        releaseReport.WriteToReport(mateSpace);
                        releaseReport.WriteToReport(mateErrors);

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
        private void OpenDrawing(string filepath)
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
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception opening assembly: {ex.Message}");
            }
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

        private void ReleaseFile(EcnFile file, ReleaseReport releaseReport)
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
                    if (releaseResult != 0)
                    {
                        releaseReport.FinishReport();
                        releaseReport.OpenReport();
                    }
                    
                    break;
                case swDocumentTypes_e.swDocDRAWING:
                    docType = swDocumentTypes_e.swDocDRAWING;
                    OpenDrawing(file.FilePath);
                    break;
                default:
                    break;
                  


            }
            if (releaseResult != 0)
            {
                releaseReport.OpenReport();
            }
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
            if (canClose)
            {
                CloseSWFile(filePath);


            }


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

