using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;



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
        public EcnRelease(string releaseNumber)
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
            var thisRelease = new EcnRelease("50001");

            var testReleaseList = new List<string>();

            testReleaseList.Add("M:\\181\\1810203000.SLDDRW");
            testReleaseList.Add("M:\\181\\1810203000.SLDASM");
            testReleaseList.Add("M:\\181\\1810214000.SLDDRW");
            testReleaseList.Add("M:\\181\\1810214000.SLDASM");
            testReleaseList.Add("M:\\181\\1810214200.SLDASM");
            testReleaseList.Add("M:\\181\\1810214200.SLDDRW");
            testReleaseList.Add("M:\\181\\1810214215.SLDPRT");
            testReleaseList.Add("M:\\181\\1810214215.SLDDRW");
            testReleaseList.Add("M:\\181\\1810905000.SLDASM");
            testReleaseList.Add("M:\\181\\1810905000.SLDDRW");
            testReleaseList.Add("M:\\181\\1810212047.SLDPRT");
            testReleaseList.Add("M:\\181\\1810212047.SLDDRW");
            //testReleaseList.Add("M:\\196\\1960000000.SLDASM");
            //testReleaseList.Add("M:\\181\\1810200000.SLDDRW");
            //testReleaseList.Add("M:\\181\\1810200003.SLDDRW");
            SldWorks swApp = parentAddin.SolidWorksApplication;

            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swLargeAsmModeSuspendAutoRebuild,true);
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
                //ReleaseFile(currentFile);
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
                ReleaseFile(currentFile);
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
                        //ReleaseFile(file);
                        //CloseSWFile(file.FilePath);
                        //thisRelease.AddCompletedFile(file);
                        thisRelease.ProcessFilesPush(file);

                    }
                    else 
                    {
                        thisRelease.ProcessFilesPush(file);
                    }
                    file.LoadedFilesRemaining--;
                }
                //CloseSWFile(currentFile.FilePath);


            }
            swApp.CloseAllDocuments(true);
            
            
            
            /*
            foreach (EcnFile file in thisRelease.LeafFiles) 

            {
                FileTraversal(file, file.FilePath, thisRelease);

            }
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
        private void OpenAssembly(string filepath)
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

        private void ReleaseFile(EcnFile file)
        {
            swDocumentTypes_e docType= file.DocumentType;
            switch (docType)
            {
                case swDocumentTypes_e.swDocPART:
                    OpenPart(file.FilePath);
                    break;
                case swDocumentTypes_e.swDocASSEMBLY:
                    OpenAssembly(file.FilePath);
                    break;
                case swDocumentTypes_e.swDocDRAWING:
                    docType = swDocumentTypes_e.swDocDRAWING;
                    OpenDrawing(file.FilePath);
                    break;
                default:
                    break;


            }
        }
        private void FileTraversal(EcnFile currentFile, string filePath, EcnRelease thisRelease)
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
            ReleaseFile(currentFile);
            foreach (EcnFile parentFile in currentFile.Parents)
            {
                parentFile.LoadedFilesRemaining--;
                if (parentFile.LoadedFilesRemaining > 0) { 
                    canClose = false;
                    continue;
                }
                FileTraversal(parentFile, parentFile.FilePath, thisRelease);
                
                
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

