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
        public List<string> SearchPaths { get; set; }
        public HashSet<EcnFile> Parents { get; set; }
        public swDocumentTypes_e DocumentType { get; set; }
        public EcnFile()
        {
            SearchPaths = new List<string>();
            Parents = new HashSet<EcnFile>();
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
        public HashSet<EcnFile> BaseDependents { get; set; }
        public HashSet<EcnFile> CompletedFiles { get; set; } 

        public EcnRelease(string releaseNumber)
        {
            ReleaseNumber = releaseNumber;
            Files = new Dictionary<string, EcnFile>();
            BaseDependents = new HashSet<EcnFile>();
            CompletedFiles = new HashSet<EcnFile>();   
        }
        public void FileSetup(EcnFile currentFile, string currentFileName)
        {


        }
        public void AddFile(EcnFile file, string fileName)
        {
            Files[fileName] = file;
        }
        public void AddBaseDependent(EcnFile file)
        {
            BaseDependents.Add(file);
        }
        public void RemoveBaseDependent(EcnFile file)
        {
            BaseDependents.Remove(file);
        }
        public void AddCompletedFile(EcnFile file)
        {
            CompletedFiles.Add(file);
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
            for (int i = 0; i < testReleaseList.Count; i++)
            {
                string currentFile = testReleaseList[i];
                var currentFileObj = new EcnFile();
                currentFileObj.FilePath = testReleaseList[i];
                currentFileObj.FileName = GetFileWithExt(currentFile);
                thisRelease.AddFile(currentFileObj, currentFileObj.FileName);
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
                  
                    if (thisRelease.BaseDependents.Contains(parentFile))
                    {
                        thisRelease.RemoveBaseDependent(parentFile);
                    }
                }
                
            }

            foreach (EcnFile file in thisRelease.BaseDependents) 
            {
                FileTraversal(file, file.FilePath, thisRelease);
            }

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
                        thisRelease.AddBaseDependent(currentFile);
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
                            if (thisRelease.BaseDependents.Contains(currentFile))
                            {
                                thisRelease.RemoveBaseDependent(currentFile);
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
                
                int options = (int)swOpenDocOptions_e.swOpenDocOptions_Silent;
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

                // Define document type and options

                int options = (int)swOpenDocOptions_e.swOpenDocOptions_Silent;
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
                swApp.CloseDoc(filepath);
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

        private void ReleaseFile(string filepath)
        {
            string fileExt = Path.GetExtension(filepath);
            swDocumentTypes_e docType;

            switch (fileExt)
            {
                case ".SLDPRT":
                    docType = swDocumentTypes_e.swDocPART;
                    OpenPart(filepath);
                    break;
                case ".SLDASM":
                    docType = swDocumentTypes_e.swDocASSEMBLY;
                    OpenAssembly(filepath);
                    break;
                case ".SLDDRW":
                    docType = swDocumentTypes_e.swDocDRAWING;
                    OpenDrawing(filepath);
                    break;
                default:
                    break;


            }
        }
        private void FileTraversal(EcnFile currentFile, string filePath, EcnRelease thisRelease)
        {
            if (thisRelease.CompletedFiles.Contains(currentFile))
            {
                return;
            }
            thisRelease.AddCompletedFile(currentFile);
            ApplySWSearchPaths(currentFile.SearchPaths);
            ReleaseFile(currentFile.FilePath);
            Thread.Sleep(100);
            foreach (EcnFile parentFile in currentFile.Parents)
            {
                FileTraversal(parentFile, parentFile.FilePath, thisRelease);
            }
            Thread.Sleep(100);
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

