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


namespace SolidworksAddTest
{
    [ComVisible(true)]
    [Guid("4099c769-bd7d-49a6-ac97-1ec1e38ddcf9")]
    [ProgId("SolidworksAddTest.DependenciesResult")]
    public partial class DependenciesResult : UserControl
    {
        private SWTestRP parentAddin;
        public HashSet<string> baseDependents = new HashSet<string>();


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
            var testReleaseList = new List<string>();
  
            testReleaseList.Add("M:\\181\\1810203000.SLDDRW");

            for (int i = 0; i < testReleaseList.Count; i++)
            {
                string currentFile = testReleaseList[i];
                SetSearchPaths(currentFile);
                ReleaseFile(currentFile);
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
            ManageSearchPaths(searchPathPriority);
      

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
                    MessageBox.Show($"main {DocName} dependent: {currentDependent}");

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
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception opening assembly: {ex.Message}");
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

        // Helper method to interpret error codes
        private string GetOpenDocumentError(int errorCode)
        {
            swFileLoadError_e error = (swFileLoadError_e)errorCode;
            return error.ToString();
        }
        private void ManageSearchPaths(List<string> searchPathPriority)
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

