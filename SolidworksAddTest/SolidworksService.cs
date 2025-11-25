using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolidworksAddTest
{
    public class SolidworksServiceResponse
    {

    }

    public class SolidworksService
    {
        SldWorks SolidWorksApp { get; set; }

        ConfigurationManager ConfigManager { get; set; }
        public Dictionary<int, string> ReleaseAnnotations {  get; set; }
        public const string PARTFILEEXT = ".SLDPRT";
        public const string ASSEMBLYFILEEXT = ".SLDASM";
        public const string DRAWINGFILEEXT = ".SLDDRW";

        public HashSet<string> SubFeatureTypeExceptions { get; set; }
        public HashSet<string> FeatureTypeExceptions { get; set; }

        public SolidworksService(SldWorks solidworksApp)
        {

            SolidWorksApp = solidworksApp;

            ReleaseAnnotations = new Dictionary<int, string>();
            ReleaseAnnotations[(int)swAnnotationType_e.swNote] = "Balloon";
            ReleaseAnnotations[(int)swAnnotationType_e.swDisplayDimension] = "Dimension";
            ReleaseAnnotations[(int)swAnnotationType_e.swDatumOrigin] = "Datum Origin";
            ReleaseAnnotations[(int)swAnnotationType_e.swDatumTag] = "Datum Tag";
            ReleaseAnnotations[(int)swAnnotationType_e.swDatumTargetSym] = "Datum Target";
            ReleaseAnnotations[(int)swAnnotationType_e.swGTol] = "Geo Tol";
            ReleaseAnnotations[(int)swAnnotationType_e.swWeldSymbol] = "Weld Symbol";
            ReleaseAnnotations[(int)swAnnotationType_e.swSFSymbol] = "Surface Finish Symbol";
            ReleaseAnnotations[(int)swAnnotationType_e.swDowelSym] = "Dowel Pin Symbol";
            ReleaseAnnotations[(int)swAnnotationType_e.swCenterMarkSym] = "Center Mark";
            ReleaseAnnotations[(int)swAnnotationType_e.swCenterLine] = "CetnerLine";
            ReleaseAnnotations[(int)swAnnotationType_e.swLeader] = "General Leader";
            ReleaseAnnotations[(int)swAnnotationType_e.swCustomSymbol] = "Custom Symbol";

            SubFeatureTypeExceptions = new HashSet<string>();

            SubFeatureTypeExceptions.Add("FlatPattern");
            SubFeatureTypeExceptions.Add("RefPlane");
            SubFeatureTypeExceptions.Add("RefAxis");



            FeatureTypeExceptions = new HashSet<string>();
            FeatureTypeExceptions.Add("ProfileFeature");
            FeatureTypeExceptions.Add("MirrorStock");
        }
        public ModelDoc2 OpenAssembly(string filepath)
        {
            try
            {
                int options = (int)swOpenDocOptions_e.swOpenDocOptions_Silent | (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel;

                if (SolidWorksApp == null)
                {
                    MessageBox.Show("Parent add-in is not set.");
                    return null;
                }


                // Check if file exists
                if (!System.IO.File.Exists(filepath))
                {
                    MessageBox.Show($"File not found: {filepath}");
                    return null;
                }

                // Define document type and options
                int docType = (int)swDocumentTypes_e.swDocASSEMBLY;

                // Open the document
                ModelDoc2 doc = OpenFile(filepath, docType, options);


                ModelDocExtension swDocExt = doc.Extension;

                return doc;

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception opening assembly: {ex.Message}");
                return null;
            }
        }
        public ModelDoc2 OpenFile(string filepath, int docType, int options)
        {
            int configuration = 0;
            string configName = "";
            int errors = 0;
            int warnings = 0;
            ModelDoc2 doc = SolidWorksApp.OpenDoc6(
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
            return doc;
        }
        public void ApplySearchPaths(List<string> searchPathPriority)
        {
            try
            {

                SolidWorksApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swUseFolderSearchRules, true);
                // CLEAR EXISTING SEARCH PATHS. 
                SolidWorksApp.SetSearchFolders((int)swSearchFolderTypes_e.swDocumentType, null);
                string searchFolders = SolidWorksApp.GetSearchFolders((int)swUserPreferenceToggle_e.swUseFolderSearchRules);

                //VALIDATE PATHS EXIST
                var validPaths = searchPathPriority.Where(path => Directory.Exists(path)).ToArray();

                string foldersString = "";
                foreach (string folder in validPaths)
                {
                    foldersString += folder + ";";
                }
                //APPLY NON EMPTY SEARCH PATHS

                if (validPaths.Length > 0)
                {
                    SolidWorksApp.SetSearchFolders((int)swSearchFolderTypes_e.swDocumentType, foldersString);
                }
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
        public void CloseFile(string filepath)
        {
            try
            {
                if (SolidWorksApp == null)
                {
                    MessageBox.Show("Parent add-in is not set.");
                    return;
                }

                SolidWorksApp.CloseDoc(filepath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error Closing SW File: {ex.Message}");
            }
        }
        public string[] GetDocumentDependencies(string DocName)
        {
            
           string[] DepList = SolidWorksApp.GetDocumentDependencies2(DocName, false, false, false);
            return DepList;
        }
        public HashSet<string> getSuppressedComponentMates(ModelDoc2 doc)
        {
            AssemblyDoc currentAssembly = (AssemblyDoc)doc;
            object[] Mates = null;
            HashSet<string> suppressedMatesResult = new HashSet<string>();
            object[] components = currentAssembly.GetComponents(true);
            if (components != null)
            {
                foreach (object component in components)
                {
                    Component2 swComponent = (Component2)component;
                    if (!swComponent.IsSuppressed()) continue;
                    Mates = (Object[])swComponent.GetMates();
                    if (Mates == null) continue;
                    foreach (object mate in Mates)
                    {
                        Feature mateFeat = (Feature)mate;
                        suppressedMatesResult.Add(mateFeat.Name);

                    }

                }
            }
            return suppressedMatesResult;
        }
        public ModelDoc2 OpenPart(string filepath)
        {
            try
            {
                if (SolidWorksApp == null)
                {
                    MessageBox.Show("Parent add-in is not set.");
                     return null;
                }


                // Check if file exists
                if (!System.IO.File.Exists(filepath))
                {
                    MessageBox.Show($"File not found: {filepath}");
                    return null;
                }

                // Define document type and options

                int options = (int)swOpenDocOptions_e.swOpenDocOptions_Silent;
                int configuration = 0;
                string configName = "";
                int errors = 0;
                int warnings = 0;

                // Open the document
                ModelDoc2 doc = OpenFile(filepath,(int)swDocumentTypes_e.swDocPART, options);

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
        public ModelDoc2 OpenDrawing(string filepath)
        {
            try
            {
                if (SolidWorksApp == null)
                {
                    MessageBox.Show("Parent add-in is not set.");
                    return null;
                }

                // Check if file exists
                if (!System.IO.File.Exists(filepath))
                {
                    MessageBox.Show($"File not found: {filepath}");
                    return null;
                }
                SolidWorksApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swAutomaticDrawingViewUpdate, false);
                // Define document type and options

                int options = (int)swOpenDocOptions_e.swOpenDocOptions_Silent  ;
                int configuration = 0;
                string configName = "";
                int errors = 0;
                int warnings = 0;

                // Open the document
                ModelDoc2 doc = SolidWorksApp.OpenDoc6(
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
        private string GetOpenDocumentError(int errorCode)
        {
            swFileLoadError_e error = (swFileLoadError_e)errorCode;
            return error.ToString();
        }
        public object[] GetDrawingSheets(DrawingDoc doc)
        {
            object[] sheets = doc.GetViews();
            return sheets;

        }
        public string[] GetDrawingSheetNames(DrawingDoc doc)
        {
            string[] SheetNames = doc.GetSheetNames();
            return SheetNames;
        }

        public int GetAnnotationCount(SolidWorks.Interop.sldworks.View view)
        {
            return view.GetAnnotationCount();
        }
        public object[] GetAnnotations(SolidWorks.Interop.sldworks.View view)
        {
            return view.GetAnnotations();
        }
        public bool isAnnotationDangling(Annotation annotation)
        {
            return annotation.IsDangling();
        }
        public void SelectAnnotation(Annotation annotation, SelectData swSelData)
        {

            annotation.Select3(true, swSelData);

        }
        public void DeleteAllSelections(ModelDoc2 doc)
        {
            doc.EditDelete();
            return;
        }
        public string GetAnnotationType(Annotation annotation)
        {
            int annotationType = (int)annotation.GetType();
            if (ReleaseAnnotations.ContainsKey(annotationType))
            {
                return ReleaseAnnotations[annotationType];
            }
            else
            {
                return "Unidentified Release Annotation Found";
            }
        }
        public string[] GetConfigurationNames(ModelDoc2 doc)
        {
            string[] configNames = doc.GetConfigurationNames();
            return configNames;
        }
        public bool SaveSWDocument(ModelDoc2 doc)
        {
            return doc.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0);

        }
        public int MoveSWFile(string srcPath, string srcFileName, string dstFolder)
        {
            try
            {
                string fullDstPath = Path.Combine(dstFolder, srcFileName);
                File.Copy(srcPath, fullDstPath, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error moving SW file: {ex.Message}");
                return 1;
            }
            return 0;
        }
    }
}
