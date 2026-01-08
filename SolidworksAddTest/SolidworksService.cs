using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolidworksAddTest
{
    public class SolidworksServiceResult
    {

    }
    public class SolidworksServiceResult<T>
    { 
        public bool Success { get; set; }
        public T response { get; set; }
        public string ErrorMessage {  get; set; }


        public SolidworksServiceResult() 
        {
            Success = false;

            

        }
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

            SubFeatureTypeExceptions = new HashSet<string> {
            "FlatPattern",
            "RefPlane",
            "RefAxis"};

            FeatureTypeExceptions = new HashSet<string> {
            "ProfileFeature",
            "MirrorStock"};

        }
        public SolidworksServiceResult<ModelDoc2> OpenAssembly(string filepath)
        {
            SolidworksServiceResult<ModelDoc2> openAssemblyResult = new SolidworksServiceResult<ModelDoc2>();
            SolidworksServiceResult<ModelDoc2> openFileResult = new SolidworksServiceResult<ModelDoc2>();

            try
            {
                int options = (int)swOpenDocOptions_e.swOpenDocOptions_Silent | (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel;

                if (SolidWorksApp == null)
                {
                    openAssemblyResult.ErrorMessage = "Parent add-in is not set.";
                    return openAssemblyResult;

                }


                // Check if file exists
                if (!System.IO.File.Exists(filepath))
                {
                    openAssemblyResult.ErrorMessage = $"File not Found: {filepath}";
                    return openAssemblyResult;
                }

                // Define document type and options
                int docType = (int)swDocumentTypes_e.swDocASSEMBLY;

                // Open the document

                openFileResult = OpenFile(filepath, docType, options);
                if (openFileResult.Success)
                {
                    ModelDoc2 doc = openFileResult.response;
                    openAssemblyResult.Success = true;
                    ModelDocExtension swDocExt = doc.Extension;
                    openAssemblyResult.response = doc;

                }

                return openAssemblyResult;

            }
            catch (Exception ex)
            {
                string errorMsg = ($"Error opening assembly: {ex.Message}");
                openAssemblyResult.ErrorMessage = errorMsg;
                return openAssemblyResult;
            }
        }
        public SolidworksServiceResult<ModelDoc2> OpenFile(string filepath, int docType, int options)
        {
            SolidworksServiceResult<ModelDoc2> OpenFileResult = new SolidworksServiceResult<ModelDoc2>();
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
                OpenFileResult.response = doc;
                OpenFileResult.Success = true;
                // Optional: Get about the assembly
                string title = doc.GetTitle();
                string pathName = doc.GetPathName();

            }
            else
            {

                string SWErrorMsg = GetOpenDocumentError(errors);
                string errorMessage = ($"Failed to open document.\nError: {SWErrorMsg}\nWarnings: {warnings}");
                OpenFileResult.ErrorMessage = errorMessage;
            }
            return OpenFileResult;
        }
        public SolidworksServiceResult<bool> ApplySearchPaths(List<string> searchPathPriority)
        {
            SolidworksServiceResult<bool> applySearchPathResult = new SolidworksServiceResult<bool>();
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
            applySearchPathResult.Success = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error managing search paths class: {ex.Message}");
            }
            applySearchPathResult.response = applySearchPathResult.Success;
            return applySearchPathResult;

        }
        public SolidworksServiceResult<bool> CloseAllDocuments()
        {
            SolidworksServiceResult<bool> solidworksServiceResult = new SolidworksServiceResult<bool>();
            solidworksServiceResult.Success = SolidWorksApp.CloseAllDocuments(true);
            solidworksServiceResult.response = solidworksServiceResult.Success;
            return solidworksServiceResult;

        }
        public SolidworksServiceResult<bool> CloseFile(string filepath)
        {
            SolidworksServiceResult<bool> closeFileResult = new SolidworksServiceResult<bool>();

            try
            {
                if (SolidWorksApp == null)
                {
                    MessageBox.Show("Parent add-in is not set.");
                    closeFileResult.ErrorMessage = "Addin is not set correctly";
                    closeFileResult.response = closeFileResult.Success;
                }

                SolidWorksApp.CloseDoc(filepath);
                closeFileResult.Success = true;
                closeFileResult.response = closeFileResult.Success;
            }
            catch (Exception ex)
            {
                closeFileResult.ErrorMessage = $"Error Closing SW File: {ex.Message}";
                closeFileResult.response= closeFileResult.Success;
            }
            return closeFileResult;
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
        public SolidworksServiceResult<ModelDoc2> OpenPart(string filepath)
        {
            SolidworksServiceResult<ModelDoc2> openPartResult = new SolidworksServiceResult<ModelDoc2>();
            SolidworksServiceResult<ModelDoc2> openFileResult = new SolidworksServiceResult<ModelDoc2>();
            try
            {
                if (SolidWorksApp == null)
                {
                    string errorMessage = "Parent add-in is not set.";
                    openPartResult.ErrorMessage = errorMessage;

                     return openPartResult;
                }


                // Check if file exists
                if (!System.IO.File.Exists(filepath))
                {
                    string errorMessage = $"File not found: {filepath}";
                    openPartResult.ErrorMessage = errorMessage;
                    return openPartResult;
                }

                // Define document type and options

                int options = (int)swOpenDocOptions_e.swOpenDocOptions_Silent;
                int configuration = 0;
                string configName = "";
                int errors = 0;
                int warnings = 0;

                // Open the document
                
                openFileResult = OpenFile(filepath,(int)swDocumentTypes_e.swDocPART, options);


                if (openFileResult.Success)
                {
                    openPartResult.Success = true;
                    openPartResult.response = openFileResult.response;
       
                }
                else
                {
                    string SWerrorMsg = GetOpenDocumentError(errors);
                    string ErrorMsg = $"Failed to open document.\nError: {SWerrorMsg}\nWarnings: {warnings}";
                }
                return openPartResult;

            }
            catch (Exception ex)
            {

                string ErrorMsg = $"Error opening assembly: {ex.Message}";
                openPartResult.ErrorMessage = ErrorMsg;
                return openPartResult;
            }
        }
        public SolidworksServiceResult<ModelDoc2> OpenDrawing(string filepath)
        {
            SolidworksServiceResult<ModelDoc2> OpenDrawingResult = new SolidworksServiceResult<ModelDoc2>();
            SolidworksServiceResult<ModelDoc2> openFileResult = new SolidworksServiceResult<ModelDoc2>();
            try
            {

                if (SolidWorksApp == null)
                {
                    string errorMessage = "Parent add-in is not set.";
                    OpenDrawingResult.ErrorMessage = errorMessage;
                    return OpenDrawingResult;
                }

                // Check if file exists
                if (!System.IO.File.Exists(filepath))
                {
                    string errorMessage = "File not found: {filepath}";
                    OpenDrawingResult.ErrorMessage = errorMessage;
                    return OpenDrawingResult;
                }
                SolidWorksApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swAutomaticDrawingViewUpdate, false);
                // Define document type and options

                int options = (int)swOpenDocOptions_e.swOpenDocOptions_Silent;


                // Open the document
                openFileResult = OpenFile(filepath, (int)swDocumentTypes_e.swDocDRAWING, options);
                if (!openFileResult.Success)
                {
                    
                    openFileResult.ErrorMessage = $"Failed to open document {openFileResult.ErrorMessage}";
                    return openFileResult;
                }
                OpenDrawingResult.Success = true;
                OpenDrawingResult.response = openFileResult.response;
                return OpenDrawingResult;
            }
            catch (Exception ex)
            {
                OpenDrawingResult.ErrorMessage = ($"Error while opening Drawing: {ex.Message}");
                return OpenDrawingResult;
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
            if (doc == null)
            {
                MessageBox.Show("DOC is NULL FOr Configs");
            }
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
