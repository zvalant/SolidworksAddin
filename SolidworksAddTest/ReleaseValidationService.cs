using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.ExceptionServices;
using SolidWorks.Interop.swdocumentmgr;
using System.Threading;

namespace SolidworksAddTest

#region ValidationResultClasses
{
    /// <summary>
    ///  Contains all classes involving the validation results and their internal subclasses
    /// </summary>
    public class ValidationResultBase
    {
        public bool CriticalError { get; set; }

        public ValidationResultBase()
        {
            CriticalError = false;
        }
    }
    public class EcnFileValidationResult : ValidationResultBase
    {
        public HashSet<string> FilesNotInECN { get; set; }
        public HashSet<string> ECNFileRevMismatch { get; set; }
        public bool FoundECNFileRevMismatch { get; set; }
        public HashSet<string> FailedApproval { get; set; }
        public bool FoundFailedApproval { get; set; }
        public HashSet<string> DuplicateEcnFile {  get; set; }
        public bool FoundDuplicateECNFile { get; set; }
        public HashSet<string> MissingDrawingFile { get; set; }
        public bool FoundMissingDrawingFile    { get; set; }
        public Dictionary<string, bool> ECNFilesRequiredFound { get; set; }




        public EcnFileValidationResult()
        {
            FilesNotInECN = new HashSet<string>();
            ECNFileRevMismatch = new HashSet<string>();
            FailedApproval = new HashSet<string>();
            DuplicateEcnFile = new HashSet<string>();
            MissingDrawingFile = new HashSet<string>();

            ECNFilesRequiredFound = new Dictionary<string, bool>();
            
            CriticalError = false;
            FoundECNFileRevMismatch = false;
            FoundFailedApproval = false;
            FoundDuplicateECNFile = false;
            FoundMissingDrawingFile = false;
        }

    }
    public class DrawingValidationResult : ValidationResultBase
    {
        public bool FoundDanglingAnnotations { get; set; }
        public List<SheetInfo> TotalSheets { get; set; }

        public DrawingValidationResult()
        {
            CriticalError = false;
            FoundDanglingAnnotations = false;
            TotalSheets = new List<SheetInfo>();
        }
    }
    public class PartValidationResult : ValidationResultBase
    {
        public bool FoundFeatureErrors { get; set; }
        public List<PartFeatureInfo> TotalSketchErrors { get; set; }
        public PartValidationResult()
        {
            FoundFeatureErrors = false;
            TotalSketchErrors = new List<PartFeatureInfo>();
        }
    }
    public class AssemblyValidationResult : ValidationResultBase
    {
        public bool FoundComponentErrors { get; set; }
        public List<ComponentInfo> TotalComponentErrors { get; set; }
        public HashSet<int> ValidMateErrorCodes { get; set; }
        public AssemblyValidationResult()
        {
            List<int> validErrorCodes = new List<int> { 0, 48 };
            FoundComponentErrors = false;
            TotalComponentErrors = new List<ComponentInfo>();
            ValidMateErrorCodes = new HashSet<int>();
            foreach (int code in validErrorCodes)
            {
                ValidMateErrorCodes.Add(code);
            }
        }
    }

    public class SearchAndDependenciesValidationResult : ValidationResultBase
    {
        public List<string> ValidPaths { get; set; }
        public List<string> ArchivePaths { get; set; }
        public Dictionary<string, HashSet<string>> InvalidDependencies;
        public bool FoundInvalidDependencies { get; set; }
        public List<string> SearchPaths { get; set; }

        public SearchAndDependenciesValidationResult(string releaseFolderTemp, string releaseFolderSrc)
        {
            ValidPaths = new List<string> { "S:\\Engineering\\Archive",
            "M:",
            releaseFolderSrc,
            releaseFolderTemp,
            };
            ArchivePaths = new List<string> {
                "S:\\Engineering\\Archive",
                "M:"
            };
            FoundInvalidDependencies = false;
            // First String in Tuple is parent file and second is dependent.  
            InvalidDependencies = new Dictionary<string, HashSet<string>>();
            SearchPaths = new List<string>();

        }

    }
    public class ComponentInfo
    {
        public string Name { get; set; }
        public string ConstraintStatus { get; set; }
        public string Configuration { get; set; }
        public bool FoundError { get; set; }
        public ComponentInfo(string componentName, string constraintStatus, string configuration)
        {
            Name = componentName;
            ConstraintStatus = constraintStatus;
            Configuration = configuration;
            FoundError = false;
        }
    }
    public class PartFeatureInfo
    {
        public string FeatureName { get; set; }
        public string SketchName { get; set; }
        public string ConfigurationName { get; set; }
        public PartFeatureInfo(string featureName, string sketchName, string configurationName)
        {
            FeatureName = featureName;
            SketchName = sketchName;
            ConfigurationName = configurationName;
        }
    }
    public class SheetInfo
    {
        public string Name { get; set; }
        public List<ViewInfo> Views = new List<ViewInfo>();
        public bool FoundDanglingAnnotations { get; set; }

        public SheetInfo(string name)
        {
            Name = name;
            FoundDanglingAnnotations = false;
            Views = new List<ViewInfo>();
        }
    }
    public class ViewInfo
    {
        public string Name { get; set; }
        public bool FoundDanglingAnnotations { get; set; }

        public Dictionary<string, int> AnnotationCount { get; set; }
        public ViewInfo(string name)
        {
            FoundDanglingAnnotations = false;
            Name = name;
            AnnotationCount = new Dictionary<string, int>();

        }
    }
    #endregion
    public class ValidationBase
    {
        protected SolidworksService ThisSolidworksService { get; set; }

        public ValidationBase(SolidworksService thisSolidworksService)

        {
            ThisSolidworksService = thisSolidworksService;
        }
    }
    public class ECNFileValidation : ValidationBase
    {
        public enum ECNTxtFileIdxs
        {
            FileName,
            QADApproved,
            NewRevLetter,
            OldRevLetter,
            Obsolete,
            ENCNumber,
            QADRev
        }
 
        public ECNFileValidation(SolidworksService thisSolidworksService) : base(thisSolidworksService)
        {
            return;

        }
        public EcnFileValidationResult RunEcnFileValidation(List<List<string>> ecnData, List<string> filesInEcnFolder)
        {
            EcnFileValidationResult currentValidationResult = new EcnFileValidationResult();
        
            for (int i = 0; i < ecnData.Count; i++)
            {

                List<string> currentLine = ecnData[i];
                if (currentLine.Count < 6)
                {
                }

                string currentFileNameWOExt = Path.GetFileNameWithoutExtension(currentLine[0]);
                if (currentValidationResult.ECNFilesRequiredFound.ContainsKey(currentLine[(int)ECNTxtFileIdxs.FileName]))
                {
                    currentValidationResult.FoundDuplicateECNFile = true;
                    currentValidationResult.DuplicateEcnFile.Add(currentLine[(int)(ECNTxtFileIdxs.FileName)]);
                }

                if (currentLine[(int)(ECNTxtFileIdxs.NewRevLetter)] != currentLine[(int)(ECNTxtFileIdxs.QADRev)])
                {
                    currentValidationResult.FoundECNFileRevMismatch = true;
                    currentValidationResult.ECNFileRevMismatch.Add(currentLine[(int)(ECNTxtFileIdxs.FileName)]);
                }

                if (currentLine[(int)ECNTxtFileIdxs.QADApproved] != "Y")
                {

                    currentValidationResult.FoundFailedApproval = true;
                    currentValidationResult.FailedApproval.Add(currentLine[(int)(ECNTxtFileIdxs.QADApproved)]);
                }

                currentValidationResult.ECNFilesRequiredFound[currentFileNameWOExt] = false;


            }

            for (int i = 0; i < filesInEcnFolder.Count; i++)
            {

                string currentFile = filesInEcnFolder[i];
                string fileNameWOExt = Path.GetFileNameWithoutExtension(currentFile);
                string fileExt = Path.GetExtension(currentFile);
                if (currentValidationResult.ECNFilesRequiredFound.ContainsKey(fileNameWOExt) && (fileExt == SolidworksService.EXCELFILEEXT || fileExt == SolidworksService.DRAWINGFILEEXT))
                {
                    currentValidationResult.ECNFilesRequiredFound[fileNameWOExt] = true;
                }
                else if (!currentValidationResult.ECNFilesRequiredFound.ContainsKey(fileNameWOExt))
                {
                    currentValidationResult.FilesNotInECN.Add(fileNameWOExt);

                }
            }
            foreach (KeyValuePair<string, bool> filePair in currentValidationResult.ECNFilesRequiredFound)
            {
                if (filePair.Value == false)
                {
                    currentValidationResult.FoundMissingDrawingFile = true;
                    currentValidationResult.MissingDrawingFile.Add(filePair.Key);
                }
            }
 
            return currentValidationResult;
        }
    


    }
    public class SearchAndDependenciesValidation:ValidationBase
    {
        public SearchAndDependenciesValidation(SolidworksService thisSolidworksService): base(thisSolidworksService)
        {
            return;
        }
        public int DependentValidation(string referenceDocPath, string parentDocPath, Dictionary<string, int> folderCount, SearchAndDependenciesValidationResult currentValidationResult, EcnRelease ThisEcnRelease)

        {
            Utility utilityFunctions = new Utility();
            string referenceFileName = utilityFunctions.GetFileWithExt(referenceDocPath);
            string parentFileName = utilityFunctions.GetFileWithExt(parentDocPath);

            string folderKey = "";
            /*paths segments will split up path by folders and drives this works for current file architecture 
             * and may break down if file arechitecture changes.
           */
            string[] pathSegments = referenceDocPath.Split(Path.DirectorySeparatorChar);
            string fileNameWithExt = pathSegments[pathSegments.Length - 1];
            bool pathValidation = false;
            // if reference is in ecn folder it will be passed in as correct reference thru searchpaths
            if (ThisEcnRelease.Files.ContainsKey(referenceFileName))
            {
                pathValidation = true;
            }
            foreach (string archivePath in currentValidationResult.ArchivePaths)
            {
                // if parent matches an archive path allow to pass validation
                if (archivePath.Length > parentDocPath.Length)
                {
                    continue;
                }
                if (archivePath == parentDocPath.Substring(0, archivePath.Length))
                {
                    pathValidation = true;
                }
            }

            foreach (string validPath in currentValidationResult.ValidPaths)
            {
                string[] validPathSegments = validPath.Split(Path.DirectorySeparatorChar);
                int validPathSegmentsLength = validPathSegments.Length;
                string[] dependentPathSegments = referenceDocPath.Split(Path.DirectorySeparatorChar);
                int dependentPathLen = referenceDocPath.Length;
                bool isSegmentationEqual = true;
                // want to check every segment for every valid path for the reference to see if any match 
                for(int i = 0; i < validPathSegmentsLength;i++)
                {
                    if (validPathSegments[i] != dependentPathSegments[i])
                    {
                        isSegmentationEqual = false;
                    }

                }
                if (isSegmentationEqual)
                {
                    pathValidation = true;
                }

            }
            if (pathValidation)
            {
                if (pathSegments.Length > 1)
                {
                    if (int.TryParse(pathSegments[1], out int result))
                    {
                        folderKey = pathSegments[1];

                    }
                }
            }
            //This cannot be removed from virtual name and solidworks does not allow rename to include ^ char
            else if (fileNameWithExt.Contains('^'))
            {
                return 0;
            }
            else
            {

                if (ThisEcnRelease.Files.ContainsKey(parentFileName))
                {
                    parentDocPath = ThisEcnRelease.Files[parentFileName].FilePathSrc;
                }
                if (!currentValidationResult.InvalidDependencies.ContainsKey(parentDocPath))
                {
                    currentValidationResult.InvalidDependencies[parentDocPath] = new HashSet<string>();
                }
                currentValidationResult.InvalidDependencies[parentDocPath].Add(referenceDocPath);
                currentValidationResult.FoundInvalidDependencies = true;
                return 1;
            }
            if (folderKey.Length > 0)
            {
                if (folderCount.ContainsKey(folderKey))
                {
                    folderCount[folderKey] += 1;
                }
                else
                {
                    folderCount[folderKey] = 1;
                }
            }
            return 0;


        }

    }
    public class PartValidation: ValidationBase
    {
        public PartValidation(SolidworksService thisSolidworksService) : base(thisSolidworksService)
        {
            return;
        }

        public PartValidationResult RunPartValidation(ModelDoc2 doc)
        {
            if (doc == null)
            {
                MessageBox.Show("DOC is NULL for part validations");
            }

            PartValidationResult currentPartResult = new PartValidationResult();
            PartDoc swPart = (PartDoc)doc;
            SolidworksServiceResult<string[]> getConfNamesResult = ThisSolidworksService.GetConfigurationNames(doc);
            if (!getConfNamesResult.Success)
            {
                currentPartResult.CriticalError = true;
                return currentPartResult;
            }
            string[] configurationNames = getConfNamesResult.response;
            if (swPart == null)
            {
                currentPartResult.CriticalError = true;
                return currentPartResult;
            }
            foreach (string currentConfigurationName in configurationNames)
            {
                swPart = (PartDoc)doc;
                bool configSwitch = doc.ShowConfiguration2(currentConfigurationName);
                Configuration activeConfiguration = doc.GetActiveConfiguration();

                if (!configSwitch & activeConfiguration.Name != currentConfigurationName)
                {
                    currentPartResult.CriticalError = true;

                    return currentPartResult;
                }
                Feature currentFeature = (Feature)swPart.FirstFeature();
                if (currentFeature == null)
                {
                    currentPartResult.CriticalError = true;
                    return currentPartResult;
                }
                Feature currentSubFeature = (Feature)currentFeature.GetFirstSubFeature();
                while (currentFeature != null)
                {

                    if (currentFeature.GetSpecificFeature2() is Sketch)
                    {
                        Sketch currentFeatureSketch = (Sketch)currentFeature.GetSpecificFeature2();
                        int constrinStatus = (int)currentFeatureSketch.GetConstrainedStatus();
                    }



                    currentFeature.GetSpecificFeature();
                    string currentFeatureType = currentFeature.GetTypeName2();
                    currentSubFeature = (Feature)currentFeature.GetFirstSubFeature();
                    SolidworksServiceResult<bool[]> isFeatureSuppressedResult = ThisSolidworksService.IsFeatureSuppressed(currentFeature, currentConfigurationName);
                    bool[] isFeatureSuppressed = isFeatureSuppressedResult.response;
                    while (currentSubFeature != null && !isFeatureSuppressed[0] && !ThisSolidworksService.FeatureTypeExceptions.Contains(currentFeatureType))
                    {
                        SolidworksServiceResult <bool> isFeatureSketchResult = ThisSolidworksService.IsFeatureSketch(currentSubFeature);

                        if (isFeatureSketchResult.response)
                        {
                            SolidworksServiceResult<Sketch> getCurrentSketchResult = ThisSolidworksService.GetSketch(currentSubFeature);
                            Sketch currentSubFeatureSketch = getCurrentSketchResult.response;
                            SolidworksServiceResult<string> getFeatureTypeNameResult = ThisSolidworksService.GetFeatureTypeName(currentFeature);
                            string featureType = getFeatureTypeNameResult.response;
                            int subConstrainStatus = (int)currentSubFeatureSketch.GetConstrainedStatus();
                            if ((!ThisSolidworksService.SubFeatureTypeExceptions.Contains(featureType)) && subConstrainStatus != 3)
                            {
                                PartFeatureInfo currentFeatureInfo = new PartFeatureInfo(currentFeature.Name, currentSubFeature.Name, currentConfigurationName);
                                currentPartResult.FoundFeatureErrors = true;
                                currentPartResult.TotalSketchErrors.Add(currentFeatureInfo);
                            }

                        }
                        currentSubFeature = currentSubFeature.GetNextSubFeature();
                    }
                    currentFeature = currentFeature.GetNextFeature();
                }
            }


            return currentPartResult;

        }

    }
    
    public class ReleaseValidationService
    {
        private SolidworksService ThisSolidworksService { get; set; }

        public ReleaseValidationService(SolidworksService thisSolidworksService)

        {
            ThisSolidworksService = thisSolidworksService;
        }
        public AssemblyValidationResult CheckAssembly(ModelDoc2 doc)
        {
            AssemblyValidationResult currentAssemblyResult = new AssemblyValidationResult();
            bool isWarning = true;
            SolidworksServiceResult<string[]> getConfigNamesResult = ThisSolidworksService.GetConfigurationNames(doc);
            if (!getConfigNamesResult.Success)
            {
                currentAssemblyResult.CriticalError = true;
                return currentAssemblyResult;
            }
            string[] configurationNames = getConfigNamesResult.response;

            foreach (string configurationName in configurationNames)
            {

                bool configSwitch = doc.ShowConfiguration2(configurationName);
                Configuration activeConfiguration = doc.GetActiveConfiguration();
                AssemblyDoc currentAssembly = (AssemblyDoc)doc;
                if (activeConfiguration.IsSpeedPak()) continue;
                Feature currentFeature = doc.FirstFeature();
                object[] components = currentAssembly.GetComponents(true);
                object[] Mates = null;
                SolidworksServiceResult<HashSet<string>> suppressedMatesResult = ThisSolidworksService.getSuppressedComponentMates(doc);
                HashSet<string> suppressedMatesSet = suppressedMatesResult.response;

                if (!configSwitch & activeConfiguration.Name != configurationName)
                {
                    currentAssemblyResult.CriticalError = true;
                    return currentAssemblyResult;
                }
                foreach (object component in components)

                {
                    List<string> MateSuppress = new List<string>();
                    Component2 swComponent = (Component2)component;
                    string[] swComponentSplitName = swComponent.Name.Split('-');
                    string formattedSwComponentName = "";
                    for (int i = 0; i < swComponentSplitName.Length - 1; i++)
                    {
                        formattedSwComponentName += swComponentSplitName[i];
                    }
                    formattedSwComponentName += '<' + swComponentSplitName[1] + '>';

                    ComponentInfo currentComponentInfo = new ComponentInfo(formattedSwComponentName, "FULLY DEFINED", configurationName);
                    Mates = (Object[])swComponent.GetMates();
                    int solveResult = swComponent.GetConstrainedStatus();
                    if (swComponent.IsPatternInstance() || swComponent.IsSuppressed())
                    {
                        continue;
                    }
                    switch (solveResult)
                    {
                        case (int)swConstrainedStatus_e.swUnderConstrained:
                            currentComponentInfo.ConstraintStatus = "UNDERDEFINED";
                            currentComponentInfo.FoundError = true;
                            break;
                        case (int)swConstrainedStatus_e.swOverConstrained:
                            currentComponentInfo.ConstraintStatus = "OVERDEFINED";
                            currentComponentInfo.FoundError = true;

                            break;
                        default:
                            break;
                    }
                    if (currentComponentInfo.FoundError == true)
                    {
                        currentAssemblyResult.TotalComponentErrors.Add(currentComponentInfo);
                        currentAssemblyResult.FoundComponentErrors = true;
                        continue;
                    }
                    if (Mates != null)
                    {
                        foreach (Object SingleMate in Mates)

                        {
                            if (!(SingleMate is Mate2))
                            {
                                continue;
                            }
                            Feature mateFeat = (Feature)SingleMate;
                            int mateErrorCode = mateFeat.GetErrorCode2(out isWarning);
                            bool mateIsSuppressed = mateFeat.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration, null)[0];
                            string mateName = mateFeat.Name;
                            if (suppressedMatesSet.Contains(mateName))
                            {
                                continue;
                            }

                            if (!currentAssemblyResult.ValidMateErrorCodes.Contains(mateErrorCode) && mateIsSuppressed==false)
                            {
                                MessageBox.Show($"{mateName} code: {mateErrorCode}");
                                currentComponentInfo.FoundError = true;
                                currentComponentInfo.ConstraintStatus = "FULLY DEFINED WITH MATE ERRORS";
                                currentAssemblyResult.FoundComponentErrors = true;
                                currentAssemblyResult.TotalComponentErrors.Add(currentComponentInfo);
                            }
                        }


                    }

                 
                }


            }


            return currentAssemblyResult;
        }
        public DrawingValidationResult CheckDrawing(ModelDoc2 doc, string filePath)
        {
            DrawingValidationResult currentDrawingResult = new DrawingValidationResult();
            if (doc == null)
            {
                currentDrawingResult.CriticalError = true;
                return currentDrawingResult;
            }
            
            ModelDocExtension swModExt = default(ModelDocExtension);
            int danglingCount = 0;
            bool annotationsSeleted = false;
            SelectionMgr swSelmgr;
            SelectData swSelData;

            swSelmgr = (SelectionMgr)doc.SelectionManager;
            swModExt = (ModelDocExtension)doc.Extension;
            //bool deleteAnnotations = true;
            bool deleteAnnotations = true;
            int sheetIdx = 0;

            swSelData = swSelmgr.CreateSelectData();



      
            int totalAnnotations = 0;
            DrawingDoc swDrawingDoc = (DrawingDoc)doc;
            SolidworksServiceResult<object[]> getSheetsResult = ThisSolidworksService.GetDrawingSheets(swDrawingDoc);
            object[] sheets = getSheetsResult.response;
            SolidworksServiceResult<string[]> getSheetNameResult = ThisSolidworksService.GetDrawingSheetNames(swDrawingDoc);
            string[] sheetNames = getSheetNameResult.response;

            if (sheets == null || sheets.Length == 0)
            {
                currentDrawingResult.CriticalError = true;
                return currentDrawingResult;
            }
            foreach (object[] sheetObj in sheets)
            {
                if (sheetObj == null) continue;
                string currentSheetName = sheetNames[sheetIdx];
                SheetInfo currentSheetInfo = new SheetInfo(currentSheetName);
                currentDrawingResult.TotalSheets.Add(currentSheetInfo);
                bool sheetSet = swDrawingDoc.ActivateSheet(currentSheetName);
                Sheet currentSheet = swDrawingDoc.GetCurrentSheet();
                if (currentSheet.RevisionTable != null)
                {
                    //MessageBox.Show($"{currentSheet.RevisionTable}");
                }

                foreach (object view in sheetObj)
                {
                    if (view == null) continue;
                    SolidWorks.Interop.sldworks.View currentView = (SolidWorks.Interop.sldworks.View)view;
                    string currentViewName = currentView.Name;
                    ViewInfo currentViewInfo = new ViewInfo(currentViewName);
                    currentSheetInfo.Views.Add(currentViewInfo);
                    SolidworksServiceResult<int> getAnnotiationCountResult = ThisSolidworksService.GetAnnotationCount(currentView);
                    int annotationsCount = getAnnotiationCountResult.response;
                    SolidworksServiceResult<object[]> getViewAnnotationsResult = ThisSolidworksService.GetAnnotations(currentView);

                    object[] viewAnnotations = getViewAnnotationsResult.response;
                    if (viewAnnotations == null) continue;
                    foreach (object annotation in viewAnnotations)
                    {
                        if (annotation == null) continue;
                        Annotation currentAnnotation = (Annotation)annotation;
                        SolidworksServiceResult<bool> isAnnotationDanglingResult = ThisSolidworksService.isAnnotationDangling(currentAnnotation);

                        if (!isAnnotationDanglingResult.response) continue;

                        if (DanglingValidation(currentAnnotation, swSelmgr, swSelData,
                            currentSheetName, currentViewName)) continue;

                        else
                        {
                            currentDrawingResult.FoundDanglingAnnotations = true;
                            currentViewInfo.FoundDanglingAnnotations = true;
                            currentSheetInfo.FoundDanglingAnnotations = true;
                            SolidworksServiceResult<bool> selectAnnotationResult = ThisSolidworksService.SelectAnnotation(currentAnnotation, swSelData);
                            if (!selectAnnotationResult.response) continue;
                            string currentAnnotationType = ThisSolidworksService.GetAnnotationType(currentAnnotation);
                            if (currentViewInfo.AnnotationCount.ContainsKey(currentAnnotationType))
                            {
                                currentViewInfo.AnnotationCount[currentAnnotationType]++;

                            }
                            else
                            {
                                currentViewInfo.AnnotationCount[currentAnnotationType] = 1;
                            }
                        }
                    }
                }
                if (deleteAnnotations)
                {
                    ThisSolidworksService.DeleteAllSelections(doc);
                }
                
                sheetIdx++;
            }
            if (deleteAnnotations)
            {
                ThisSolidworksService.SaveSWDocument(doc);
            }
       
            return currentDrawingResult;
        }
        private bool DanglingValidation(Annotation currentAnnotation, SelectionMgr swSelmgr, SelectData swSelData,
           string currentSheetName, string currentViewName)
        {
            string danglingItem = "";
            bool validAnnotationCheck = true;
            if (ThisSolidworksService.ReleaseAnnotations.ContainsKey((int)currentAnnotation.GetType()))
            {
                if ((int)currentAnnotation.GetType() == (int)swAnnotationType_e.swNote)
                {
                    Note currentNote = (Note)currentAnnotation.GetSpecificAnnotation();
                    if (currentNote.IsBomBalloon() || currentNote.IsStackedBalloon())
                    {
                        validAnnotationCheck = false;
                    }
                }
                else
                {
                    validAnnotationCheck = false;
                }
            }
            else
            {
                validAnnotationCheck = true;
            }
            return validAnnotationCheck;
        }

        
    }

}


