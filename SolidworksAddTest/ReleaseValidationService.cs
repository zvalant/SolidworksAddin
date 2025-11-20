using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolidworksAddTest
{
    public class DrawingValidationResult
    {
        public bool CriticalError { get; set; }
        public bool FoundDanglingAnnotations { get; set; }
        public List<SheetInfo> TotalSheets { get; set; }

        public DrawingValidationResult()
        {
            CriticalError = false;
            FoundDanglingAnnotations = false;
            TotalSheets = new List<SheetInfo>();
        }
    }
    public class PartValidationResult
    {
        public bool CriticalError { get; set; }
        public bool FoundFeatureErrors { get; set; }
        public List<PartFeatureInfo> TotalSketchErrors { get; set; }
        public PartValidationResult()
        { 
            CriticalError=false;
            FoundFeatureErrors=false;
            TotalSketchErrors = new List<PartFeatureInfo>();
        }
    }
    public class AssemblyValidationResult
    {
        public bool CritalError { get; set; }
        public bool FoundComponentErrors { get; set; }
        public List<ComponentInfo> TotalComponentErrors { get; set; }
        public AssemblyValidationResult()
        {
            CritalError = false;
            FoundComponentErrors = false;
            TotalComponentErrors = new List<ComponentInfo>();
        }
    }
    public class ComponentInfo
    { 
        public string Name { get; set; }
        public string ConstraintStatus { get; set; }
        public string Configuration {  get; set; }
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

            string[] configurationNames = ThisSolidworksService.GetConfigurationNames(doc);

            foreach (string configurationName in configurationNames)
            {

                bool configSwitch = doc.ShowConfiguration2(configurationName);
                Configuration activeConfiguration = doc.GetActiveConfiguration();
                AssemblyDoc currentAssembly = (AssemblyDoc)doc;
                if (activeConfiguration.IsSpeedPak()) continue;
                Feature currentFeature = doc.FirstFeature();
                object[] components = currentAssembly.GetComponents(true);
                object[] Mates = null;
                HashSet<string> suppressedMatesSet = new HashSet<string>();

                suppressedMatesSet = ThisSolidworksService.getSuppressedComponentMates(doc);

                if (!configSwitch & activeConfiguration.Name != configurationName)
                {
                    currentAssemblyResult.CritalError = true;
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
                            int mateErrorCode = mateFeat.GetErrorCode();
                            bool mateIsSuppressed = mateFeat.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration, null)[0];
                            string mateName = mateFeat.Name;
                            if (suppressedMatesSet.Contains(mateName))
                            {
                                continue;
                            }

                            if (mateErrorCode != 0 && mateIsSuppressed==false)
                            {
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
            object[] sheets = ThisSolidworksService.GetDrawingSheets(swDrawingDoc);
            string[] sheetNames = ThisSolidworksService.GetDrawingSheetNames(swDrawingDoc);

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

                foreach (object view in sheetObj)
                {
                    if (view == null) continue;
                    SolidWorks.Interop.sldworks.View currentView = (SolidWorks.Interop.sldworks.View)view;
                    string currentViewName = currentView.Name;
                    ViewInfo currentViewInfo = new ViewInfo(currentViewName);
                    currentSheetInfo.Views.Add(currentViewInfo);
                    int annotationsCount = ThisSolidworksService.GetAnnotationCount(currentView);
                    object[] viewAnnotations = ThisSolidworksService.GetAnnotations(currentView);
                    if (viewAnnotations == null) continue;
                    foreach (object annotation in viewAnnotations)
                    {
                        if (annotation == null) continue;
                        Annotation currentAnnotation = (Annotation)annotation;
                        if (!ThisSolidworksService.isAnnotationDangling(currentAnnotation)) continue;

                        if (DanglingValidation(currentAnnotation, swSelmgr, swSelData,
                            currentSheetName, currentViewName)) continue;

                        else
                        {
                            currentDrawingResult.FoundDanglingAnnotations = true;
                            currentViewInfo.FoundDanglingAnnotations = true;
                            currentSheetInfo.FoundDanglingAnnotations = true;
                            ThisSolidworksService.SelectAnnotation(currentAnnotation, swSelData);
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
        public PartValidationResult CheckPart(ModelDoc2 doc, string filepath)
        {
            PartValidationResult currentPartResult = new PartValidationResult();
            PartDoc swPart = (PartDoc)doc;
            string[] configurationNames = ThisSolidworksService.GetConfigurationNames(doc);
            if (swPart == null)
            {
                currentPartResult.CriticalError = true;
                return currentPartResult;
            }
            foreach (string configurationName in configurationNames)
            {
                swPart = (PartDoc)doc;
                bool configSwitch = doc.ShowConfiguration2(configurationName);
                Configuration activeConfiguration = doc.GetActiveConfiguration();

                if (!configSwitch &  activeConfiguration.Name!=configurationName)
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
                    bool[] isFeatureSuppressed = currentFeature.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration, configurationName);
                    while (currentSubFeature != null && !isFeatureSuppressed[0] && !ThisSolidworksService.FeatureTypeExceptions.Contains(currentFeatureType))
                    {
                        if (currentSubFeature.GetSpecificFeature2() is Sketch)
                        {
                            Sketch currentSubFeatureSketch = (Sketch)currentSubFeature.GetSpecificFeature2();
                            string featureType = currentFeature.GetTypeName2();
                            int subConstrainStatus = (int)currentSubFeatureSketch.GetConstrainedStatus();
                            if ((!ThisSolidworksService.SubFeatureTypeExceptions.Contains(featureType)) && subConstrainStatus != 3)
                            {
                                PartFeatureInfo currentFeatureInfo = new PartFeatureInfo(currentFeature.Name, currentSubFeature.Name, configurationName);
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
}


