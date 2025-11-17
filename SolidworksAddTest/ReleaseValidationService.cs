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
    public class PartFeatureInfo
    {
        public string FeatureName { get; set; }
        public string SketchName { get; set; }
        public PartFeatureInfo(string featureName, string sketchName) 
        { 

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
        public int CheckAssembly(ModelDoc2 doc)
        {
            AssemblyDoc currentAssembly = (AssemblyDoc)doc;
            Feature currentFeature = doc.FirstFeature();
            object[] components = currentAssembly.GetComponents(true);
            object[] Mates = null;
            List<string> mateSpace = new List<string>();
            List<string> assyName = new List<string>();
            int validRelease = 0;
            int componenetError = 0;
            int mateError = 0;
            string previousMate = null;
            string currentMateError = null;
            List<string> compResult = new List<string>();
            HashSet<string> suppressedMatesSet = new HashSet<string>();
            suppressedMatesSet = ThisSolidworksService.getSuppressedMates(doc);

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
                    formattedSwComponentName += swComponentSplitName[i];
                }
                formattedSwComponentName += '<' + swComponentSplitName[1] + '>';


                Mates = (Object[])swComponent.GetMates();
                int solveResult = swComponent.GetConstrainedStatus();
                string partMessage = $"Part: {swComponent.Name2} resolve: {solveResult}";
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
                        bool[] isSuppressed = mateFeat.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration, null);
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

                        if (errorCodes != 0 && componenetError == 0)
                        {
                            componenetError = 1;
                            componentErrors.Add($"  {formattedSwComponentName} PROPERLY DEFINED BUT HAS MATE ERRORS ");
                            componentErrors.Add($"Mate: {mateFeat.Name} errorNUM: {componenetError} comp: {swComponent.Name2}");
                            continue;
                        }

                        previousMate = mateFeat.Name;
                    }


                }

                if (componenetError != 0)
                {
                    if (mateErrors.Count > 0)
                    {
                        //releaseReport.WriteToReport(mateSpace);
                        //releaseReport.WriteToReport(mateErrors);

                    }
                    if (componenetError != 0)
                    {
                        validRelease = 1;
                    }
                }
            }


            return validRelease;
        }
        public DrawingValidationResult CheckDrawing(ModelDoc2 doc, string filePath)
        {
            DrawingValidationResult currentDrawingResult = new DrawingValidationResult();
            ModelDocExtension swModExt = default(ModelDocExtension);
            int danglingCount = 0;
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
                currentDrawingResult.CriticalError = true;
                return currentDrawingResult;
            }
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
                sheetIdx++;


            }

            /*
            if (deleteAnnotations)
            {
                ThisSolidworksService.DeleteAllSelections(doc);
            }
            */
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
                        return false;
                    }
                }
                else
                {
                    validAnnotationCheck = false;
                    return false;
                }

            }
            else
            {
                validAnnotationCheck = true;
                return true;
            }
            if (!validAnnotationCheck)
            {
                //select annotation that did not pass validation.
                currentAnnotation.Select3(true, swSelData);
            }

            return validAnnotationCheck;
        }
        public PartValidationResult CheckPart(ModelDoc2 doc, string filepath)
        {
            PartValidationResult currentPartResult = new PartValidationResult();
            PartDoc swPart = (PartDoc)doc;
            if (swPart == null)
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
                currentSubFeature = (Feature)currentFeature.GetFirstSubFeature();
                while (currentSubFeature != null)
                {
                    if (currentSubFeature.GetSpecificFeature2() is Sketch)
                    {
                        Sketch currentSubFeatureSketch = (Sketch)currentSubFeature.GetSpecificFeature2();
                        string featureType = currentFeature.GetTypeName2();
                        int subConstrainStatus = (int)currentSubFeatureSketch.GetConstrainedStatus();
                        if ((!ThisSolidworksService.FeatureTypeExceptions.Contains(featureType)) && subConstrainStatus!=3)
                        {
                            PartFeatureInfo currentFeatureInfo = new PartFeatureInfo(currentFeature.Name, currentSubFeature.Name);
                            currentPartResult.FoundFeatureErrors = true;
                            currentPartResult.TotalSketchErrors.Add(currentFeatureInfo);
                        }
                  
                    }
                    currentSubFeature = currentSubFeature.GetNextSubFeature();
                }
                currentFeature = currentFeature.GetNextFeature();
            }
            return currentPartResult;

        }
    }
}


