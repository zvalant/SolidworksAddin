using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;

namespace SolidworksAddTest
    ///sw testground: this is just meant to benchmark subtasks in release program and is not representative of final strucutre. 
{
    [ComVisible(true)]
    [Guid("12345678-1234-1234-1234-1234523789AC")]
    [ProgId("SolidworksAddTest.SWTestRP")]
    public class SWTestRP: SwAddin
    {
        public SldWorks SolidWorksApplication => mSolidworksApplication;

        #region Private Members
        /// <summary>
        /// cookie to the current instance of SW
        /// </summary>
        private int mSWCookie;
        /// <summary>
        /// taskpane view for the addin
        /// </summary>
        private TaskpaneView mTaskpaneView;
        /// <summary>
        /// the ui control that is going to be in SW taskpane
        /// </summary>
        private object mTaskpaneHost;
        /// <summary>
        /// Current Instance of Solidworks applicaiton
        /// </summary>
        private SldWorks mSolidworksApplication;
        #endregion
        #region Public Members
        /// <summary>
        /// unique id for addin progam for the task id
        /// </summary>
        public const string SWTASKPANE_PROGID = "SolidworksAddTest.TaskpaneHostUI";

        #endregion
        #region Solidworks Add-in Callbacks


        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            mSolidworksApplication = (SldWorks)ThisSW;
            //store cookie id
            mSWCookie = Cookie;


            var ok = mSolidworksApplication.SetAddinCallbackInfo2(0,this,mSWCookie);
            LoadUI();
            return true;
        }


        public bool DisconnectFromSW()
        {
            UnloadUI();
            return true;
        }
        #endregion
        #region UI Creation
        private void LoadUI() {
            //find location of icon
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(SWTestRP).Assembly.CodeBase).Replace(@"file:\", ""), "baaderlogo.png");
            //create taskpane
            mTaskpaneView = mSolidworksApplication.CreateTaskpaneView2(imagePath, "SWADDIN");
            System.Windows.Forms.MessageBox.Show($"Looking for image at: {imagePath}\nFile exists: {File.Exists(imagePath)}");


            //Load our UI into the Taskpane
            mTaskpaneHost = (TaskpaneHostUI)mTaskpaneView.AddControl(SWTestRP.SWTASKPANE_PROGID, string.Empty);
            ((TaskpaneHostUI)mTaskpaneHost).SetParentAddin(this);
            //throw new NotImplementedException();

        }
        private void UnloadUI() {
            mTaskpaneHost = null;
            //remove taskpane view
            mTaskpaneView.DeleteView();
            // release com object and clean up memory
            Marshal.ReleaseComObject(mTaskpaneView);
            mTaskpaneView = null;

            //throw new NotImplementedException();
        }
        #endregion
        #region COM Registration
        [ComRegisterFunctionAttribute()]
        private static void ComRegister(Type t) { 
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);
            using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(keyPath))
            {
                rk.SetValue(null, t);
                rk.SetValue("Title", "My SW Addin");
                rk.SetValue("Description", "All your pixels are mine");
            }

        }
        private static void ComUnregister(Type t)
        {
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(keyPath);
        }
        #endregion
        public void ChangeToResultsView()
        {
            try
            {
                if (mTaskpaneHost != null)
                {
                    mTaskpaneHost =  null;
                    
                }

                System.Windows.Forms.MessageBox.Show("About to call AddControl for EcnProcessService...");

                var control = mTaskpaneView.AddControl("SolidworksAddTest.EcnProcessService", string.Empty);

                System.Windows.Forms.MessageBox.Show($"AddControl returned type: {control.GetType().FullName}");

                // Only try to cast if it's the right type
 
               
                var resultsControl = (EcnProcessService)control;
                resultsControl.SetParentAddin(this);
                mTaskpaneHost = resultsControl;
                System.Windows.Forms.MessageBox.Show("Successfully switched to EcnProcessService!");
             
             
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}");
            }
        }
        public void ChangeToMainView()
        {
            // Clear current control  
            mTaskpaneHost = null;

            // Add new control
            var mainControl = (TaskpaneHostUI)mTaskpaneView.AddControl(SWTASKPANE_PROGID, string.Empty);
            mainControl.SetParentAddin(this);
            mTaskpaneHost = mainControl;
        }
    }
}       
