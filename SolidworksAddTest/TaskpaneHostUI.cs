
using System;
using System.Windows.Forms;

namespace SolidworksAddTest
{
    public partial class TaskpaneHostUI : UserControl
    {
        private SWTestRP parentAddin;
        private Panel contentPanel;
        private Button switchButton;
        private bool mainViewFlag;

        public TaskpaneHostUI()
        {
            InitializeComponent();
            InitializeDynamicUI();
        }

        public void SetParentAddin(SWTestRP parent)
        {
            parentAddin = parent;
        }

        private void InitializeDynamicUI()
        {
            // Create and configure the content panel
            contentPanel = new Panel
            {
                Dock = DockStyle.Fill
            };
            this.Controls.Add(contentPanel);

            // Create and configure the switch button
            switchButton = new Button
            {
                Text = "Switch View",
                Dock = DockStyle.Top
            };
            switchButton.Click += SwitchButton_Click;
            this.Controls.Add(switchButton);

            // Load the default view
            LoadMainView();
        }

        private void LoadMainView()
        {
            mainViewFlag = true;
            contentPanel.Controls.Clear();

            Label mainLabel = new Label
            {
                Text = "Main View Loaded",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };

            contentPanel.Controls.Add(mainLabel);
        }

        private void SwitchButton_Click(object sender, EventArgs e)
        {
            if (mainViewFlag)
            {
                LoadDependenciesResultView();

            }
            else 
            {
                LoadMainView();
            }
        }

        private void LoadDependenciesResultView()
        {
            try
            {
                mainViewFlag = false;
                contentPanel.Controls.Clear();

                DependenciesResult resultsControl = new DependenciesResult();
                resultsControl.SetParentAddin(parentAddin);
                resultsControl.Dock = DockStyle.Fill;

                contentPanel.Controls.Add(resultsControl);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading DependenciesResult: {ex.Message}");
            }
        }
        private void button1_Click(object sender, EventArgs e)
        { 
        }
        private void TaskpaneHostUI_Load(object sender, EventArgs e) 
        { 
        }
    }
}
