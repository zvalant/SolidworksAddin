using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolidworksAddTest
{
    public class ReleaseReport
    {
        public string EcnNumber { get; set; }
        public string ReleaseType { get; set; }
        private string reportFilePath { get; set; }
        public DateTime runTime { get; set; }
        public DateTime startTime { get; set; }
        public Dictionary<EcnFile, List<string>> Files { get; set; }
        public ReleaseReport(string ecnNumber, bool isReadiness)
        {
            reportFilePath = @"C:\Users\zacv\Documents\releaseTest";
            EcnNumber = ecnNumber;
            Files = new Dictionary<EcnFile, List<string>>();
            if (isReadiness)
            {
                ReleaseType = "RELEASE";
            }
            else
            {
                ReleaseType = "READINESS FOR RELEASE";
            }
            string reportFolder = @"C:\Users\zacv\Documents\releaseTest";
            if (!System.IO.Directory.Exists(reportFolder))
            {
                System.IO.Directory.CreateDirectory(reportFolder);
            }
            startTime = DateTime.Now;
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            reportFilePath = System.IO.Path.Combine(reportFolder, $"{EcnNumber}_{timestamp}.txt");

            // Create initial report file
            CreateReportFile();

        }
        private void CreateReportFile()
        {
            try
            {
                using (System.IO.StreamWriter writer = new System.IO.StreamWriter(reportFilePath))
                {
                    writer.WriteLine($"ECN {ReleaseType}: {EcnNumber}");
                    writer.WriteLine(reportFilePath);
                    writer.WriteLine("=".PadRight(90, '='));
                    writer.WriteLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating report: {ex.Message}");
            }
        }
        public void WriteSectionHeader(string sectionName)
        {
            try
            {
                List<string> reportLines = new List<string>();
                
                reportLines.Add("=".PadRight(90, '='));
                reportLines.Add($"{sectionName}");
                reportLines.Add("");
                WriteToReportMultiline(reportLines);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating report: {ex.Message}");
            }
        }
        public void FinishReport()
        {
            int DateTimeRunTime = (int)(DateTime.Now - startTime).TotalMilliseconds;

            string runtimeString = $"Total Runtime: {DateTimeRunTime / 1000}.{DateTimeRunTime % 1000} S";


            List<string> FinalRuntime = new List<string>();
            FinalRuntime.Add("");
            FinalRuntime.Add("");
            FinalRuntime.Add(runtimeString);
            WriteToReportMultiline(FinalRuntime);
        }

        public void AddFile(EcnFile file)
        {
            Files[file] = new List<string>();
        }
        public void WriteToReportMultiline(List<string> lines)
        {
            try
            {
                System.IO.File.AppendAllLines(reportFilePath, lines);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error writing to report: {ex.Message}");
            }
        }
        public void WriteToReportSingleline(string line)
        { 
            List<string> lines = new List<string>();
            lines.Add(line);
            try
            {
                System.IO.File.AppendAllLines(reportFilePath, lines);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error writing to report: {ex.Message}");
            }
        }
public void OpenReport()
        {
            try
            {
                if (System.IO.File.Exists(reportFilePath))
                {
                    System.Diagnostics.Process.Start("notepad.exe", reportFilePath);
                }
                else
                {
                    MessageBox.Show("Report file not found");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening report: {ex.Message}");
            }
        }
        public void WriteValidationStatus(bool validationStatus, List<string> reportLines)
        {
            if (validationStatus)
            {
                reportLines.Insert(0, ("Validation Status: Passed"));
                reportLines.Add("");
            }
            else
            {
                reportLines.Insert(0, ("Validation Status: Failed"));
                reportLines.Add("");
            }
        }
    }
}
