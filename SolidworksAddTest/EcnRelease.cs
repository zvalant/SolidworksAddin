using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidworksAddTest
{
    public class EcnRelease
    {
        public string ReleaseNumber { get; set; }
        public string ReleaseFolderSrc { get; set; }
        public string ReleaseFolderTemp { get; set; }
        public string ReleaseTxtFile { get; set; }
        public bool ReadinessForRelease { get; set; }
        public Dictionary<string, EcnFile> Files { get; set; }
        public HashSet<string> FileNames {  get; set; }
        public HashSet<EcnFile> LeafFiles { get; set; }
        public HashSet<EcnFile> ReleasedFiles { get; set; }
        public Queue<EcnFile> ProcessingFileQueue { get; set; }
        public Stack<EcnFile> OpenFilesStack { get; set; }
        public HashSet<string> validReleaseExtensions { get; set; }

        public EcnRelease(string releaseNumber, bool isReadiness)
        {
            ReleaseNumber = releaseNumber;
            ReadinessForRelease = isReadiness;
            ReleaseFolderSrc = $"R:\\{releaseNumber}";
            ReleaseFolderTemp = $"C:\\releaseECN\\{releaseNumber}";
            ReleaseTxtFile = $"G:\\Parent ECR Files\\{releaseNumber}_Parent_ECN.txt";
            Files = new Dictionary<string, EcnFile>();
            FileNames = new HashSet<string>();
            LeafFiles = new HashSet<EcnFile>();
            ReleasedFiles = new HashSet<EcnFile>();
            ProcessingFileQueue = new Queue<EcnFile>();
            OpenFilesStack = new Stack<EcnFile>();
            validReleaseExtensions = new HashSet<string>();

            List<string> validReleaseExtensionsList = new List<string>{ "SLDPRT", "SLDASM", "SLDDRW","xls" };
            foreach (string validReleaseExtension in validReleaseExtensionsList)
            {
                validReleaseExtensions.Add(validReleaseExtension);
            }

        }

        public void AddFile(EcnFile file, string fileName)
        {
            Files[fileName] = file;
        }
        public void AddLeafFile(EcnFile file)
        {
            LeafFiles.Add(file);
        }
        public void RemoveLeafFile(EcnFile file)
        {
            LeafFiles.Remove(file);
        }
        public void AddReleasedFile(EcnFile file)
        {
            ReleasedFiles.Add(file);
        }
        public void ProcessFilesPush(EcnFile file)
        {
            ProcessingFileQueue.Enqueue(file);
        }
        public EcnFile ProcessFilesPop()
        {
            return ProcessingFileQueue.Dequeue();
        }
        public void PushOpenFileStack(EcnFile file)
        {
            OpenFilesStack.Push(file);
        }
        public EcnFile PopOpenFileStack()
        {
            return OpenFilesStack.Pop();
        }

    }
}
