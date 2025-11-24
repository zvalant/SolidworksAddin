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
        public int ReleaseType { get; set; }

        public Dictionary<string, EcnFile> Files { get; set; }
        public HashSet<EcnFile> LeafFiles { get; set; }
        public HashSet<EcnFile> CompletedFiles { get; set; }
        public Queue<EcnFile> ProcessingFileQueue { get; set; }
        public Stack<EcnFile> OpenFilesStack { get; set; }
        public EcnRelease(string releaseNumber, int releaseType)
        {
            ReleaseNumber = releaseNumber;
            ReleaseType = releaseType;
            ReleaseFolderSrc = $"R:\\{releaseNumber}";
            ReleaseFolderTemp = $"C:\\releaseECN\\{releaseNumber}";
            ReleaseTxtFile = $"G:\\Parent ECR Files\\{releaseNumber}_Parent_ECN.txt";
            Files = new Dictionary<string, EcnFile>();
            LeafFiles = new HashSet<EcnFile>();
            CompletedFiles = new HashSet<EcnFile>();
            ProcessingFileQueue = new Queue<EcnFile>();
            OpenFilesStack = new Stack<EcnFile>();
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
        public void AddCompletedFile(EcnFile file)
        {
            CompletedFiles.Add(file);
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
