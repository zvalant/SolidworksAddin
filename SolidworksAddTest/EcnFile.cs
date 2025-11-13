using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidworksAddTest
{
    public class EcnFile
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public int LoadedFilesRemaining { get; set; }
        public List<string> SearchPaths { get; set; }
        public HashSet<EcnFile> Parents { get; set; }
        public swDocumentTypes_e DocumentType { get; set; }
        public EcnFile()
        {
            SearchPaths = new List<string>();
            Parents = new HashSet<EcnFile>();
            LoadedFilesRemaining = 0;
        }
        public void InsertSearchPaths(List<string> searchPaths)
        {
            SearchPaths = searchPaths;
        }
        public void InsertParent(EcnFile parent)
        {
            Parents.Add(parent);
        }


    }
}
