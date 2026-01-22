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
        public string FilePathSrc { get; set; } 
        public int LoadedFilesRemaining { get; set; }
        public List<string> SearchPaths { get; set; }
        public HashSet<EcnFile> Parents { get; set; }
        public swDocumentTypes_e SWDocumentType { get; set; }
        public bool IsExcelDocument { get; set; }
        public bool IsDrawingDocument { get; set; }
        public string Revision { get; set; }
        public EcnFile()
        {
            SearchPaths = new List<string>();
            Parents = new HashSet<EcnFile>();
            LoadedFilesRemaining = 0;
            IsExcelDocument = false;
            Revision = "_";
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
