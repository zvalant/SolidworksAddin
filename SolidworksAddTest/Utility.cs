using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidworksAddTest
{
    public class Utility
    {
        public string GetFileWithExt(string docName)
        {
            string[] path = docName.Split(new char[] { '\\' });
            string fileName = path[path.Length - 1];
            return fileName;

        }
        public string GetFileNameWithoutExt(string docName)
        {
            string[] path = docName.Split(new char[] {'.'});
            string fileName = path[0];
            return fileName;
        }
        public string GetFileExt(string docName)
        {
            string fileName = GetFileWithExt((string)docName);
            string[] fileSegments = fileName.Split(new char[] { '.' });
            return fileSegments[fileSegments.Length - 1];

        }
    }
}
