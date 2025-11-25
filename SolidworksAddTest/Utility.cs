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
    }
}
