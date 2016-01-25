using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SecurityReader
{
    public abstract class FileReader
    {
        public abstract bool OpenFile(string filePath);
        public abstract bool CloseFile();
        public abstract Dictionary<string, List<Dictionary<string, string>>> ReadFile();
    }
}
