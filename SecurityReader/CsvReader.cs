using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SecurityReader
{
    class CsvReader : FileReader
    {
        public override bool CloseFile()
        {
            throw new NotImplementedException();
        }

        public override bool OpenFile(string filePath)
        {
            throw new NotImplementedException();
        }

        public override Dictionary<string, List<Dictionary<string, string>>> ReadFile()
        {
            throw new NotImplementedException();
        }
    }
}
