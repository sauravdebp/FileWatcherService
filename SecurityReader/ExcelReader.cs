using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SecurityReader
{
    public class ExcelReader : FileReader
    {
        string connectionString;
        OleDbDataAdapter adapter;
        DataTable data;
        public string SheetName { get; set; }

        public override Dictionary<string, List<Dictionary<string, string>>> ReadFile()
        {
            adapter = new OleDbDataAdapter("SELECT * FROM [" + SheetName + "$]", connectionString);
            data = new DataTable(SheetName);
            adapter.Fill(data);

            Dictionary<string, List<Dictionary<string, string>>> dataDict = new Dictionary<string, List<Dictionary<string, string>>>();
            dataDict.Add(SheetName, new List<Dictionary<string, string>>());
            foreach(DataRow row in data.Rows)
            {
                dataDict[SheetName].Add(new Dictionary<string, string>());
                foreach(DataColumn col in data.Columns)
                {
                    //dataDict[SheetName].Add(col.ColumnName, row[col].ToString());
                    dataDict[SheetName].Last().Add(col.ColumnName, row[col].ToString());
                }
            }

            data.Dispose();
            return dataDict;
        }

        public override bool OpenFile(string filePath)
        {
            connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties = \"Excel 12.0 Xml;HDR=YES\"; ", filePath);
            return true;
        }

        public override bool CloseFile()
        {
            adapter.Dispose();
            return true;
        }
    }
}
