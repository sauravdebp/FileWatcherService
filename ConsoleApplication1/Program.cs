using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SecurityReader;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExcelReader reader = new ExcelReader();
            //reader.OpenFile(@"C:\Users\saura_000\Downloads\Data for securities.xlsx");
            //reader.SheetName = "Equities";
            //var data = reader.ReadFile();
            //reader.CloseFile();
            SecurityReader.SecurityReader reader = new SecurityReader.SecurityReader();
            reader.ReadSecuritiesFromFile(@"C:\Users\saura_000\Downloads\Equities.xlsx");
        }
    }
}
