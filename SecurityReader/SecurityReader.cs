using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SecMaster_DAL.DataModel;

namespace SecurityReader
{
    public class SecurityReader
    {
        FileReader reader;

        public SecurityReader()
        {

        }

        public List<Security> ReadSecuritiesFromFile(string filePath)
        {
            List<Security> securities = new List<Security>();
            string securityName = Path.GetFileNameWithoutExtension(filePath);
            Dictionary<string, List<Dictionary<string, string>>> securitiesData = new Dictionary<string, List<Dictionary<string, string>>>();
            if(File.Exists(filePath))
            {
                InstantiateReader(Path.GetExtension(filePath));
                reader.OpenFile(filePath);
                if (reader.GetType() == typeof(ExcelReader))
                    (reader as ExcelReader).SheetName = securityName;
                securitiesData = reader.ReadFile();
                reader.CloseFile();

                foreach(var d in securitiesData[securityName])
                {
                    Security security = GetSecurityObject(securityName);
                    //security.GetType().GetProperties()["Name"].SetValue(security, null);
                }
            }
            
            return securities;
        }

        FileReader InstantiateReader(string fileExtension)
        {
            switch(fileExtension)
            {
                case ".csv":
                    reader = new CsvReader();
                    break;
                case ".xls":
                case ".xlsx":
                    reader = new ExcelReader();
                    break;
            }
            return reader;
        }

        Security GetSecurityObject(string securityName)
        {
            Security security = null;
            switch(securityName)
            {
                case "Equities":
                    security = new Equity();
                    break;
                case "Bonds":
                    security = new CorporateBond();
                    break;
            }
            return security;
        }
    }
}
