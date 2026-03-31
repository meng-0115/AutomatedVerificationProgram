using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using Spire.Xls;

namespace AutomatedVerificationProgram
{
    internal class FindListName
    {
        public static List<string> findCompositionOfHomogeneous(string vendorCode, string partNumber)
        {
            List<string> compositionOfHomogeneous = new List<string>();
          
            Spire.Xls.Workbook theListOfMaterialHierarchy = new Spire.Xls.Workbook();
            foreach (string fname in System.IO.Directory.GetFileSystemEntries(@"./list/", "The List of Homogeneous Type_" + vendorCode + "*.xlsx"))
            {
                theListOfMaterialHierarchy.LoadFromFile(fname);
                Spire.Xls.Worksheet materialHierarchy = theListOfMaterialHierarchy.Worksheets["Material Hierarchy"];
                for (int i = 3; i < materialHierarchy.Rows.Length; i++)
                {
                    if (materialHierarchy.Rows[i].Columns[3].Value != "")
                    {
                        if (materialHierarchy.Rows[i].Columns[3].Value.ToUpper().Contains(partNumber.ToUpper()))
                        {
                            compositionOfHomogeneous = materialHierarchy.Rows[i].Columns[3].Value.Split(',').ToList();                           
                        }

                    }
                }
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return compositionOfHomogeneous;
        }
    }
}
