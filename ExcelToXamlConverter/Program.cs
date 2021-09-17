using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelToXamlConverter
{
  class Program
  {
    static void Main(string[] args)
    {
      using (ExcelExtractor extractor = new ExcelExtractor())
      {
        extractor.ReadXaml();
        //extractor.ReadFile();
        //extractor.WriteToXaml();
       // extractor.SaveToNewExcelFile();
      }
    }

  
  }
}
