using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xaml;
using System.Windows.Markup;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using XamlWriter = System.Windows.Markup.XamlWriter;

namespace ExcelToXamlConverter
{
  public class ExcelExtractor : IDisposable
  {
    private static string docLocation;
    private Application xclApp;
    private Workbook xlWorkbook;
    private _Worksheet xlWorksheet;
    private Range xlRange;

    public ExcelExtractor()
    {
      xclApp = new Application();
      randomFolder = DateTime.Now.ToString(CultureInfo.InvariantCulture).Replace(":", "_").Replace("/", "_").Replace(" ", "_");
    }

    public void ReadFile()
    {
      try
      {
        Console.WriteLine("Location of your .xlsx document");
        docLocation = Console.ReadLine();
        xlWorkbook = xclApp.Workbooks.Open(docLocation);
        xlWorksheet = xlWorkbook.Sheets[1];
        xlRange = xlWorksheet.UsedRange;
        ReadRows();
      }
      catch (Exception e)
      {
        Console.WriteLine(e.ToString());
        ReadFile();
      }
    }

    private void ReadRows()
    {
      var colCount = xlRange.Columns.Count;
      var rowCount = xlRange.Rows.Count;
      ExcelFileRowView currentRow;
      //iterate over the rows and columns and print to the console as it appears in the file
      //excel is not zero based!!
      for (int i = 2; i <= rowCount; i++)
      {
        currentRow = new ExcelFileRowView();
        for (int j = 1; j <= colCount; j++)
        {
          if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
          {
            switch (j)
            {
              case 1:
                currentRow.Project = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 2:
                currentRow.File = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 3:
                currentRow.Key = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 5:
                currentRow.Basic = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 7:
                currentRow.German = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 9:
                currentRow.Spanish = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 11:
                currentRow.Italian = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 13:
                currentRow.Japanese = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 15:
                currentRow.Portuguese = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 17:
                currentRow.Russian = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 19:
                currentRow.Turkish = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 21:
                currentRow.Chinese = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 23:
                currentRow.Czech = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 25:
                currentRow.English = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 27:
                currentRow.SerbianCyr = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 29:
                currentRow.SerbianLat = GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
            }
          }

        }

        //very important
        if (string.IsNullOrEmpty(currentRow.English))
        {
          currentRow.English = currentRow.Basic;
          currentRow.Key = RedoKey(currentRow.File, currentRow.Key);
        }

        double percentage = (double) i / (double) rowCount * 100;
        Console.WriteLine($"Row {i} out of {rowCount} ({percentage:N1})%");
        if (!fileRowViews.Select(z => z.Key).Contains(currentRow.Key))
        {
          fileRowViews.Add(currentRow);
        }
      }
      Console.WriteLine("Done reading!");
    }

    private string RedoKey(string file, string key)
    {
      var addProject = file + "_" + key;
      var changeSlashes = addProject.Replace(@"\\", "_");
      var changeSlash2 = changeSlashes.Replace("\\", "_");
      var changeDots = changeSlash2.Replace(".", "_");
      return changeDots;
    }

    private string GetProperValue(string original)
    {
      var changeHyph = original.Replace("&", "_");
      return RemoveHtml(changeHyph)?.Trim();
    }

    private string RemoveHtml(string original)
    {
      string result = original;
      try
      {
        if (original.Contains("<") && original.Contains(">"))
        {
          var htmlStart = original.IndexOf("<", StringComparison.Ordinal);
          var htmlEnd = original.IndexOf(">", StringComparison.Ordinal);
          if (htmlEnd < htmlStart)
          {
            original = original.Remove(htmlEnd, 1);
            htmlEnd = original.IndexOf(">", StringComparison.Ordinal);
          }

          var withoutHtml = original.Remove(htmlStart, htmlEnd - htmlStart + 1);

          if (withoutHtml.Contains("<") && withoutHtml.Contains(">"))
          {
            result = RemoveHtml(withoutHtml);
          }
          else
          {
            result = withoutHtml;
          }
        }
      }
      catch (Exception e)
      {
        Console.WriteLine(e);
        throw;
      }
   
      return result;
    }

    private void ReleaseUnmanagedResources()
    {
      //cleanup
      GC.Collect();
      GC.WaitForPendingFinalizers();
      Marshal.ReleaseComObject(xlRange);
      Marshal.ReleaseComObject(xlWorksheet);

      //close and release
      xlWorkbook.Close();
      Marshal.ReleaseComObject(xlWorkbook);

      //quit and release
      xclApp.Quit();
      Marshal.ReleaseComObject(xclApp);
    }

    public void SaveToNewExcelFile()
    {
      Console.WriteLine("Write the number of the document version (integer) so the doc will be created next to the old one plus the number");
      var number = int.Parse(Console.ReadLine() ?? throw new InvalidOperationException());

      Console.WriteLine("Writing to the new file!");

      var dataTable = ConvertToDataTable(fileRowViews);
      DataSet dataSet = new DataSet();
      dataSet.Tables.Add(dataTable);

      // create a excel app along side with workbook and worksheet and give a name to it  
     
      var excelWorkBook = xclApp.Workbooks.Add();
      var worksheet = excelWorkBook.Sheets[1];
      var range = worksheet.UsedRange;
      foreach (DataTable table in dataSet.Tables)
      {
        //Add a new worksheet to workbook with the Datatable name  
        var excelWorkSheet = excelWorkBook.Sheets.Add();
        excelWorkSheet.Name = table.TableName;

        // add all the columns  
        for (int i = 1; i < table.Columns.Count + 1; i++)
        {
          excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
        }

        var countRows = table.Rows.Count;
        // add all the rows  
        for (int j = 0; j < countRows; j++)
        {
          for (int k = 0; k < table.Columns.Count; k++)
          {
            excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
          }
          Console.WriteLine($"Written row {j} out of {countRows}");
        }
        
        string newLoc = docLocation.Replace(".xlsx", string.Empty) + "_" + number + ".xlsx";
        excelWorkBook.SaveAs(newLoc); // -> this will do the custom  
        excelWorkBook.Close();

        Console.WriteLine($"Saved to {newLoc}!");
        Console.ReadKey();
      }
    }


    public DataTable ConvertToDataTable<T>(IList<T> data)
    {
      PropertyDescriptorCollection properties =
        TypeDescriptor.GetProperties(typeof(T));
      DataTable table = new DataTable();
      foreach (PropertyDescriptor prop in properties)
        table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
      foreach (T item in data)
      {
        DataRow row = table.NewRow();
        foreach (PropertyDescriptor prop in properties)
          row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
        table.Rows.Add(row);
      }
      return table;
    }

    private List<ExcelFileRowView> fileRowViews = new List<ExcelFileRowView>();
    private int _countToWrite = 0;
    private const string ResDict = "<ResourceDictionary ";
    private string _lineIndent;
    public void WriteToXaml()
    {
      IEnumerable<IGrouping<string, ExcelFileRowView>> groupByFiles = fileRowViews.GroupBy(z => z.File);
      var byFiles = groupByFiles as IGrouping<string, ExcelFileRowView>[] ?? groupByFiles.ToArray();
      _countToWrite = byFiles.Count() * 13;
      counterOfWritten = 0;
      foreach (var group in byFiles)
      {
        string folder = group.FirstOrDefault()?.File ?? "unidentified";
        folder = folder.Replace("\\", "_");
        var basicGroup = group.ToDictionary(x => x.Key, y => y.Basic);
        var germanGroup = group.ToDictionary(x => x.Key, y => y.German);
        var spanishGroup = group.ToDictionary(x => x.Key, y => y.Spanish);
        var italianGroup = group.ToDictionary(x => x.Key, y => y.Italian);
        var japaneseGroup = group.ToDictionary(x => x.Key, y => y.Japanese);
        var portugueseGroup = group.ToDictionary(x => x.Key, y => y.Portuguese);
        var russianGroup = group.ToDictionary(x => x.Key, y => y.Russian);
        var turkishGroup = group.ToDictionary(x => x.Key, y => y.Turkish);
        var chineseGroup = group.ToDictionary(x => x.Key, y => y.Chinese);
        var czechGroup = group.ToDictionary(x => x.Key, y => y.Czech);
        var englishGroup = group.ToDictionary(x => x.Key, y => y.English);
        var serbianCyrGroup = group.ToDictionary(x => x.Key, y => y.SerbianCyr);
        var serbianLatGroup = group.ToDictionary(x => x.Key, y => y.SerbianLat);

        WriteToSingleFile(basicGroup, folder, string.Empty);
        WriteToSingleFile(germanGroup, folder, "de-DE");
        WriteToSingleFile(spanishGroup, folder, "es-Es");
        WriteToSingleFile(italianGroup, folder, "it-IT");
        WriteToSingleFile(japaneseGroup, folder, "ja-JP");
        WriteToSingleFile(portugueseGroup, folder, "pt-BR");
        WriteToSingleFile(russianGroup, folder, "ru-RU");
        WriteToSingleFile(turkishGroup, folder, "tr-TR");
        WriteToSingleFile(chineseGroup, folder, "zh-CN");
        WriteToSingleFile(czechGroup, folder, "cs-CZ");
        WriteToSingleFile(englishGroup, folder, "en-US");
        WriteToSingleFile(serbianCyrGroup, folder, "sr-Cyrl-RS");
        WriteToSingleFile(serbianLatGroup, folder, "sr-Latn-RS");
      }
    }

    private string randomFolder;
    private int counterOfWritten;
    private void WriteToSingleFile(Dictionary<string, string> resDictionary, string folder, string abr)
    {
      _lineIndent = indent();
      string smallIndent = _lineIndent.Remove(4);
      string rootFolder = docLocation.Remove(docLocation.LastIndexOf('\\') + 1);
   
      System.IO.Directory.CreateDirectory(rootFolder + randomFolder);

      System.IO.Directory.CreateDirectory(rootFolder + randomFolder + "\\" + folder);
      StringBuilder beginningTag = new StringBuilder();
      beginningTag.AppendLine(ResDict + "xmlns \"http://schemas.microsoft.com/winfx/2006/xaml/presentation\"");
      beginningTag.AppendLine(_lineIndent + "xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\"");
      beginningTag.AppendLine(_lineIndent + "xmlns:local=\"clr -namespace:EvidenceCenter.Ui.WPF.View.UserControls.HomeScreen\"");
      beginningTag.AppendLine(_lineIndent + "xmlns:system=\"clr -namespace:System;assembly=mscorlib\"> ");

      foreach (var res in resDictionary)
      {
        beginningTag.AppendLine(smallIndent + $"<system:String x:Key=\"{res.Key}\">{res.Value}</system:String>");
      }

      beginningTag.AppendLine("</ResourceDictionary>");
      string finalFile = rootFolder + randomFolder + "\\" + folder + "\\" + folder + abr + ".xaml";
      System.IO.File.WriteAllText(finalFile, beginningTag.ToString());
      counterOfWritten++;
      Console.WriteLine($"Created {finalFile} /n out of {_countToWrite} = {(double)counterOfWritten / (double)_countToWrite * 100}");
    }

    private string indent()
    {
      StringBuilder sb = new StringBuilder();
      foreach (var x in ResDict)
      {
        sb.Append(" ");
      }

      return sb.ToString();
    }

    public void Dispose()
    {
      ReleaseUnmanagedResources();
      GC.SuppressFinalize(this);
    }

    ~ExcelExtractor()
    {
      ReleaseUnmanagedResources();
    }

   
  }
}