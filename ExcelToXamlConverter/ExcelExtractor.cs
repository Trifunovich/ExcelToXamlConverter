using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Text;
using System.Xaml;
using System.Windows.Markup;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using XamlReader = System.Windows.Markup.XamlReader;
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

    private void ReadFileLocation(string fileType)
    {
      Console.WriteLine($"Location of your {fileType}");
      docLocation = Console.ReadLine();
    }

    public void ReadFile()
    {
      try
      {
        ReadFileLocation(".xlsx document");
        xlWorkbook = xclApp.Workbooks.Open(docLocation);
        if (docLocation == null)
        {
          ReadFile();
        }
        var rootFolder = docLocation.Contains('\\') ? docLocation?.Remove(docLocation.LastIndexOf('\\') + 1) : docLocation;
        rootRandomFolder = rootFolder + randomFolder;
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

      StringBuilder basicBuilder = new StringBuilder();
      StringBuilder germanBuilder = new StringBuilder();
      StringBuilder spanishBuilder = new StringBuilder();
      StringBuilder italianBuilder = new StringBuilder();
      StringBuilder japaneseBuilder = new StringBuilder();
      StringBuilder portugueseBuilder = new StringBuilder();
      StringBuilder russianBuilder = new StringBuilder();
      StringBuilder turkishBuilder = new StringBuilder();
      StringBuilder chineseBuilder = new StringBuilder();
      StringBuilder czechBuilder = new StringBuilder();
      StringBuilder englishBuilder = new StringBuilder();
      StringBuilder serbianCyrBuilder = new StringBuilder();
      StringBuilder serbianLatnBuilder = new StringBuilder();

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

        AddToMasterStringBuilder(basicBuilder, WriteToSingleFile(basicGroup, folder, string.Empty));
        AddToMasterStringBuilder(germanBuilder, WriteToSingleFile(germanGroup, folder, ".de-DE"));
        AddToMasterStringBuilder(spanishBuilder, WriteToSingleFile(spanishGroup, folder, ".es-Es"));
        AddToMasterStringBuilder(italianBuilder, WriteToSingleFile(italianGroup, folder, ".it-IT"));
        AddToMasterStringBuilder(japaneseBuilder, WriteToSingleFile(japaneseGroup, folder, ".ja-JP"));
        AddToMasterStringBuilder(portugueseBuilder, WriteToSingleFile(portugueseGroup, folder, ".pt-BR"));
        AddToMasterStringBuilder(russianBuilder, WriteToSingleFile(russianGroup, folder, ".ru-RU"));
        AddToMasterStringBuilder(turkishBuilder, WriteToSingleFile(turkishGroup, folder, ".tr-TR"));
        AddToMasterStringBuilder(chineseBuilder, WriteToSingleFile(chineseGroup, folder, ".zh-CN"));
        AddToMasterStringBuilder(czechBuilder, WriteToSingleFile(czechGroup, folder, ".cs-CZ"));
        AddToMasterStringBuilder(englishBuilder, WriteToSingleFile(englishGroup, folder, ".en-US"));
        AddToMasterStringBuilder(serbianCyrBuilder, WriteToSingleFile(serbianCyrGroup, folder, ".sr-Cyrl-RS"));
        AddToMasterStringBuilder(serbianLatnBuilder, WriteToSingleFile(serbianLatGroup, folder, ".sr-Latn-RS"));
      }

      WriteToMasterFile(basicBuilder,string.Empty);
      WriteToMasterFile(germanBuilder, ".de-DE");
      WriteToMasterFile(spanishBuilder,".es-Es");
      WriteToMasterFile(italianBuilder, ".it-IT");
      WriteToMasterFile(japaneseBuilder, ".ja-JP");
      WriteToMasterFile(portugueseBuilder, ".pt-BR");
      WriteToMasterFile(russianBuilder, ".ru-RU");
      WriteToMasterFile(turkishBuilder, ".tr-TR");
      WriteToMasterFile(chineseBuilder, ".zh-CN");
      WriteToMasterFile(czechBuilder, ".cs-CZ");
      WriteToMasterFile(englishBuilder, ".en-US");
      WriteToMasterFile(serbianCyrBuilder, ".sr-Cyrl-RS");
      WriteToMasterFile(serbianLatnBuilder, ".sr-Latn-RS");
    }

    private string randomFolder;
    private string rootRandomFolder;
    private int counterOfWritten;
    private Tuple<string, string> WriteToSingleFile(Dictionary<string, string> resDictionary, string folder, string abr)
    {
      _lineIndent = indent();
      string smallIndent = _lineIndent.Remove(4);
   
      System.IO.Directory.CreateDirectory(rootRandomFolder);

      System.IO.Directory.CreateDirectory(rootRandomFolder + "\\" + folder);
      StringBuilder beginningTag = new StringBuilder();
      beginningTag.AppendLine(ResDict + "xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\"");
      beginningTag.AppendLine(_lineIndent + "xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\"");
      beginningTag.AppendLine(_lineIndent + "xmlns:system=\"clr-namespace:System;assembly=mscorlib\"> ");

      foreach (var res in resDictionary)
      {
        if (!string.IsNullOrEmpty(res.Value))
        {
          beginningTag.AppendLine(smallIndent + $"<system:String x:Key=\"{res.Key}\">{res.Value}</system:String>");
        }
      }

      beginningTag.AppendLine("</ResourceDictionary>");
      string finalFile = rootRandomFolder + "\\" + folder + "\\" + folder + abr + ".xaml";
      System.IO.File.WriteAllText(finalFile, beginningTag.ToString());
      counterOfWritten++;
      Console.WriteLine($"Created {finalFile} /n out of {_countToWrite} = {(double)counterOfWritten / (double)_countToWrite * 100}");
      return new Tuple<string, string>(folder, folder + abr + ".xaml");
    }

    private void WriteToMasterFile(StringBuilder files, string abr)
    {
      string filePath = rootRandomFolder + "\\Locale" + abr + ".xaml";
      StringBuilder sb = new StringBuilder();
      sb.AppendLine("<ResourceDictionary xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\">");
      sb.AppendLine("    <ResourceDictionary.MergedDictionaries>");
      sb.AppendLine(files.ToString());
      sb.AppendLine("    </ResourceDictionary.MergedDictionaries>");
      sb.AppendLine("</ResourceDictionary>");
      System.IO.File.WriteAllText(filePath, sb.ToString());
      Console.WriteLine($"Created master file {filePath}");
    }

    private void AddToMasterStringBuilder(StringBuilder sb, Tuple<string, string> folderName)
    {
      sb.AppendLine($"<ResourceDictionary Source=\"/Belkasoft.Ui.WPF;component/Localization/{folderName.Item1}/{folderName.Item2}\"/>");
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

    private List<string> _readXamlLines = new List<string>();
    private List<ExcelFileRowView> _rowsToExport = new List<ExcelFileRowView>();
    public void ReadXaml()
    {
      ReadFileLocation("whole folder where your .xaml files are");
      foreach (string readFile in Directory.EnumerateFiles(docLocation, "*.xaml"))
      {
        ReadOneFile(readFile);
      }

      DataTable dt = ConvertToDataTable(_rowsToExport);
      ExportToExcel(dt);
    }

    private void ReadOneFile(string doc)
    {
      using (FileStream fsSource = new FileStream(doc,
        FileMode.Open, FileAccess.Read))
      {
        using (var streamReader = new StreamReader(fsSource, Encoding.UTF8))
        {
          while (!streamReader.EndOfStream)
          {
            string oneLine = streamReader.ReadLine();
            if (!badLine(oneLine))
            {
              var kv = ExtractKeyValue(oneLine);
              UpdateField(doc, kv.Key, kv.Value);
            }
          }

        }
      }

    }

    private KeyValuePair<string, string> ExtractKeyValue (string line)
    {
      line = line.Trim();
      var twoParts = line.Split(new string[] { "\">" }, StringSplitOptions.None); ;
      string firstPart = twoParts[0].Split('"')[1].Replace("<system:String x:Key=\\\'", string.Empty);
      string secondPart = twoParts[1].Replace("</system:String>", string.Empty);
      return new KeyValuePair<string, string>(firstPart, secondPart);
    }
    
    
    private void UpdateField(string doc, string key, string value)
    {
      ExcelFileRowView row = _rowsToExport.FirstOrDefault(z => z.Key.Equals(key));
      if (row == null)
      {
        row = new ExcelFileRowView();
        row.Key = key;
        _rowsToExport.Add(row);
      }

      if (doc.Contains(".de-DE"))
      {
        row.German = value;
      }
      else if (doc.Contains(".es-Es"))
      {
        row.Spanish = value;
      }
      else if (doc.Contains(".it-IT"))
      {
        row.Italian = value;
      }
      else if (doc.Contains(".ja-JP"))
      {
        row.Japanese = value;
      }
      else if (doc.Contains(".pt-BR"))
      {
        row.Portuguese = value;
      }
      else if (doc.Contains(".ru-RU"))
      {
        row.Russian = value;
      }
      else if (doc.Contains(".tr-TR"))
      {
        row.Turkish = value;
      }
      else if (doc.Contains(".zh-CN"))
      {
        row.Chinese = value;
      }
      else if (doc.Contains(".cs-CZ"))
      {
        row.Czech = value;
      }
      else if (doc.Contains(".en-US"))
      {
        row.English = value;
      }
      else if (doc.Contains(".sr-Cyrl-RS"))
      {
        row.SerbianCyr = value;
      }
      else if (doc.Contains(".sr-Latn-RS"))
      {
        row.SerbianLat = value;
      }
      else
      {
        row.Basic = value;
      }
    }

    private bool badLine(string line)
    {
      return !line.Contains(@"x:Key");
    }

    private void ExportToExcel(DataTable dataTable)
    {
      var path = docLocation + "/TranslationOutput.xlsx";
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
        
        excelWorkBook.SaveAs(path); // -> this will do the custom  
        excelWorkBook.Close();

        Console.WriteLine($"Saved to {path}!");
        Console.ReadKey();
      }
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