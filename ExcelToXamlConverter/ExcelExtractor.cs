using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

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
      for (int i = 1; i <= rowCount; i++)
      {
        currentRow = new ExcelFileRowView();
        for (int j = 1; j <= colCount; j++)
        {
          if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
          {
            switch (j)
            {
              case 1:
                currentRow.Project =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 2:
                currentRow.File =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 3:
                currentRow.Key =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 5:
                currentRow.Basic =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 7:
                currentRow.German =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 9:
                currentRow.Spanish =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 11:
                currentRow.Italian =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 13:
                currentRow.Japanese =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 15:
                currentRow.Portuguese =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 17:
                currentRow.Russian =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 19:
                currentRow.Turkish =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 21:
                currentRow.Chinese =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 23:
                currentRow.Czech =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 25:
                currentRow.English =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 27:
                currentRow.SerbianCyr =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
                break;
              case 29:
                currentRow.SerbianLat =GetProperValue(xlRange.Cells[i, j].Value2.ToString());
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

        double percentage = (double)i / (double)rowCount * 100;
        Console.WriteLine($"Row {i} out of {rowCount} ({percentage:N1})%");
        fileRowViews.Add(currentRow);
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