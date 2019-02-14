using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using ClosedXML;
using CsvHelper;

namespace CSVToXLSX
{
  public static class Program
  {
    static void Main(string[] args)
    {
      if(args.Count() < 1)
      {
        Console.WriteLine("Please enter a file path:");
        args = new string[1];
        args[0] = Console.ReadLine();
      }
      var path = args[0];

      if (!File.Exists(path) || !path.ToUpper().Contains("CSV"))
      {
        Console.WriteLine("Must provide a valid CSV");
        System.Threading.Thread.Sleep(500);
        return;
      }

      List<dynamic> issues;
      using (var reader = new StreamReader(path))
      {
        using (var csv = new CsvReader(reader))
        {
          issues = csv.GetRecords<dynamic>().ToList();
        }
      }
      using (var wb = new ClosedXML.Excel.XLWorkbook())
      {
        DataTable table = ToDataTable(issues); 
        wb.AddWorksheet(table, "Sheet1");
        foreach (var ws in wb.Worksheets)
        {
          ws.Columns().AdjustToContents();
        }
        var output = path.Substring(0, path.Length-3) + "xlsx";
        wb.SaveAs(output);
        Console.WriteLine("wrote to : "+ output);
        System.Threading.Thread.Sleep(500);
        return;
      }
    }
    public static DataTable ToDataTable(IEnumerable<dynamic> items)
    {
      var data = items.ToArray();
      if (data.Count() == 0) return null;

      var dt = new DataTable();
      foreach (var key in ((IDictionary<string, object>)data[0]).Keys)
      {
        dt.Columns.Add(key);
      }
      foreach (var d in data)
      {
        dt.Rows.Add(((IDictionary<string, object>)d).Values.ToArray());
      }
      return dt;
    }
  }
}
