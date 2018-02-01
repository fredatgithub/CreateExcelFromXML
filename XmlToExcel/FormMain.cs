using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Xml;
using System.Xml.Linq;
using XmlToExcel.Properties;
using Excel = Microsoft.Office.Interop.Excel;

namespace XmlToExcel
{
  public partial class FormMain : Form
  {
    public FormMain()
    {
      InitializeComponent();
    }

    string[,] lines = new string[100, 5];

    private void button1_Click(object sender, EventArgs e)
    {
      Excel.Application xlApp;
      Excel.Workbook xlWorkBook;
      Excel.Worksheet xlWorkSheet;
      object misValue = System.Reflection.Missing.Value;

      DataSet ds = new DataSet();
      XmlReader xmlFile;
      int i = 0;
      int j = 0;

      //xlApp = new Excel.ApplicationClass();
      //xlApp = new Excel.ApplicationClass();
      //xlWorkBook = xlApp.Workbooks.Add(misValue);
      //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item[1];

      xmlFile = XmlReader.Create("Product.xml", new XmlReaderSettings());
      ds.ReadXml(xmlFile);

      for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
      {
        for (j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)
        {
          //xlWorkSheet.Cells[i + 1, j + 1] = ds.Tables[0].Rows[i].ItemArray[j].ToString();
        }
      }

      //xlWorkBook.SaveAs("xml2excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
      //xlWorkBook.Close(true, misValue, misValue);
      //xlApp.Quit();

      //releaseObject(xlApp);
      //releaseObject(xlWorkBook);
      //releaseObject(xlWorkSheet);

      MessageBox.Show("Done");
    }

    private void releaseObject(object obj)
    {
      try
      {
        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        obj = null;
      }
      catch (Exception)
      {
        obj = null;
      }
      finally
      {
        GC.Collect();
      }
    }

    private void ButtonReadXml_Click(object sender, EventArgs e)
    {
      XDocument xmlDoc;
      string xmlFileName = "sample.xml";
      try
      {
        xmlDoc = XDocument.Load(xmlFileName);
      }
      catch (Exception exception)
      {
        MessageBox.Show("error" + exception.Message);
        return;
      }

      XmlDocument doc = new XmlDocument();
      doc.Load(xmlFileName);
      XmlNodeList elemList = doc.GetElementsByTagName("Issue");

      var items = from item in xmlDoc.Descendants("Project").Elements("Issue")
                  where item.HasAttributes
                  where item.HasElements
                  select new
                  {
                    TypeId = (string)item.Attribute("TypeId"),
                    FileName = (string)item.Attribute("File"),
                    Offset = (string)item.Attribute("Offset"),
                    lineNumber = (int)item.Attribute("Line"),
                    Message = (string)item.Attribute("Message")
                  };

      var listOfIssues = items.Select(item => new Issue(item.TypeId, item.FileName, item.Offset, item.lineNumber, item.Message)).ToList();

      var numberOfIssuesfound = listOfIssues.Count;
      // foreach (var q in result)
      // {
      //   if (!_allQuotes.ListOfQuotes.Contains(new Quote(q.authorValue, q.languageValue, q.sentenceValue)) &&
      //       q.authorValue != string.Empty && q.languageValue != string.Empty && q.sentenceValue != string.Empty)
      //   {
      //     _allQuotes.Add(new Quote(q.authorValue, q.languageValue, q.sentenceValue));
      //   }
      // }

      // _allQuotes.QuoteFileSaved = true;
    }
  }
}