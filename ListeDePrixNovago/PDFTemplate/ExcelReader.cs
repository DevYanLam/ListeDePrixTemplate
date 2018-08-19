using MigraDoc.DocumentObjectModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using System.Data.OleDb;
using Microsoft.Office.Interop;

namespace ListeDePrixNovago.PDFTemplate
{
    public class ExcelReader
    {
        private string excelFilePath;
        private OleDbConnection con;


        public string ExcelFilePath { get => excelFilePath; set => excelFilePath = value; }
        
        public ExcelReader(string excelFilePath)
        {
            this.ExcelFilePath = excelFilePath;
        }

        private OleDbDataReader Reader()
        {
            string workSheetName = GetWorksheetName(excelFilePath);
            string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties=Excel 12.0;";
            con = new OleDbConnection(connStr);

            con.Open();

            StringBuilder stbQuery = new StringBuilder();
            stbQuery.Append("SELECT * FROM [" + workSheetName + "$]");
            OleDbCommand command = new OleDbCommand(stbQuery.ToString(), con);

            OleDbDataReader dataReader = command.ExecuteReader();

            return dataReader;
        }

        public Table ExcelTable(Table t)
        {
            
            OleDbDataReader dataReader = Reader();

            t.Borders.Bottom.Visible = true;

            double columnSize = (t.Document.DefaultPageSetup.PageWidth - (t.Document.DefaultPageSetup.RightMargin * 2)) / dataReader.FieldCount;
            
            for (int x = 0; x < dataReader.FieldCount; x++)
            {
                t.AddColumn(new Unit(columnSize, UnitType.Point));
            }

            t.AddRow();

            Row lastRow = (Row)t.Rows.LastObject;

            for (int x = 0; x < dataReader.FieldCount; x++)
            {
                lastRow.Cells[x].AddParagraph(dataReader.GetName(x));
            }


            bool isFirstRow = true;

            try
            {
                while (dataReader.Read())
                {
                    //if (isFirstRow)
                    t.AddRow();
                    for (int x = 0; x < dataReader.FieldCount; x++)
                    {
                        lastRow = (Row)t.Rows.LastObject;
                        lastRow.Cells[x].AddParagraph(dataReader[x].ToString());
                        if (x == dataReader.FieldCount - 1)
                        {
                            t.AddRow();
                        }
                    }
                    isFirstRow = false;
                }
                lastRow = (Row)t.Rows.LastObject;
                t.Rows.RemoveObjectAt(lastRow.Index);
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                dataReader.Close();
                con.Close();
            }
            
            

            return t;
        }

        public static string GetWorksheetName(string fileName)
        {
            Microsoft.Office.Interop.Excel.Application ExcelObj = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook theWorkbook = null;

            theWorkbook = ExcelObj.Workbooks.Open(fileName);

            Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;

            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(1);

            return worksheet.Name;
        }
        
        
    }
}
