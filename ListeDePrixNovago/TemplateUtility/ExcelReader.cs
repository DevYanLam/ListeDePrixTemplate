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
using System.Collections.ObjectModel;
using ListeDePrixNovago.Utility;

namespace ListeDePrixNovago.PDFTemplate
{
    internal enum TableType { PriceList, CatalogList};
    public class ExcelReader
    {
        private string excelFilePath;
        private OleDbConnection con;
        private double scaleRowHeight = 2;
        private List<Utility.Column> columns = new List<Utility.Column>()
        {
            new Utility.Column(){Name = "Catégorie", Size = 25 },
            new Utility.Column(){Name = "Produit", Size = 30 },
            new Utility.Column(){Name = "Description", Size = 50 },
            new Utility.Column(){Name = "U/M", Size = 10 },
            new Utility.Column(){Name = "Prix1", Size = 10 },
            new Utility.Column(){Name = "Prix2", Size = 10 },
            new Utility.Column(){Name = "Prix3", Size = 10 },
            new Utility.Column(){Name = "Prix4", Size = 10 },
            new Utility.Column(){Name = "Prix5", Size = 10 },

        };
        

        public string ExcelFilePath { get => excelFilePath; set => excelFilePath = value; }
        
        public ExcelReader(string excelFilePath)
        {
            this.ExcelFilePath = excelFilePath;
        }

        private OleDbDataReader Reader(string fieldsToSelect = "*")
        {
            string workSheetName = GetWorksheetName(excelFilePath);
            string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties=Excel 12.0;";
            con = new OleDbConnection(connStr);

            con.Open();

            StringBuilder stbQuery = new StringBuilder();
            stbQuery.Append("SELECT " + fieldsToSelect + " FROM [" + workSheetName + "$]");
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
                Paragraph para = lastRow.Cells[x].AddParagraph(dataReader.GetName(x).ToUpper());
                para.Format.Font.Bold = true;
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

        public void AddPriceCatalogTables(Section section)
        {
            section.PageSetup.PageWidth = new Unit(8.5, UnitType.Inch);
            section.PageSetup.PageHeight = new Unit(11, UnitType.Inch);
            section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;
            var att = GetAttributes();
            var catalog = GetCatalogItem();
            bool isFirstRow = true;
            bool isFirstSubRow = true;
            bool isFirstTable = true;
            foreach(KeyValuePair<string, List<string>> entry in att)
            {
                isFirstRow = true;
                section.AddTable();
                Table lastTable = section.LastTable;
                lastTable.Format.Font.Size = new Unit(10);
                double pageWidth = section.Document.DefaultPageSetup.PageWidth - (section.Document.DefaultPageSetup.RightMargin * 2);
                foreach (var column in columns)
                {
                    lastTable.AddColumn(new Unit((column.Size / 100) * pageWidth));
                }
                int i = 0;
                if (isFirstTable)
                {
                    lastTable.AddRow();
                    Row lastRow = (Row)lastTable.Rows.LastObject;
                    lastRow.Height = new Unit(lastRow.Height.Point * scaleRowHeight);
                    foreach (var column in columns)
                    {
                        lastRow.Cells[i].AddParagraph(column.Name);
                        i++;
                    }
                    lastRow.Borders.Visible = true;
                    lastRow.Shading.Color = new Color(192, 192, 192);
                    isFirstTable = false;
                }
                if(isFirstRow)
                {
                    lastTable.AddRow();
                    Row lastRow = (Row)lastTable.Rows.LastObject;
                    lastRow.Height = new Unit(lastRow.Height.Point * scaleRowHeight);
                    lastRow.Cells[0].AddParagraph(entry.Key);
                    lastRow.Cells[1].AddParagraph();
                    lastRow.Cells[2].AddParagraph();
                    lastRow.Cells[3].AddParagraph();
                    lastRow.Cells[4].AddParagraph();
                    lastRow.Cells[5].AddParagraph();
                    lastRow.Cells[6].AddParagraph();
                    lastRow.Cells[7].AddParagraph();
                    lastRow.Cells[8].AddParagraph();
                    lastRow.Shading.Color = new Color(192, 192, 192);
                    lastRow.Cells[0].MergeRight = 8;
                    lastRow.Borders.Visible = true;
                    isFirstRow = false;
                }
                foreach(string subAtt in entry.Value)
                {
                    isFirstSubRow = true;
                    foreach(CatalogItem item in catalog.Where(a => a.Attribut2 == subAtt))
                    {
                        lastTable.AddRow();
                        Row lastRow = (Row)lastTable.Rows.LastObject;
                        lastRow.Cells[0].Borders.Left.Visible = true;
                        lastRow.Height = new Unit(lastRow.Height.Point * scaleRowHeight);
                        if (isFirstSubRow)
                            lastRow.Cells[0].AddParagraph(subAtt);
                        lastRow.Cells[1].AddParagraph(item.Id);
                        lastRow.Cells[1].Borders.Right.Visible = true;
                        
                        lastRow.Cells[2].AddParagraph(item.Description);
                        lastRow.Cells[2].Borders.Right.Visible = true;
                        lastRow.Cells[3].AddParagraph(item.Um);
                        lastRow.Cells[3].Borders.Right.Visible = true;
                        lastRow.Cells[4].AddParagraph(item.Prix1.ToString());
                        lastRow.Cells[4].Borders.Right.Visible = true;
                        lastRow.Cells[5].AddParagraph(item.Prix2.ToString());
                        lastRow.Cells[5].Borders.Right.Visible = true;
                        lastRow.Cells[6].AddParagraph(item.Prix3.ToString());
                        lastRow.Cells[6].Borders.Right.Visible = true;
                        lastRow.Cells[7].AddParagraph(item.Prix4.ToString());
                        lastRow.Cells[7].Borders.Right.Visible = true;
                        lastRow.Cells[8].AddParagraph(item.Prix5.ToString());
                        lastRow.Cells[8].Borders.Right.Visible = true;
                        isFirstSubRow = false;
                    }
                    
                }
                Row lastR = (Row)lastTable.Rows.LastObject;
                lastR.Borders.Bottom.Visible = true;
                section.AddParagraph();
            }
            
        }

        private List<string> GetColumns()
        {
            List<string> columns = new List<string>();
            OleDbDataReader dataReader = Reader();
            for(int x = 0; x < dataReader.FieldCount; x++)
            {
                columns.Add(dataReader.GetName(x));
            }
            return columns;
        }

        private Dictionary<string, List<string>> GetAttributes()
        {
            Dictionary<string, List<string>> res = new Dictionary<string, List<string>>();
            OleDbDataReader dataReader = Reader("distinct attribut1, attribut2");
            while (dataReader.Read())
            {
                if (!res.ContainsKey(dataReader[0].ToString()))
                {
                    List<string> attList = new List<string>();
                    attList.Add(dataReader[1].ToString());
                    res.Add(dataReader[0].ToString(), attList);
                }
                else
                {
                    res.First(a => a.Key == dataReader[0].ToString()).Value.Add(dataReader[1].ToString());
                }
            }
            return res;
        }

        private List<CatalogItem> GetCatalogItem()
        {
            List<CatalogItem> ci = new List<CatalogItem>();
            OleDbDataReader dataReader = Reader();
            while (dataReader.Read())
            {
                ci.Add(new CatalogItem()
                {
                    Attribut1 = dataReader[0].ToString(),
                    Attribut2 = dataReader[1].ToString(),
                    Id = dataReader[2].ToString(),
                    Description = dataReader[3].ToString(),
                    Um = dataReader[4].ToString(),
                    Prix1= dataReader.GetDouble(5),
                    Prix2 = dataReader.GetDouble(6),
                    Prix3 = dataReader.GetDouble(7),
                    Prix4 = dataReader.GetDouble(8),
                    Prix5 = dataReader.GetDouble(9)
                });
            }
                return ci;
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
