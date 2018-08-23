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
using System.Reflection;

namespace ListeDePrixNovago.PDFTemplate
{
    internal enum TableType { PriceList, CatalogList};
    public class ExcelReader
    {
        private string excelFilePath;
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
        private Dictionary<string, List<string>> catalogAttributes;
        private List<CatalogItem> catalogItems;
        private List<ListItem> listItems;
        private OleDbDataReader dataReader;


        public string ExcelFilePath { get => excelFilePath; set => excelFilePath = value; }
        
        public ExcelReader(string excelFilePath)
        {
            this.ExcelFilePath = excelFilePath;
        }

        private void Reader(TableType type)
        {
            string workSheetName = GetWorksheetName(excelFilePath);
            string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties=Excel 12.0;";
            using (OleDbConnection con = new OleDbConnection(connStr))
            {
                con.Open();

                StringBuilder stbQuery = new StringBuilder();
                stbQuery.Append("SELECT * FROM [" + workSheetName + "$]");
                OleDbCommand command = new OleDbCommand(stbQuery.ToString(), con);

                dataReader = command.ExecuteReader();

                if (type == TableType.CatalogList)
                {
                    GetCatalogItem(dataReader);
                    GetAttributes();
                }
                else if (type == TableType.PriceList)
                {
                    GetListItems(dataReader);
                }
            }
        }



        public Table AddListPrice(Table t)
        {
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
            foreach(ListItem item in listItems)
            {
                t.AddRow();
                
                lastRow = (Row)t.Rows.LastObject;
                lastRow.Cells[0].AddParagraph(item.Id);
                lastRow.Cells[1].AddParagraph(item.Description);
                lastRow.Cells[2].AddParagraph(item.Price.ToString());
            }
            
            lastRow = (Row)t.Rows.LastObject;
            t.Rows.RemoveObjectAt(lastRow.Index);
            return t;
        }

        private void GetListItems(OleDbDataReader reader)
        {
            using (reader)
            {
                listItems = new List<ListItem>();
                while (reader.Read())
                {
                    listItems.Add(new ListItem()
                    {
                        Id = reader[0].ToString(),
                        Description = reader[1].ToString(),
                        Price = reader.GetDouble(2)
                    });
                }
            }
        }

        public void AddPriceCatalogTables(Section section)
        {
            Reader(TableType.CatalogList);
            section.PageSetup.PageWidth = new Unit(8.5, UnitType.Inch);
            section.PageSetup.PageHeight = new Unit(11, UnitType.Inch);
            section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;
            bool isFirstRow = true;
            bool isFirstSubRow = true;
            bool isFirstTable = true;
            foreach(KeyValuePair<string, List<string>> entry in catalogAttributes)
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
                    foreach(CatalogItem item in catalogItems.Where(a => a.Attribut2 == subAtt))
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
        private void GetAttributes()
        {
            catalogAttributes = new Dictionary<string, List<string>>();

            var req = (from a in catalogItems
                       select new
                       {
                           Att1 = a.Attribut1,
                           Att2 = a.Attribut2
                       }).Distinct();

            foreach(var item in req)
            {
                if(!catalogAttributes.ContainsKey(item.Att1))
                {
                    List<string> att = new List<string>();
                    att.Add(item.Att2);
                    catalogAttributes.Add(item.Att1, att);
                } else
                {
                    catalogAttributes.First(a => a.Key == item.Att1).Value.Add(item.Att2);
                }
            }
        }

        private void GetCatalogItem(OleDbDataReader reader)
        {
            using (reader)
            {
                catalogItems = new List<CatalogItem>();
                while (reader.Read())
                {
                    catalogItems.Add(new CatalogItem()
                    {
                        Attribut1 = reader[0].ToString(),
                        Attribut2 = reader[1].ToString(),
                        Id = reader[2].ToString(),
                        Description = reader[3].ToString(),
                        Um = reader[4].ToString(),
                        Prix1 = reader.GetDouble(5),
                        Prix2 = reader.GetDouble(6),
                        Prix3 = reader.GetDouble(7),
                        Prix4 = reader.GetDouble(8),
                        Prix5 = reader.GetDouble(9)
                    });
                }
            }
        }

        public List<string> GetPriceColumns()
        {
            //Method to return the a list of excel columns that contains "prix"
            return null;
        }

        public static string GetWorksheetName(string fileName)
        {
            Microsoft.Office.Interop.Excel.Application ExcelObj = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook theWorkbook = null;

            theWorkbook = ExcelObj.Workbooks.Open(fileName);

            Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;

            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(1);

            string worksheetName = worksheet.Name;

            theWorkbook.Close();

            ExcelObj.Quit();

            return worksheetName;
        }
        
        
    }
}
