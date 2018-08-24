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
    internal enum TableType { PriceList, CatalogList, ColumnList, ListTypeList};
    public class ExcelReader
    {
        private string excelFilePath;
        private double scaleRowHeight = 2;
        private List<Utility.Column> catalogColumns = new List<Utility.Column>()
        {
            new Utility.Column(){Name = "Catégorie", Size = 25 },
            new Utility.Column(){Name = "Produit", Size = 30 },
            new Utility.Column(){Name = "Description", Size = 50 },
            new Utility.Column(){Name = "U/M", Size = 10 }

        };
        private List<Utility.Column> priceListColumns = new List<Utility.Column>()
        {
            new Utility.Column(){Name = "Produit", Size = 25 },
            new Utility.Column(){Name = "Description", Size = 50 }

        };
        private Dictionary<string, List<string>> catalogAttributes;
        private List<CatalogItem> catalogItems;
        private List<ListItem> listItems;
        private double fieldCount;
        private List<Price> columnList;
        private List<string> listTypeList;
        private OleDbDataReader dataReader;



        public ExcelReader(string excelFilePath)
        {
            this.excelFilePath = excelFilePath;
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
                else if(type == TableType.ColumnList)
                {
                    GetColumns(dataReader);
                } else if(type == TableType.ListTypeList)
                {
                    GetListType(dataReader);
                }
            }
        }
        
        public Table AddListPrice(Table t, string listTitle, List<Price> priceList)
        {
            Reader(TableType.PriceList);
            t.Borders.Bottom.Visible = true;
            double pageWidth = t.Document.DefaultPageSetup.PageWidth - (t.Document.DefaultPageSetup.RightMargin * 2);
            foreach (var column in priceListColumns)
            { 
                t.AddColumn(new Unit((column.Size / 100) * pageWidth, UnitType.Point));
            }
            foreach(var col in priceList)
            {
                t.AddColumn(new Unit(0.1 * pageWidth, UnitType.Point));
            }
            t.AddRow();
            Row lastRow = (Row)t.Rows.LastObject;
            int x = 0;
            foreach (var column in priceListColumns)
            {
                Paragraph para = lastRow.Cells[x].AddParagraph(column.Name.ToUpper());
                para.Format.Font.Bold = true;
                x++;
            }
            foreach(var col in priceList)
            {
                Paragraph para = lastRow.Cells[x].AddParagraph(col.Name.ToUpper());
                para.Format.Font.Bold = true;
                x++;
            }
            foreach(ListItem item in listItems.Where(a => a.ListType.Equals(listTitle)))
            {
                t.AddRow();
                
                lastRow = (Row)t.Rows.LastObject;
                lastRow.Cells[0].AddParagraph(item.Id);
                lastRow.Cells[1].AddParagraph(item.Description);
                int z = 2;
                foreach(var i in item.Price)
                {
                    foreach (var price in priceList)
                    {
                        if(i.Name.Equals(price.Name))
                        {
                            lastRow.Cells[z].AddParagraph(i.Amount.ToString());
                            z++;
                        }
                    }
                }
                
            }
            
            lastRow = (Row)t.Rows.LastObject;
            t.Rows.RemoveObjectAt(lastRow.Index);
            return t;
        }

        private void GetListItems(OleDbDataReader reader)
        {
            using (reader)
            {
                if (reader.GetName(0).Equals("liste"))
                {
                    fieldCount = reader.FieldCount;
                    listItems = new List<ListItem>();
                    while (reader.Read())
                    {
                        listItems.Add(new ListItem()
                        {
                            ListType = reader[0].ToString(),
                            Id = reader[1].ToString(),
                            Description = reader[2].ToString(),
                            Price = new List<Price>()
                            {
                                new Price()
                                {
                                    Name = "prix1",
                                    Amount = reader.GetDouble(3)
                                },
                                new Price()
                                {
                                    Name = "prix2",
                                    Amount = reader.GetDouble(4)
                                },
                                new Price()
                                {
                                    Name = "prix3",
                                    Amount = reader.GetDouble(5)
                                },
                                new Price()
                                {
                                    Name = "prix4",
                                    Amount = reader.GetDouble(6)
                                },
                                new Price()
                                {
                                    Name = "prix5",
                                    Amount = reader.GetDouble(7)
                                },
                                new Price()
                                {
                                    Name = "prix6",
                                    Amount = reader.GetDouble(8)
                                },
                                new Price()
                                {
                                    Name = "prix7",
                                    Amount = reader.GetDouble(9)
                                },
                                new Price()
                                {
                                    Name = "prix8",
                                    Amount = reader.GetDouble(10)
                                },
                                new Price()
                                {
                                    Name = "prix9",
                                    Amount = reader.GetDouble(11)
                                },
                                new Price()
                                {
                                    Name = "prix10",
                                    Amount = reader.GetDouble(12)
                                }
                            }

                        });
                    }
                }
                else
                    throw new Exception("Gabarit incorrect");
            }
        }

        public void AddPriceCatalogTables(Section section, List<Price> priceList)
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
                foreach (var column in catalogColumns)
                {
                    lastTable.AddColumn(new Unit((column.Size / 100) * pageWidth));
                }
                foreach(var price in priceList)
                {
                    lastTable.AddColumn(new Unit(0.1 * pageWidth));
                }
                int i = 0;
                if (isFirstTable)
                {
                    lastTable.AddRow();
                    Row lastRow = (Row)lastTable.Rows.LastObject;
                    lastRow.Height = new Unit(lastRow.Height.Point * scaleRowHeight);
                    foreach (var column in catalogColumns)
                    {
                        lastRow.Cells[i].AddParagraph(column.Name);
                        i++;
                    }
                    foreach(var price in priceList)
                    {
                        lastRow.Cells[i].AddParagraph(price.Name);
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
                    for(int q = 1; q < catalogColumns.Count; q++)
                    {
                        lastRow.Cells[q].AddParagraph();
                    }
                    int u = 4;
                    foreach(var price in priceList)
                    {
                        lastRow.Cells[u].AddParagraph();
                        u++;
                    }
                    lastRow.Shading.Color = new Color(192, 192, 192);
                    lastRow.Cells[0].MergeRight = catalogColumns.Count+priceList.Count-1;
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
                        int y = 4;
                        foreach(var p in item.Prices)
                        {
                            foreach(var price in priceList)
                            {
                                if(p.Name.Equals(price.Name))
                                {
                                    lastRow.Cells[y].AddParagraph(p.Amount.ToString());
                                    lastRow.Cells[y].Borders.Right.Visible = true;
                                    y++;
                                }
                            }
                        }

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
                if (reader.GetName(0).Equals("attribut1") && reader.GetName(1).Equals("attribut2"))
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
                            Prices = new List<Price>()
                            {
                                new Price()
                                {
                                    Name = "prix1",
                                    Amount = reader.GetDouble(5)
                                },
                                new Price()
                                {
                                    Name = "prix2",
                                    Amount = reader.GetDouble(6)
                                },
                                new Price()
                                {
                                    Name = "prix3",
                                    Amount = reader.GetDouble(7)
                                },
                                new Price()
                                {
                                    Name = "prix4",
                                    Amount = reader.GetDouble(8)
                                },
                                new Price()
                                {
                                    Name = "prix5",
                                    Amount = reader.GetDouble(9)
                                },
                                new Price()
                                {
                                    Name = "prix6",
                                    Amount = reader.GetDouble(10)
                                },
                                new Price()
                                {
                                    Name = "prix7",
                                    Amount = reader.GetDouble(11)
                                },
                                new Price()
                                {
                                    Name = "prix8",
                                    Amount = reader.GetDouble(12)
                                },
                                new Price()
                                {
                                    Name = "prix9",
                                    Amount = reader.GetDouble(13)
                                },
                                new Price()
                                {
                                    Name = "prix10",
                                    Amount = reader.GetDouble(14)
                                }
                            }
                        });
                    }
                }
                else
                    throw new Exception("Gabarit incorrect");
            }
        }

       

        public IEnumerable<Price> GetPriceColumns()
        {
            Reader(TableType.ColumnList);
            return columnList;
        }

        private void GetColumns(OleDbDataReader reader)
        {
            columnList = new List<Price>();
            using (reader)
            {
                string columnName = "";
                int x = 0;
                while(reader.Read() && reader.FieldCount > x)
                {
                    columnName = reader.GetName(x).ToString();
                    if (columnName.Contains("prix"))
                    {
                        columnList.Add(new Price()
                        {
                            Name = columnName,
                            IsChecked = false
                        }
                            );
                    }
                    x++;
                }
            }
        }

        public IEnumerable<string> GetListTypeList()
        {
            Reader(TableType.ListTypeList);
            return listTypeList;
        }

        private void GetListType(OleDbDataReader reader)
        {
            listTypeList = new List<string>();
            using (reader)
            {
                while (reader.Read())
                {
                    if (!reader.GetName(0).ToString().Equals("liste"))
                        break;
                    if(!listTypeList.Contains(reader[0].ToString()))
                    {
                        listTypeList.Add(reader[0].ToString());
                    }
                }
            }
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
