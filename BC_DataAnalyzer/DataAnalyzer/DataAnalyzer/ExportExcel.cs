using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;
namespace ExportExcelTools
{
    public class DataItems
    {
        public int time { get; set; }
        public string title { get; set; }
        public string audio_id { get; set; }
        public int duration { get; set; }
        public string type1 { get; set; }
        public string type2 { get; set; }
        public string execution { get; set; }
    }
    public class DefSheet
    {
        public string title { get; set; }
        public UInt32 color { get; set; }
    }
    public class ExportExcel
    {
        static private Excel.Application excelApp;
        static private Excel.Workbook workBook;
        static private object misValue;
  
        static public void creatExcel()
        {

            misValue = System.Reflection.Missing.Value;
            excelApp = new Excel.Application();
            if (excelApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }
            excelApp.DisplayAlerts = false;
            // Make the object visible.
            excelApp.Visible = true;

            // Create a new, empty workbook and add it to the collection returned 
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a praticular template. 
            // Because no argument is sent in this example, Add creates a new workbook. 
            workBook = excelApp.Workbooks.Add();
            
            //The default number of sheets is three so we need to add another one
            workBook.Worksheets.Add();
            workBook.Worksheets.Add();
        }
        static public void saveExcel(string filePath)
        {
            workBook.SaveAs(filePath, AccessMode: Excel.XlSaveAsAccessMode.xlShared);
            //workBook.SaveAs("c:\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            workBook.Close(true,misValue,misValue);
            excelApp.Quit();
        }
        static public void exportContent(System.Data.DataTable dataItems, int whichSheet,string date)
        {
            //all Sheets contain 4 sheets and their details
            var allSheets = new List<DefSheet>{
                new DefSheet{
                    title="VODACOM RTCE du ",
                    color=(UInt32)0xC07000,

                }, 
                new DefSheet{
                    title="ORANGE RTCE du ",
                    color=(UInt32)0x317DED,

                },
                new DefSheet{
                    title="AIRTEL RTCE du ",
                    color=(UInt32)0x0000FF,
                },
                new DefSheet{
                    title="AFRICELL RTCE du ",
                    color=(UInt32)0xA03070,

                },
                new DefSheet{
                    title="MARSAVCO RTCE du ",
                    color=(UInt32)0x3BA707,

                }
            };
            // Choose to the second workSheet, which is the blue one
            Excel._Worksheet workSheet = workBook.Worksheets.get_Item(whichSheet+1);

            // The name for the worksheet
            workSheet.Name = allSheets[whichSheet].title + date;
            workSheet.Tab.Color = allSheets[whichSheet].color;



            // Establish column headings in cells A1 and B1.
            workSheet.Cells[1, "A"] = "Time";
            workSheet.Cells[1, "B"] = "Title";
            workSheet.Cells[1, "C"] = "Audio_ID";
            workSheet.Cells[1, "D"] = "Duration";
            workSheet.Cells[1, "E"] = "Type1";
            workSheet.Cells[1, "F"] = "Type2";
            workSheet.Cells[1, "G"] = "Execution";
          
            // Call to fill the color for the chart's title
            workSheet.Range["A1", "G1"].Interior.Color = allSheets[whichSheet].color;
            workSheet.Range["A1", "G1"].Font.Color = Excel.XlRgbColor.rgbWhite;




            var row = 1;
            foreach (System.Data.DataRow item in dataItems.Rows)
            {
                row++;
                workSheet.Cells[row, "A"] = item["Time"];
                workSheet.Cells[row, "B"] = item["title"];
                workSheet.Cells[row, "C"] = item["audio_id"];
                workSheet.Cells[row, "D"] = item["duration"];
                workSheet.Cells[row, "E"] = item["TYPE1"];
                workSheet.Cells[row, "F"] = item["TYPE2"];
                workSheet.Cells[row, "G"] = item["EXECUTION"];
            }
            //I don't know this ...too lazy to search.
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();

            //Add the summary

            workSheet.Cells[7, "J"] = "Quantitative daily report ";
            workSheet.Cells[8, "L"] = "Broadcasted 	 ";
            workSheet.Cells[8, "N"]="Ordered";
            workSheet.Cells[9, "J"] = "Commercials ";
            workSheet.Cells[9, "K"] = "Duration ";
            workSheet.Cells[9, "L"] = "Qty";
            workSheet.Cells[9, "M"] = "Time spent";
            workSheet.Cells[9, "N"] = "";
            workSheet.Cells[9, "O"] = "Variation";
            workSheet.Cells[9, "P"] = "Execution rate ";
            workSheet.Cells[10, "J"] = "Forfait Int  Pay / Mass / M-MONEY / Spot";
            workSheet.Cells[11, "J"] = "Mputu  JS8 / Spot";
            workSheet.Cells[12, "J"] = "Tarif 3G / YOUTH / SAV / Spot";
            workSheet.Cells[13, "J"] = "Total";

            // Call to fill the color for the chart's title
            workSheet.Range["J7"].Interior.Color = allSheets[whichSheet].color;
            workSheet.Range["J7"].Font.Color = Excel.XlRgbColor.rgbWhite;
            workSheet.Range["L8", "N8"].Interior.Color = allSheets[whichSheet].color;
            workSheet.Range["L8", "N8"].Font.Color = Excel.XlRgbColor.rgbWhite;
            workSheet.Range["J9", "P9"].Interior.Color = allSheets[whichSheet].color;
            workSheet.Range["J9", "P9"].Font.Color = Excel.XlRgbColor.rgbWhite;
            workSheet.Range["J10", "J13"].Interior.Color = allSheets[whichSheet].color;
            workSheet.Range["J10", "J13"].Font.Color = Excel.XlRgbColor.rgbWhite;
            workSheet.Range["K13", "P13"].Interior.Color = allSheets[whichSheet].color;
            workSheet.Range["K13", "P13"].Font.Color = Excel.XlRgbColor.rgbWhite;

            workSheet.UsedRange.Font.Name = "dengxian";//设置字体
            workSheet.UsedRange.Font.Size = 11;//设置字体大小
            workSheet.Columns.AutoFit();//单元格高度宽度自动

            int lastRowNumber = workSheet.UsedRange.Rows.Count;

        
        }

    }
}