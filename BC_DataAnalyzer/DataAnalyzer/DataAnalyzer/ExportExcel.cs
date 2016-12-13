using System;
using System.Collections.Generic;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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

            workBook.Worksheets.Add();
            workBook.Worksheets.Add();
            workBook.Worksheets.Add();
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
        static public void exportContent(System.Data.DataTable dataItems, int whichSheet,string date, String[] sheetTitles, String[] sheetChannels)
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
                },

                //global channels report
                new DefSheet{
                    title="GLOBAL DAILY REPORT VODACOM",
                    color= (UInt32)0xC07000,
                },
                new DefSheet{
                    title="GLOBAL DAILY REPORT ORANGE",
                    color= (UInt32)0x317DED,
                },
                new DefSheet{
                    title="GLOBAL DAILY REPORT AIRTEL",
                    color= (UInt32)0x0000FF,
                },
                new DefSheet{
                    title="GLOBAL DAILY REPORT AFRICEL",
                    color= (UInt32)0xA03070,
                },
                new DefSheet{
                    title="GLOBAL DAILY REPORT MARSAVC",
                    color= (UInt32)0x3BA707,
                },

            };
            Excel._Worksheet workSheet = workBook.Worksheets.get_Item(whichSheet+1);
            Excel._Worksheet workSheetGlobal = workBook.Worksheets.get_Item(whichSheet + 6);
            // The name for the worksheet
            workSheet.Name = allSheets[whichSheet].title + date;
            workSheet.Tab.Color = allSheets[whichSheet].color;

            workSheetGlobal.Name = allSheets[whichSheet+5].title;
            workSheetGlobal.Tab.Color = allSheets[whichSheet+5].color;




            // Establish column headings in cells A1 and B1.
            workSheet.Cells[1, "A"] = "Time";
            workSheet.Cells[1, "B"] = "Title";
            workSheet.Cells[1, "C"] = "Audio_ID";
            workSheet.Cells[1, "D"] = "Duration";
            workSheet.Cells[1, "E"] = "Type1";
            workSheet.Cells[1, "F"] = "Type2";
            workSheet.Cells[1, "G"] = "Execution";
            workSheet.Cells[1, "H"] = "ChannelName";
          
            // Call to fill the color for the chart's title
            workSheet.Range["A1", "H1"].Interior.Color = allSheets[whichSheet].color;
            workSheet.Range["A1", "H1"].Font.Color = Excel.XlRgbColor.rgbWhite;



            //for (int i = 0; i < dataItems.Rows.Count; i++)
            //{
            //    string title = dataItems.Rows[i]["title"].ToString();

            //    string timeStr = dataItems.Rows[i]["Time"].ToString();
            //    double timeSeconds = TimeSpan.Parse(timeStr).TotalSeconds;//time的秒数

            //    string durationStr = dataItems.Rows[i]["duration"].ToString();
            //    double durationSeconds = TimeSpan.Parse(durationStr).TotalSeconds;//duration的秒数

            //    double adEndTime = timeSeconds + durationSeconds;//广告结束时间
            //    string channelname1 = dataItems.Rows[i]["ChannelName"].ToString();
            //    //去除重复广告
            //    for (int j = i + 1; j < dataItems.Rows.Count;)
            //    {
            //        string nowTimeStr = dataItems.Rows[j]["Time"].ToString();
            //        string nowTitle = dataItems.Rows[j]["title"].ToString();
            //        string channelname2 = dataItems.Rows[j]["ChannelName"].ToString();
            //        double nowtimeSeconds = TimeSpan.Parse(nowTimeStr).TotalSeconds;
            //        if (nowtimeSeconds <= adEndTime + 1 && title == nowTitle)
            //        {
            //            dataItems.Rows[j].Delete();
            //            j--;
            //        }
            //        j++;
            //    }
            //}



              //这个地方是输出报表主体的
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
                workSheet.Cells[row, "H"] = item["ChannelName"];

            }



             //这个地方是小报表的字段

            workSheet.Cells[7, "J"] = "Quantitative daily report";
            workSheet.Cells[8, "L"] = "Broadcasted";
            workSheet.Cells[8, "N"]="Ordered";
            workSheet.Cells[9, "J"] = "Commercials";
            workSheet.Cells[9, "K"] = "Duration";
            workSheet.Cells[9, "L"] = "Qty";
            workSheet.Cells[9, "M"] = "Time spent";
            workSheet.Cells[9, "N"] = "";
            workSheet.Cells[9, "O"] = "Variation";
            workSheet.Cells[9, "P"] = "Execution rate";



            
            int a = 10;
            int b = Asc("B");
            foreach(String title in sheetTitles)
            {
                //这个地方是处理小报表的格式和内容
                workSheet.Cells[a, "J"] = title;
                
                DataRow[] rowdata = dataItems.Select("title = '"+title+"'");
                workSheet.Cells[a, "K"] = rowdata[0]["duration"];

                workSheet.Range["J" + a.ToString()].Interior.Color = allSheets[whichSheet].color;
                workSheet.Range["J" + a.ToString()].Font.Color = Excel.XlRgbColor.rgbWhite;

                workSheet.Cells[a, "L"] = "=COUNTIF(B:B,J" + a + ")";
                workSheet.Cells[a, "M"] = "=TEXT((HOUR(K"+ a.ToString()+")*L"+ a.ToString()+"*3600+MINUTE(K"+ a.ToString()+")*L"+ a.ToString()+"*60+SECOND(K"+ a.ToString()+")*L"+ a.ToString()+@")/24/3600,""hh::mm:ss"")";


                workSheet.Cells[a, "O"] = "=N" + a.ToString() + "-" + "L" + a.ToString();
                workSheet.Cells[a, "P"] = "=L" + a.ToString() + "/" + "N" + a.ToString();
                a++;
                //这个地方是处理总表的表头
                workSheetGlobal.Cells[4, Chr(b)] = title;
                b++;
               
            }
            workSheet.Cells[a, "J"] = "Total";
            workSheet.Range["J" + a.ToString(), "P" + a.ToString()].Interior.Color = allSheets[whichSheet].color;
            workSheet.Range["J" + a.ToString(),"P"+a.ToString()].Font.Color = Excel.XlRgbColor.rgbWhite;
            workSheetGlobal.Cells[4, Chr(b)] = "Total";
            workSheetGlobal.Cells[4, Chr(b + 1)] = "Time spent per channel";
            workSheetGlobal.Cells[4, Chr(b + 2)] = "Average Excution rate per channel";
            workSheetGlobal.Range[Chr(b + 1) + "4", Chr(b + 2) + "4"].Interior.Color = allSheets[whichSheet].color;
            workSheetGlobal.Range[Chr(b + 1) + "4", Chr(b + 2) + "4"].Font.Color = Excel.XlRgbColor.rgbWhite;


            //创建总表头

            workSheetGlobal.Range["B3", Chr(b+1)+"3"].Interior.Color = allSheets[whichSheet + 5].color;
            workSheetGlobal.Range["B3", Chr(b+1)+"3"].Font.Color = Excel.XlRgbColor.rgbWhite;
            workSheetGlobal.Range["B3", Chr(b+1)+"3"].Merge();//合并单元格
            workSheetGlobal.Cells["3", "B"] = "Broadcasted";
            workSheetGlobal.Range["B3", Chr(b+1) + "3"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            




            // Call to fill the color for the chart's title
            workSheet.Range["J7"].Interior.Color = allSheets[whichSheet].color;
            workSheet.Range["J7"].Font.Color = Excel.XlRgbColor.rgbWhite;
            workSheet.Range["L8", "N8"].Interior.Color = allSheets[whichSheet].color;
            workSheet.Range["L8", "N8"].Font.Color = Excel.XlRgbColor.rgbWhite;
            workSheet.Range["J9", "P9"].Interior.Color = allSheets[whichSheet].color;
            workSheet.Range["J9", "P9"].Font.Color = Excel.XlRgbColor.rgbWhite;










            //下面是global报表



            int c = 1;
            foreach(String channelName in sheetChannels)
            {
                int cellNumber = c + 4;
                workSheetGlobal.Cells[cellNumber, "A"] = "Station " + c.ToString() + " = " + channelName;
                for (int i = Asc("B"); i < b; i++)
                {
                    workSheetGlobal.Cells[cellNumber, Chr(i)] = "=COUNTIFS('" + allSheets[whichSheet].title + date + "'!H:H,\"" + channelName + "\",'" + allSheets[whichSheet].title + date + "'!B:B,B" + (cellNumber - 1) + ")";
                }
                c++;
            }

            workSheetGlobal.Cells[c + 4, "A"] = "TOTAL";
            workSheetGlobal.Range["A" + (c + 4).ToString(), Chr(b + 2).ToString() + (c + 4).ToString()].Interior.Color = allSheets[whichSheet].color;
            workSheetGlobal.Range["A" + (c + 4).ToString(), Chr(b + 2).ToString() + (c + 4).ToString()].Font.Color = Excel.XlRgbColor.rgbWhite;
            for (int i = Asc("B"); i < b + 3; i++)
            {
                workSheetGlobal.Cells[c+4,Chr(i)]="=SUM("+Chr(i)+"5:"+Chr(i)+(c+3)+")";
            }






                workSheet.UsedRange.Font.Name = "dengxian";//设置字体
            workSheet.UsedRange.Font.Size = 11;//设置字体大小
            workSheet.Columns.AutoFit();//单元格高度宽度自动


            workSheetGlobal.UsedRange.Font.Name = "dengxian";//设置字体
            workSheetGlobal.UsedRange.Font.Size = 11;//设置字体大小
            workSheetGlobal.Columns.AutoFit();//单元格高度宽度自动

            int lastRowNumber = workSheet.UsedRange.Rows.Count;
        
        }

        private static int Asc(string character)
        {
            if (character.Length == 1)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                int intAsciiCode = (int)asciiEncoding.GetBytes(character)[0];
                return (intAsciiCode);
            }
            else
            {
                throw new Exception("Character is not valid.");
            }

        }
        private static string Chr(int asciiCode)
        {
            if (asciiCode >= 0 && asciiCode <= 255)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] byteArray = new byte[] { (byte)asciiCode };
                string strCharacter = asciiEncoding.GetString(byteArray);
                return (strCharacter);
            }
            else
            {
                throw new Exception("ASCII Code is not valid.");
            }
        }





    }
}