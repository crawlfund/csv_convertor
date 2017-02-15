using System;
using System.Collections;
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
        public string logo { get; set; }

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
            excelApp.Visible = false;

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
            //多余的bug测试
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


        static public double convertTimeToSecond(string timeString)
        {

            TimeSpan ts = new TimeSpan(Int32.Parse(timeString.Split(':')[0]),
                                        Int32.Parse(timeString.Split(':')[1]),
                                        Int32.Parse(timeString.Split(':')[2]));
            return ts.TotalSeconds;
        }
        static public void exportContent(System.Data.DataTable dataItems, int whichSheet,string date, String[] sheetTitles, String[] sheetChannels)
        {
            //all Sheets contain 4 sheets and their details
            var allSheets = new List<DefSheet>{
                new DefSheet{
                    title="VODACOM du ",
                    color=(UInt32)0xC07000,
                    logo=@"Vodacom.gif",

                }, 
                new DefSheet{
                    title="ORANGE du ",
                    color=(UInt32)0x317DED,
                    logo=@"Orange.jpg",

                },
                new DefSheet{
                    title="AIRTEL du ",
                    color=(UInt32)0x0000FF,
                    logo = @"Airtel.png"
                },
                new DefSheet{
                    title="AFRICELL du ",
                    color=(UInt32)0xA03070,
                    logo=@"Africell.jpg",

                },
                new DefSheet{
                    title="MARSAVCO du ",
                    color=(UInt32)0x3BA707,
                    logo=@"Marsavco.gif",
                },

                //global channels report
                new DefSheet{
                    title="GLOBAL DAILY REPORT VODACOM",
                    color= (UInt32)0xC07000,
                    logo=@"Vodacom.gif",

                },
                new DefSheet{
                    title="GLOBAL DAILY REPORT ORANGE",
                    color= (UInt32)0x317DED,
                    logo=@"Orange.jpg",
                },
                new DefSheet{
                    title="GLOBAL DAILY REPORT AIRTEL",
                    color= (UInt32)0x0000FF,
                    logo = @"Airtel.png"
                },
                new DefSheet{
                    title="GLOBAL DAILY REPORT AFRICEL",
                    color= (UInt32)0xA03070,
                    logo=@"Africell.jpg",
                },
                new DefSheet{
                    title="GLOBAL DAILY REPORT MARSAVC",
                    color= (UInt32)0x3BA707,
                    logo=@"Marsavco.gif",
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
            workSheet.Cells[1, "B"] = "Commercial";
            workSheet.Cells[1, "C"] = "Company";
            workSheet.Cells[1, "D"] = "Duration";
            workSheet.Cells[1, "E"] = "Type1";
            workSheet.Cells[1, "F"] = "Type2";
            workSheet.Cells[1, "G"] = "Execution";
            workSheet.Cells[1, "H"] = "ChannelName";

            workSheet.Shapes.AddPicture(System.IO.Directory.GetCurrentDirectory()+@"/logo/" + allSheets[whichSheet].logo, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 500, 0, 200, 60);
            workSheetGlobal.Shapes.AddPicture(System.IO.Directory.GetCurrentDirectory() + @"/logo/" + allSheets[whichSheet+5].logo, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 0, 183, 42);
            // Call to fill the color for the chart's title
            workSheet.Range["A1", "H1"].Interior.Color = allSheets[whichSheet].color;
            workSheet.Range["A1", "H1"].Font.Color = Excel.XlRgbColor.rgbWhite;

            //--------    去除无效广告 -------------------
            for (int i = 0; i < dataItems.Rows.Count; )
            {
                Console.WriteLine("************\t" + i.ToString() + "\t*************");
                string title = dataItems.Rows[i]["title"].ToString();


                string timeStr = dataItems.Rows[i]["Time"].ToString();
                double timeSeconds = convertTimeToSecond(timeStr);//time的秒数

                string durationStr = dataItems.Rows[i]["duration"].ToString();
                int durationSeconds = (int)convertTimeToSecond(durationStr);//duration的秒数

                double adEndTime = timeSeconds + durationSeconds;//广告结束时间

                string channelname1 = dataItems.Rows[i]["ChannelName"].ToString();//当前channel的值
                int timesLimit = durationSeconds / 10;
                int endPoint = 0;
                if ((i + timesLimit + 2) < dataItems.Rows.Count)
                {
                    endPoint = i + timesLimit + 2;//当前i指向的加上广告时/10的 再加个2表示误差（防止漏
                }
                else
                {
                    endPoint = dataItems.Rows.Count;//表末尾
                }
                int count = 0;//计数
                for (int j = i; j < endPoint; j++)
                {
                    string nowTitle = dataItems.Rows[j]["title"].ToString();//现在遍历到的title
                    string nowTimeStr = dataItems.Rows[j]["Time"].ToString();
                    string channelname2 = dataItems.Rows[j]["ChannelName"].ToString();//现在遍历到的channelname

                    double nowtimeSeconds = convertTimeToSecond(nowTimeStr);

                    if (nowtimeSeconds <= (adEndTime + 10) && title == nowTitle && channelname2 == channelname1)//如果channelname和title都一样count+1.
                    {
                        count++;
                    }

                }
                if ((timesLimit == 3 && count <= 1) || (timesLimit>3 && timesLimit <= 5 && count < 3) || (timesLimit > 5 && count < 4))
                {
                    Console.WriteLine("Delete----------------");
                    Console.WriteLine(dataItems.Rows[i]["Time"].ToString());
                    Console.WriteLine(dataItems.Rows[i]["title"].ToString());
                    Console.WriteLine(dataItems.Rows[i]["duration"].ToString());
                    Console.WriteLine(timesLimit);
                    Console.WriteLine(count);
                    for (int k = 0; k < count; k++)
                    {

                        dataItems.Rows[i].Delete();
                    }
                }
                else
                {
                    i += count;
                }
            



                //if ((count + 1) < timesLimit)
                //{
                //    for (int k = 0; k < count; k++)
                //    {
                //
               //         dataItems.Rows[i].Delete();
               //     }
               // }
              //  else
              //  {
              //      i += count;
              //  }
            }


            //for (int i = 0; i < invalidAdList.Count; i++)
            //{
            //    int id = (int)invalidAdList[i] - i;
                //dataItems.Rows[id].Delete();
            //}

            //------------------------------------------------------------------------------

            //-------------------  去除重复广告   --------------------------------------
            for (int i = 0; i < dataItems.Rows.Count; i++)
            {
                string title = dataItems.Rows[i]["title"].ToString();

                string timeStr = dataItems.Rows[i]["Time"].ToString();
                double timeSeconds = convertTimeToSecond(timeStr);//time的秒数

                string durationStr = dataItems.Rows[i]["duration"].ToString();
                double durationSeconds = convertTimeToSecond(durationStr);//duration的秒数

                double adEndTime = timeSeconds + durationSeconds;//广告结束时间
                string channelname1 = dataItems.Rows[i]["ChannelName"].ToString();
                for (int j = i + 1; j < dataItems.Rows.Count;)
                {
                    string nowTimeStr = dataItems.Rows[j]["Time"].ToString();
                    string nowTitle = dataItems.Rows[j]["title"].ToString();
                    string channelname2 = dataItems.Rows[j]["ChannelName"].ToString();
                    double nowtimeSeconds = convertTimeToSecond(nowTimeStr);
                    if (nowtimeSeconds <= adEndTime + 10 && title == nowTitle && channelname2 == channelname1)
                    {
                        dataItems.Rows[j].Delete();
                        j--;
                    }
                    j++;
                }
            }
            //----------------------------------------------------------------------------


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
                if (rowdata.Length != 0)
                {
                    workSheet.Cells[a, "K"] = rowdata[0]["duration"];

                    workSheet.Range["J" + a.ToString()].Interior.Color = allSheets[whichSheet].color;
                    workSheet.Range["J" + a.ToString()].Font.Color = Excel.XlRgbColor.rgbWhite;

                    workSheet.Cells[a, "L"] = "=COUNTIF(B:B,J" + a + ")";
                    workSheet.Cells[a, "M"] = "=TEXT((HOUR(K" + a.ToString() + ")*L" + a.ToString() + "*3600+MINUTE(K" + a.ToString() + ")*L" + a.ToString() + "*60+SECOND(K" + a.ToString() + ")*L" + a.ToString() + @")/24/3600,""hh::mm:ss"")";


                    workSheet.Cells[a, "O"] = "=N" + a.ToString() + "-" + "L" + a.ToString();
                    workSheet.Cells[a, "P"] = "=TEXT(ROUND(L" + a.ToString() + "/" + "N" + a.ToString() + @",2)," + "\"0.00%\"" + ")";
                    a++;
                    //这个地方是处理总表的表头
                    workSheetGlobal.Cells[4, Chr(b)] = title;
                    b++;
                }
            }
            workSheet.Cells[a, "J"] = "Total";
            workSheet.Range["J" + a.ToString(), "P" + a.ToString()].Interior.Color = allSheets[whichSheet].color;
            workSheet.Range["J" + a.ToString(),"P"+a.ToString()].Font.Color = Excel.XlRgbColor.rgbWhite;
            workSheetGlobal.Cells[4, Chr(b)] = "Total";
            workSheetGlobal.Cells[4, Chr(b + 1)] = "Time spent per channel";
            workSheetGlobal.Cells[4, Chr(b + 2)] = "Average Execution rate per channel";
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
            int d = Asc("B");
            foreach(String channelName in sheetChannels)
            {
                int cellNumber = c + 4;
                workSheetGlobal.Cells[cellNumber, "A"] = "Station " + c.ToString() + " = " + channelName;
                for ( ;d < b;d++)
                {
                    workSheetGlobal.Cells[cellNumber, Chr(d)] = "=COUNTIFS('" + allSheets[whichSheet].title + date + "'!H:H,\"" + channelName + "\",'" + allSheets[whichSheet].title + date + "'!B:B,"+Chr(d)+"4)";
                }
                workSheetGlobal.Cells[cellNumber,Chr(d)] = "=SUM(B"+cellNumber+":"+Chr(d-1)+cellNumber+")";
                workSheetGlobal.Cells[cellNumber, Chr(d + 1)] = "=TEXT(SUMIF('" + allSheets[whichSheet].title + date + "'!H:H,\"" + channelName + "\",'"+allSheets[whichSheet].title + date + "'!D:D)," + "\"[h]:mm:ss\")";
                d = Asc("B");
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
            workSheetGlobal.get_Range("A:A", System.Type.Missing).EntireColumn.ColumnWidth = 30;



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