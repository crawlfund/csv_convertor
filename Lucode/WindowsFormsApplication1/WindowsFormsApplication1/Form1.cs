using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using ExportExcelTools;
namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //all Sheets contain 4 sheets and their details
            var allSheets = new List<DefSheet>{
                new DefSheet{
                    id=1,
                    title="VODACOM RTCE du ",
                    color=(int)Excel.XlRgbColor.rgbSkyBlue,

                }, 
                new DefSheet{
                    id=2,
                     title="AIRTEL RTCE du ",
                    color=(int)Excel.XlRgbColor.rgbRed,
                   

                },
                new DefSheet{
                    id=3,
                    title="AFRICELL RTCE du ",
                    color=(int)Excel.XlRgbColor.rgbPurple,
                },
                new DefSheet{
                    id=4,
                    title="ORANGE RTCE du ",
                    color=(int)Excel.XlRgbColor.rgbOrange,

                },
                new DefSheet{
                    id=5,
                    title="GLOBAL DAILY REPORT VODACOM",
                    color=(int)Excel.XlRgbColor.rgbSkyBlue,
                },
                 new DefSheet{
                    id=6,
                    title="GLOBAL DAILY REPORT AIRTEL",
                    color=(int)Excel.XlRgbColor.rgbRed,
                },
                 new DefSheet{
                    id=7,
                    title="GLOBAL DAILY REPORT AFRICELL",
                    color=(int)Excel.XlRgbColor.rgbPurple,
                },
                 new DefSheet{
                    id=8,
                    title="GLOBAL DAILY REPORT ORANGE ",
                    color=(int)Excel.XlRgbColor.rgbOrange,
                }
            };
            var acrcloudItems = new List<DataItems>{
                new DataItems {
                    time=20,
                    title="KOMABASS",
                    audio_id="ORANGE",
                    duration=122,
                    type1="YOUTH",
                    type2="YOUTH2",
                    execution="Spot"
                },
                 new DataItems {
                    time=30,
                    title="KOMABASS",
                    audio_id="ORANGE",
                    duration=100,
                    type1="YOUTH",
                    type2="YOUTH2",
                    execution="Spot"
                },
                 new DataItems {
                    time=50,
                    title="KOMABASS",
                    audio_id="ORANGE",
                    duration=150,
                    type1="YOUTH",
                    type2="YOUTH2",
                    execution="Spot"
                }
            };
            var nameOfAD = new List<NameOfAD>{
                new NameOfAD{
                    id=1,
                    title="Ange Gardion "
                }, 
                new NameOfAD{
                    id=2,
                    title="FIDELITE "
                }, 
                new NameOfAD{
                    id=3,
                    title="MPDESA MORE "
                }, 
             };
            var channelItems = new List<ChannelItems>{
                new ChannelItems{
                    channelName="station1",
                    timeSpent="11:00",
                    averageExeRate=20,
                }, 
                new ChannelItems{
                    channelName="station2",
                    timeSpent="09:00",
                    averageExeRate=30,
                }, 
                new ChannelItems{
                    channelName="station3",
                    timeSpent="9:00",
                    averageExeRate=25,
                }, 
             };
            //Creat an Excel including 1 workbook and 4 sheets
            ExportExcel.creatExcel();

            //Fill the content into 4 different sheets
            string date="13 SEPT 016";
            for (int whichSheet = 0; whichSheet < 4; whichSheet++)
                ExportExcel.exportContent(acrcloudItems, allSheets, whichSheet,date);

           //Fill the daily report into 4 different sheets
           for (int whichSheet = 4; whichSheet < 8; whichSheet++)
                ExportExcel.exportReport(nameOfAD,channelItems,allSheets, whichSheet);

           //Save the excel to a fixed path
           ExportExcel.saveExcel("\\\\vmware-host\\Shared Folders\\Desktop\\csharp.xls");
          
        }
    }
       
}
