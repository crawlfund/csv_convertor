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
                    time="13 SEPT 016",
                    color=(int)Excel.XlRgbColor.rgbSkyBlue,

                }, 
                new DefSheet{
                    id=2,
                    title="ORANGE RTCE du ",
                    time="13 SEPT 016",
                    color=(int)Excel.XlRgbColor.rgbOrange,

                },
                new DefSheet{
                    id=3,
                    title="AIRTEL RTCE du ",
                    time="13 SEPT 016",
                    color=(int)Excel.XlRgbColor.rgbRed,
                },
                new DefSheet{
                    id=4,
                    title="Mputu RTCE du ",
                    time="13 SEPT 016",
                    color=(int)Excel.XlRgbColor.rgbPurple,

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

            //Creat an Excel including 1 workbook and 4 sheets
            ExportExcel.creatExcel();

            //Fill the content into 4 different sheets
            for (int whichSheet = 0; whichSheet < 4; whichSheet++)
                ExportExcel.exportContent(acrcloudItems, allSheets, whichSheet);

            //Save the excel to a fixed path
            ExportExcel.saveExcel("\\\\vmware-host\\Shared Folders\\Desktop\\csharp.xls");

        }
    }
       
}
