using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using CsvHelper;
using ExportExcelTools;//import Export Tools
namespace DataAnalyzer
{
    public partial class DataAnalyzerForm : Form
    {
        System.Data.DataTable dt = new System.Data.DataTable();
        public DataAnalyzerForm()
        {
            InitializeComponent();
        }

        private void importButton_Click(object sender, EventArgs e)
        {
            string filelist = "";
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Title = "Choose File";
            fileDialog.Multiselect = true;
            fileDialog.Filter = "CSV File(*.csv)|*.csv";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string file in fileDialog.FileNames)
                {
                    filelist += (file + '\n');
                    filesListBox.Items.Add(file);
                }
                MessageBox.Show("Choose:\n" + filelist, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void clearFilesButton_Click(object sender, EventArgs e)
        {
            filesListBox.Items.Clear();
        }

        private void exportButton_Click(object sender, EventArgs e)
        {
            if (dateTextBox.Text != "" && filesListBox.Items.Count != 0)
            {
                foreach (string file in filesListBox.Items)
                {
                    analyzeFile(file);
                }
            }
            else
            {
                MessageBox.Show("Please input date and choose csv files you want to analyze");
            }
            
         

           
        }
        private void parseDatatable(System.Data.DataTable sourceDt, System.Data.DataTable storageDt, String condition)
        {
            try
            {
                DataRow[] rows = sourceDt.Select(condition);

                foreach (DataRow row in rows)
                {
                    storageDt.Rows.Add(row.ItemArray);
                }
                DataView dv = storageDt.DefaultView;
                dv.Sort = "title";
                dv.ToTable();
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex);
                
            }
        }


        private void analyzeFile(String filePath)
        {
            System.Data.DataTable dtTable = CsvHelper.CsvHelper.CsvParsingHelper.CsvToDataTable(filePath, true);
            if (dtTable == null)
            {
                MessageBox.Show("Please make sure the csv files isn't occupied by other programs and the files have data");
            }
            else if (!dt.Columns.Contains("title"))
            { 


                System.Data.DataTable VODACOMTable = dtTable.Clone();
                String VODACOMCondition = "audio_id = 'VODACOM'";
                parseDatatable(dtTable, VODACOMTable, VODACOMCondition);
                String[] VODACOMTitles = VODACOMTable.AsEnumerable().Select(c => (String)c["title"]).Distinct().ToArray();
                String[] VODACOMChannels = VODACOMTable.AsEnumerable().Select(c => (String)c["ChannelName"]).Distinct().ToArray();

       

                System.Data.DataTable AIRTELTable = dtTable.Clone();
                String AIRTELCondition = "audio_id = 'AIRTEL'";
                parseDatatable(dtTable, AIRTELTable, AIRTELCondition);
                String[] AIRTELTitles = AIRTELTable.AsEnumerable().Select(c => (String)c["title"]).Distinct().ToArray();
                String[] AIRTELChannels = AIRTELTable.AsEnumerable().Select(c => (String)c["ChannelName"]).Distinct().ToArray();


                System.Data.DataTable AFRICELLTable = dtTable.Clone();
                String AFRICELLCondition = "audio_id = 'AFRICELL'";
                parseDatatable(dtTable, AFRICELLTable, AFRICELLCondition);
                String[] AFRICELLTitles = AFRICELLTable.AsEnumerable().Select(c => (String)c["title"]).Distinct().ToArray();
                String[] AFRICELLChannels = AFRICELLTable.AsEnumerable().Select(c => (String)c["ChannelName"]).Distinct().ToArray();



                System.Data.DataTable ORANGETable = dtTable.Clone();
                String ORANGECondition = "audio_id = 'ORANGE'";
                parseDatatable(dtTable, ORANGETable, ORANGECondition);
                String[] ORANGETitles = ORANGETable.AsEnumerable().Select(c => (String)c["title"]).Distinct().ToArray();
                String[] ORANGEChannels = ORANGETable.AsEnumerable().Select(c => (String)c["ChannelName"]).Distinct().ToArray();



                System.Data.DataTable MARSAVCOTable = dtTable.Clone();
                String MARSAVCOCondition = "audio_id = 'MARSAVCO'";
                parseDatatable(dtTable, MARSAVCOTable, MARSAVCOCondition);
                String[] MARSAVCOTitles = MARSAVCOTable.AsEnumerable().Select(c => (String)c["title"]).Distinct().ToArray();
                String[] MARSAVCOChannels = MARSAVCOTable.AsEnumerable().Select(c => (String)c["ChannelName"]).Distinct().ToArray();


  
                //Creat an Excel including 1 workbook and 4 sheets
                ExportExcel.creatExcel();
                string date = dateTextBox.Text;
                //Fill the content into 4 different sheets
                ExportExcel.exportContent(VODACOMTable, 0, date, VODACOMTitles, VODACOMChannels);
                ExportExcel.exportContent(ORANGETable, 1, date, ORANGETitles, ORANGEChannels);
                ExportExcel.exportContent(AIRTELTable, 2, date, AIRTELTitles, AIRTELChannels);
                ExportExcel.exportContent(AFRICELLTable, 3, date, AFRICELLTitles, AFRICELLChannels);
                ExportExcel.exportContent(MARSAVCOTable, 4, date, MARSAVCOTitles, MARSAVCOChannels);

                FolderBrowserDialog FBDialog = new FolderBrowserDialog();
                if (FBDialog.ShowDialog() == DialogResult.OK)
                {
                    string path = FBDialog.SelectedPath;
                    string P_obj_excelName = "";
                    if (path.EndsWith("\\"))
                        P_obj_excelName = path + date + "_report_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";
                    else
                        P_obj_excelName = path + "\\" + date + "_report_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";
                    MessageBox.Show(P_obj_excelName);
                    //Save the excel to a fixed path
                    ExportExcel.saveExcel(P_obj_excelName);
                }

            }
            else
            {
                MessageBox.Show("Error,the csv files doesn't have title column.");
            }

        }

    }
}
