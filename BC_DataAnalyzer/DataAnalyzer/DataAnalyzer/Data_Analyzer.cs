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
                    readFile(file);
                }
            }
            else
            {
                MessageBox.Show("Please input date and choose csv files you want to analyze");
            }
        }
        private void parseDatatable(System.Data.DataTable sourceDt, System.Data.DataTable storageDt, String condition)
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

        private void readFile(String filePath)
        {
            Console.WriteLine(filePath);
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


                System.Data.DataTable AIRTELTable = dtTable.Clone();
                String AIRTELCondition = "audio_id = 'AIRTEL'";
                parseDatatable(dtTable, AIRTELTable, AIRTELCondition);

                System.Data.DataTable AFRICELLTable = dtTable.Clone();
                String AFRICELLCondition = "audio_id = 'AFRICELL'";
                parseDatatable(dtTable, AFRICELLTable, AFRICELLCondition);

                System.Data.DataTable MARSAVCOTable = dtTable.Clone();
                String MARSAVCOCondition = "audio_id = 'MARSAVCO'";
                parseDatatable(dtTable, MARSAVCOTable, MARSAVCOCondition);

                dataGridView1.DataSource = MARSAVCOTable;

            }
            else
            {
                MessageBox.Show("Error,the csv files doesn't have title column.");
            }

        }
    }
}
