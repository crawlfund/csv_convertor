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

namespace DataAnalyzer
{
    public partial class DataAnalyzerForm : Form
    {
        System.Data.DataTable dt = new System.Data.DataTable();
        public DataAnalyzerForm()
        {
            InitializeComponent();
        }

        private void ImportButton_Click(object sender, EventArgs e)
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

        private void ClearFilesButton_Click(object sender, EventArgs e)
        {
            filesListBox.Items.Clear();
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            if (dateTextBox.Text != "" && filesListBox.Items.Count != 0)
            {
                foreach (string file in filesListBox.Items)
                {
                    Console.WriteLine("Hello World");
                }
            }
            else
            {
                MessageBox.Show("Please input date and choose csv files you want to analyze");
            }
        }
       

    }
}
