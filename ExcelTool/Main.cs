using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTool
{
    public partial class Main : Form
    {
        string path;
        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            comboBox1.Visible = false;
            comboBox2.Visible = false;
        }

        private void Main_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void Main_DragDrop(object sender, DragEventArgs e)
        {
            path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            label1.Text = path;
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox1.Visible = true;
            comboBox2.Visible = true;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(comboBox1.SelectedIndex==0)
            {
                ExcelUtility.ConvertToJson(path);
            }

            if(comboBox1.SelectedIndex == 1)
            {
                ExcelUtility.ConvertToXml(path);
            }

            if (comboBox1.SelectedIndex == 2)
            {
                ExcelUtility.ConvertToCSV(path);
            }
        }
    }
}
