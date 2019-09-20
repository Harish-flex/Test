using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
namespace WindowsFormsApplication3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private OpenFileDialog openFileDialog1;
        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1 = new OpenFileDialog();
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            string file = openFileDialog1.FileName;
            Console.WriteLine(file); // <-- For debugging use.
            textBox1.Text = file;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1 = new OpenFileDialog();
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            string file = openFileDialog1.FileName;
            Console.WriteLine(file); // <-- For debugging use.
            textBox2.Text = file;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Application xlApp1 = new Excel.Application();
            Excel.Workbook xlWorkBook;
            string userName = Environment.UserName;
            xlWorkBook = xlApp.Workbooks.Open("C:\\Users\\" + userName + "\\Desktop\\BOM_vs_Alternate.xlsm");
            xlApp.Visible = false;
            var xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
            xlWorkSheet.Cells[1, 1].value = textBox1.Text;
            xlWorkSheet.Cells[2, 1].value = textBox2.Text;
            xlApp.Run("'BOM_vs_Alternate.xlsm'!BOM_vs_Alternate.BOM_vs_Alternate");
            string outp = xlWorkSheet.Cells[3, 1].value;
            string otpath= xlWorkSheet.Cells[4, 1].value;
            Console.WriteLine(outp);
            textBox1.Text = "";
            textBox2.Text = "";
            string message = "Process Completed\n\nOutput Msg: "+ outp;
            string title = "Compeleted";
            MessageBox.Show(message, title);
            xlWorkBook.Close(false);
            xlApp.Quit();
            releaseObject(xlApp);
            releaseObject(xlWorkBook);
            xlWorkBook = xlApp1.Workbooks.Open("C:\\Automation\\Output\\" + otpath);
            xlApp1.Visible = true;
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
        }

        private void label2_Click(object sender, EventArgs e)
        {
            
    }

        private void Form1_Load(object sender, EventArgs e)
        {
           
            string name = System.DirectoryServices.AccountManagement.UserPrincipal.Current.DisplayName;
            Console.WriteLine(name);
            label2.Text = "Welcome " + name;
            linkLabel1.Text= @"C:\Automation\Output\";


        }
        public enum EXTENDED_NAME_FORMAT
        {
            NameUnknown = 0,
            NameFullyQualifiedDN = 1,
            NameSamCompatible = 2,
            NameDisplay = 3,
            NameUniqueId = 6,
            NameCanonical = 7,
            NameUserPrincipal = 8,
            NameCanonicalEx = 9,
            NameServicePrincipal = 10,
            NameDnsDomain = 12
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", @"C:\Automation\Output");
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("a");
        }
    }
    }
