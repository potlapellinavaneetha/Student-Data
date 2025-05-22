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
using ClosedXML.Excel;



namespace StudentData
{
    public partial class MainDashBoard : Form
    {
        public MainDashBoard()
        {
            InitializeComponent();
        }

      

        private void Btn_addstudentdata_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 studentForm = new Form1();
            studentForm.FormClosed += (s, args) => this.Show(); // Reopen dashboard after close
            studentForm.Show();
        }

        private void Btn_enquirylist_Click(object sender, EventArgs e)
        {
            string excelPath = @"C:\Users\Navaneetha\Desktop\Studentlist.xlsx";

            Enquirylist enquiryControl = new Enquirylist();
            enquiryControl.LoadExcelData(excelPath);
            
            PanelMain.Controls.Clear();
            enquiryControl.Dock = DockStyle.Fill;
            PanelMain.Controls.Add(enquiryControl);
        }

        private void Btn_main_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Dashboard is already active.");
        }

        private void Btn_backmain_Click(object sender, EventArgs e)
        {
            this.Close(); 
        }
    }
}
