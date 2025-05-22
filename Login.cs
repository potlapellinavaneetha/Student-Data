using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML;
using ClosedXML.Excel;


namespace StudentData
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }


            private bool CheckLogin(string filepath,string username, string password)
            {
                using (var workbook = new XLWorkbook(@"C:\Users\Navaneetha\Desktop\Login.xlsx"))
                {
                    var worksheet = workbook.Worksheet("Login");
                
                var rows = worksheet.RangeUsed().RowsUsed();

                    foreach (var row in rows.Skip(1)) // Skip header row
                    {
                        string excelUsername = row.Cell(1).GetString();
                        string excelPassword = row.Cell(2).GetString();

                        if (username == excelUsername && password == excelPassword)
                        {
                            return true;
                        }
                    }
                }

                return false;
            }
        

        private void Txt_username_TextChanged(object sender, EventArgs e)
        {

        }

       

        private void Btn_reset_Click(object sender, EventArgs e)
        {
            Txt_username.Clear();
            Password.Clear();
        }

        private void Btn_submit_Click(object sender, EventArgs e)
        {
            string filePath = "Login.xlsx"; // Ensure it's in the executable's directory or give full path
            string username = Txt_username.Text;
            string password =Password.Text;

            if (CheckLogin(filePath, username, password))
            {
                MessageBox.Show("Login successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // You can load the next form here
                MainDashBoard dashboard = new MainDashBoard();
                this.Hide();
                dashboard.Show();
            }
            else
            {
                MessageBox.Show("Invalid username or password.", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Password_TextChanged(object sender, EventArgs e)
        {

        }

        private void Btn_add_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1(); // Create instance of Form1
            form1.Show();
        }
    }
}
