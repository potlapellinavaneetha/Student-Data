using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using ClosedXML.Excel;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace StudentData
{
    public partial class Form1 : Form
    {
        private readonly string _excelPath = @"C:\Users\Navaneetha\Desktop\Studentlist.xlsx";
        private const string SheetName = "Studentlist";

        public class StudentData
        {
            public string Studentname { get; set; }

            public decimal Year { get; set; }
            public string Gender { get; set; }
            public string Location { get; set; }
            public string College { get; set; }
            public string Course { get; set; }
            public string HasLaptop { get; set; }

        }
        public List<StudentData> StudentCollection = new List<StudentData>();


        public Form1()
        {
            InitializeComponent();
            Btn_submit.Click += Btn_submit_Click;
            LoadComboBoxData();
        }




        private void Cmbx_location_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cmbx_location.Focus();
        }

        private void Rbtn_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void Rbtn_femle_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Txt_collegename_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void Txt_degree_TextChanged(object sender, EventArgs e) {
           
        }
        private void Txt_name_TextChanged(object sender, EventArgs e) {
           
        }
        private void Txt_year_TextChanged(object sender, EventArgs e) {
        }


        private void Btn_reset_Click(object sender, EventArgs e)
        {
            Txt_name.Clear();
            Txt_year.Clear();
            Rbtn.Checked = true;
            Cmbx_location.SelectedIndex = 0;
            Txt_collegename.Clear();
            Txt_degree.Clear();
            Txt_laptop.Clear();
        }
        private void LoadComboBoxData()
        {
            Cmbx_location.Items.AddRange(new string[] { "Select", "Hyderabad", "Delhi", "Mumbai", "Kolkata", "Other" });

        }




        public void Btn_submit_Click(object sender, EventArgs e)
        {
            try
            {
                // Validate Name
                if (string.IsNullOrWhiteSpace(Txt_name.Text))
                {
                    MessageBox.Show("Please enter the student's name.");
                    return;
                }

                // Validate Year
                if (!decimal.TryParse(Txt_year.Text, out decimal year) || year <= 0)
                {
                    MessageBox.Show("Please enter a valid numeric year.");
                    return;
                }

                // Validate Gender
                string gender = Rbtn.Checked ? "Male" :
                                Rbtn_femle.Checked ? "Female" : "";

                if (string.IsNullOrEmpty(gender))
                {
                    MessageBox.Show("Please select a gender.");
                    return;
                }

                // Validate Location
                string location = Cmbx_location.SelectedItem?.ToString() ?? "";
                if (string.IsNullOrEmpty(location) || location == "Select")
                {
                    MessageBox.Show("Please select a location.");
                    return;
                }

                // Validate College
                if (string.IsNullOrWhiteSpace(Txt_collegename.Text))
                {
                    MessageBox.Show("Please enter the college name.");
                    return;
                }

                // Validate Course
                if (string.IsNullOrWhiteSpace(Txt_degree.Text))
                {
                    MessageBox.Show("Please enter the course/degree.");
                    return;
                }

               try
            {
                // Load or create Excel file
                XLWorkbook workbook;
                    string sheetName = "Studentlist";
                if (File.Exists(_excelPath))
                {
                    workbook = new XLWorkbook(_excelPath);
                }
                else
                {
                    workbook = new XLWorkbook();
                    workbook.AddWorksheet("Studentlist");
                }

                var ws = workbook.Worksheets.Contains(sheetName)
                    ? workbook.Worksheet(sheetName)
                    : workbook.AddWorksheet(sheetName);

                // Write headers if first time
                var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
                if (lastRow < 1)
                {
                    ws.Cell(1, 1).Value = "Name";
                    ws.Cell(1, 2).Value = "Location";
                    ws.Cell(1, 3).Value = "Degree";
                    ws.Cell(1, 4).Value = "Passout Year";
                    ws.Cell(1, 5).Value = "Has Laptop";
                    ws.Cell(1, 6).Value = "Gender";
                        ws.Cell(1, 7).Value = "College Name";
                       
                    lastRow = 1;
                }

                // Write new row
                int newRow = lastRow + 1;
                ws.Cell(newRow, 1).Value = Txt_name.Text.Trim();
                ws.Cell(newRow, 2).Value = location;
                    ws.Cell(newRow, 3).Value = Txt_degree.Text.Trim();
                    ws.Cell(newRow, 4).Value = year;
                    ws.Cell(newRow, 5).Value = Txt_laptop.Text.Trim();
                ws.Cell(newRow, 6).Value = gender;

                    ws.Cell(newRow, 7).Value =Txt_collegename.Text.Trim();

                // Save
                workbook.SaveAs(_excelPath);
                workbook.Dispose();

                MessageBox.Show("Student data saved!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error writing to Excel:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        

               

                // Add student to the list
                StudentData student = new StudentData
                {
                    Studentname = Txt_name.Text.Trim(),
                    Year = year,
                    Gender = gender,
                    Location = location,
                    College = Txt_collegename.Text.Trim(),
                    Course = Txt_degree.Text.Trim(),
                    HasLaptop = Txt_laptop.Text.Trim(),
                };

                StudentCollection.Add(student);

                // Refresh DataGridView
               // Datagw_data.DataSource = null;
                //Datagw_data.DataSource = StudentCollection;

                MessageBox.Show("Submitted successfully!");

                // Optionally reset the form
                Btn_reset_Click(sender, e);
            }
            catch (Exception ex) {
                MessageBox.Show("Runtime error: " + ex.Message);
            }
        }
        

       

        private void Txt_laptop_TextChanged(object sender, EventArgs e)
        {

        }
    }  
}