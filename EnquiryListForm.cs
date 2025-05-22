using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace StudentData
{
    public partial class Enquirylist : UserControl
    {
        public Enquirylist()
        {
            InitializeComponent();
        }

        private void Datagridview_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        public void LoadExcelData(string filePath)
        {
            

            var dt = new DataTable();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet("Studentlist"); // or use the sheet name: "EnquiryList"

                bool isFirstRow = true;
                foreach (var row in worksheet.RowsUsed())
                {
                    if (isFirstRow)
                    {
                        foreach (var cell in row.Cells())
                            dt.Columns.Add(cell.Value.ToString());
                        isFirstRow = false;
                    }
                    else
                    {
                        var cells = row.Cells().ToList();
                        var newRow = dt.NewRow();

                        for (int i = 0; i < Math.Min(cells.Count, dt.Columns.Count); i++)
                        {
                            var value = cells[i].Value;
                            newRow[i] = string.IsNullOrWhiteSpace(value.ToString()) ? string.Empty : value.ToString().Trim();
                        }

                        dt.Rows.Add(newRow);


                    }
                }
            }
            if (Datagridview != null)
            {
                Datagridview.DataSource = dt;
                Datagridview.ScrollBars = ScrollBars.Both;
                Datagridview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

                foreach (DataGridViewColumn col in Datagridview.Columns)
                {
                    col.Width = 100;
                }
            }
        }
            

        



        private void Panel_Paint(object sender, PaintEventArgs e)
        {

        }
    }

}





        
    
