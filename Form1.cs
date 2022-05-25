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
using System.Runtime.InteropServices;
using System.IO;

namespace ExcelApplication
{
    public partial class Form1 : Form

    {
        DataSet ds = new DataSet();
        DataTable datat = new DataTable();
    
        public Form1()
        {
            InitializeComponent();
        }
 
        private void createBtn_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog2 = new OpenFileDialog();
                openFileDialog2.InitialDirectory = "c:\\";
                openFileDialog2.Filter = "Excel files (*.xlsx)|*.xlsx| Excel 2007 (*.xls)| *.xls";
                openFileDialog2.FilterIndex = 1;

                if (openFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    DataTable datat = Excel.DataGridView_To_Datatable(dataGridView1);
                    datat.exportToExcel(openFileDialog2.FileName);
                    MessageBox.Show("Data is exported");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        
        

        private void convertBtn_Click(object sender, EventArgs e)
        {
          
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Select file";
            openFileDialog1.Filter = "All files (*.*)|*_*|XML File (*.xml)|*.xml";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.ShowDialog();
            if(openFileDialog1.FileName != " ")
            {
                ds.ReadXml(openFileDialog1.FileName);
                dataGridView1.DataSource = ds.Tables[0];
                MessageBox.Show("XML Data Imported");
            }
        }




        private void openBtn_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|* .xlsx", Multiselect = false })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    DataTable dt = new DataTable();
                    using (XLWorkbook wb = new XLWorkbook(ofd.FileName))
                    {

                        bool isFirstRow = true;
                        var rows = wb.Worksheet(1).RowsUsed();
                        foreach (var row in rows)
                        {
                            if (isFirstRow)
                            {
                                foreach (IXLCell cell in row.Cells())
                                    dt.Columns.Add(cell.Value.ToString());
                                isFirstRow = false;
                            }
                            else
                            {
                                dt.Rows.Add();
                                int i = 0;
                                foreach (IXLCell cell in row.Cells())
                                    dt.Rows[dt.Rows.Count - 1][i++] = cell.Value.ToString();
                            }
                        }
                        dataGridView1.DataSource = dt.DefaultView;
                        label3.Text = $"Total records: {dataGridView1.RowCount}";
                        Cursor.Current = Cursors.Default;
                    }
                }
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                DataView dv = dataGridView1.DataSource as DataView;
                if (dv != null)
                    dv.RowFilter = txtSearch.Text;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtSearch_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
                btnSearch.PerformClick();
           
        }
    }
}
