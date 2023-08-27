using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using OfficeOpenXml;
using ExcelDataReader;
using Z.Dapper.Plus;
using System.IO;

namespace InsertDataFromExcel
{
    public partial class Form1 : Form
    {
        private DataTableCollection tables;
        private DataTable excelDataTable;
        private string connectionString = "Server=CAD001\\WEB;Database=MTH;User=sa;Password=abc123";
        public Form1()
        {
            InitializeComponent();
        }

        private void find_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtPath.Text = ofd.FileName;
                    using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });
                            tables = result.Tables;
                            comboBox1.Items.Clear();
                            foreach (DataTable table in tables)
                                comboBox1.Items.Add(table.TableName);
                        }
                    }
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                DataTable dt = tables[comboBox1.SelectedItem.ToString()];

                // Remove empty columns
                List<DataColumn> columnsToRemove = new List<DataColumn>();
                foreach (DataColumn column in dt.Columns)
                {
                    bool isColumnEmpty = true;
                    foreach (DataRow row in dt.Rows)
                    {
                        if (!string.IsNullOrWhiteSpace(row[column.ColumnName].ToString()))
                        {
                            isColumnEmpty = false;
                            break;
                        }
                    }
                    if (isColumnEmpty)
                    {
                        columnsToRemove.Add(column);
                    }
                }

                foreach (DataColumn columnToRemove in columnsToRemove)
                {
                    dt.Columns.Remove(columnToRemove);
                }

                dataGridView1.DataSource = dt;
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                DataTable dt = tables[comboBox1.SelectedItem.ToString()];

                // Đảm bảo các cột bắt buộc tồn tại trong DataTable
                if (!dt.Columns.Contains("tk") || !dt.Columns.Contains("mk"))
                {
                    MessageBox.Show("Cột 'Tên tài khoản' và 'Mật khẩu' là bắt buộc.");
                    return;
                }

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    progressBar.Minimum = 0;
                    progressBar.Maximum = dt.Rows.Count;
                    progressBar.Step = 1;
                    progressBar.Value = 0;

                    foreach (DataRow row in dt.Rows)
                    {
                        string tk = row["tk"].ToString();
                        string mk = row["mk"].ToString();
                        string tb = Environment.MachineName;

                        // Thêm dữ liệu vào bảng SQL
                        string insertQuery = "INSERT INTO excel (tk, mk, tb) VALUES (@Username, @Password, @tb)";

                        using (SqlCommand command = new SqlCommand(insertQuery, connection))
                        {
                            command.Parameters.AddWithValue("@Username", tk);
                            command.Parameters.AddWithValue("@Password", mk);
                            command.Parameters.AddWithValue("@tb", tb);
                            command.ExecuteNonQuery();
                        }
                        progressBar.PerformStep();
                        int percent = (progressBar.Value * 100) / progressBar.Maximum;
                        progressBar.CreateGraphics().DrawString(percent.ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar.Width / 2 - 10, progressBar.Height / 2 - 7));
                    }

                    MessageBox.Show("Import dữ liệu thành công!", "Trạng Thái");

                }
                Application.Exit();
            }
        }
    }
}
