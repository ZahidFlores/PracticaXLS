using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace PracticaXLS
{
    public partial class Form1 : Form
    {
        static string server = "Data Source = DESKTOP-DDN13PI\\SQLEXPRESS; Initial Catalog= Reportes; Integrated Security = True ";
        SqlConnection conectar = new SqlConnection(server);
        DataTable dataTable = new DataTable();
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            conectar.Open();
            string consulta = "select * from ventas";
            SqlDataAdapter adapter = new SqlDataAdapter(consulta, conectar);
            adapter.Fill(dataTable);
            dgvReporte.DataSource = dataTable;
        }
        private void button1_Click(object sender, EventArgs e)
        {

                try
                {
                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook workbook = excelApp.Workbooks.Open(@"C:\Users\Zahid Flores\Desktop\Formato Reporte.xlsx");
                    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets["Hoja1"];
                    for (int i = 0; i < dgvReporte.RowCount - 1; i++)
                    {
                        for (int j = 0; j < dgvReporte.ColumnCount; j++)
                        {
                            worksheet.Cells[i + 10, j + 4] = dgvReporte.Rows[i].Cells[j].Value;
                        }
                    }
                    SaveFileDialog dialogo = new SaveFileDialog();
                    if (dialogo.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }
                    excelApp.ActiveWorkbook.SaveAs(dialogo.FileName);
                    MessageBox.Show("Datos guardados correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filtro = comboBox1.SelectedItem.ToString();
            if (!string.IsNullOrEmpty(filtro))
            {
                DataView dv = new DataView(dataTable); 
                dv.RowFilter = "Año = '" + filtro + "'"; 
                dgvReporte.DataSource = dv;
                button1.Enabled = true;
            }
        }
    }
}
