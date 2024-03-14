using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        readonly string connect = "Data Source=sql1c;Initial Catalog=Biblioteka;User ID=stud;Password=123456789";
        public Form1()
        {
            InitializeComponent();
            Meta();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void Update(string sqlText)
        {
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlText, connect);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable;
        }
        private void Meta()
        {
            string sqlText = "Select Fam, Ima, Otch, Nazvanie, DataVzatiya, DataVozvrata from Chitateli, Knigi, Knigovidacha where KodKnigi = FKKodKnigi and KodChitatela = FKKodchitatela";

            SqlDataAdapter adapter = new SqlDataAdapter(sqlText, connect);
            DataTable table = new DataTable();
            adapter.Fill(table);
            dataGridView1.DataSource = table;
            dataGridView1.Columns[0].HeaderText = "Фамилия";
            dataGridView1.Columns[1].HeaderText = "Имя";
            dataGridView1.Columns[2].HeaderText = "Отчество";
            dataGridView1.Columns[3].HeaderText = "Название";
            dataGridView1.Columns[4].HeaderText = "Дата взятия";
            dataGridView1.Columns[5].HeaderText = "Дата возврата";
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox1.Text))
                        {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
                    }
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            string comand = @"Select Fam, Ima, Otch, Nazvanie, DataVzatiya, DataVozvrata from Chitateli, Knigi, Knigovidacha where KodKnigi = FKKodKnigi and KodChitatela = FKKodchitatela";
            Update(comand);

            int f;
            for (int a = 0; a < dataGridView1.RowCount - 1; a++)
            {
                f = dataGridView1.RowCount - 2;
                if (a == f)
                {
                    richTextBox1.AppendText(dataGridView1[0, a].Value.ToString());
                }
                else
                {
                    richTextBox1.AppendText(dataGridView1[0, a].Value.ToString());
                }

            }
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            if (richTextBox1.Text != "")
            {
                string comand = @"Select Fam, Ima, Otch, Nazvanie, DataVzatiya, DataVozvrata from Chitateli, Knigi, Knigovidacha where KodKnigi = FKKodKnigi and KodChitatela = FKKodchitatela and Nazvanie like '%" + textBox1.Text + "%' or Fam like '%" + textBox1.Text + "%' or DataVzatiya like '%" + textBox1.Text + "%'";
                Update(comand);
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вывести отчет?",
   "Сообщение",
   MessageBoxButtons.OKCancel,
   MessageBoxIcon.Question,
   MessageBoxDefaultButton.Button1);
            if (result == DialogResult.OK)
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.Sheets[1];
                worksheet.Name = "Отчёт";

                worksheet.Cells[2, 4] = "Отчет о книговыдаче";
                Excel.Range rng1 = worksheet.Range[worksheet.Cells[2, 4], worksheet.Cells[2, 4]];
                rng1.Cells.Font.Name = "Times New Roman";
                rng1.Cells.Font.Size = 24;
                rng1.Font.Bold = true;
                rng1.Cells.Font.Color = ColorTranslator.ToOle(Color.Black);

                worksheet.Cells[3, 2] = "Фамилия";
                worksheet.Cells[3, 3] = "Имя";
                worksheet.Cells[3, 4] = "Отчество";
                worksheet.Cells[3, 5] = "Название книги";
                worksheet.Cells[3, 6] = "Дата Взятия";
                worksheet.Cells[3, 7] = "Дата Возврата";
                worksheet.Cells[5, 2] = "Заведующая сектором обслуживания";
                worksheet.Cells[5, 4] = "Рудель Н.Е.";
                worksheet.Cells[5, 6] = "Дата составления";
                worksheet.Columns[2].columnWidth = 18;
                worksheet.Columns[3].columnWidth = 18;
                worksheet.Columns[4].columnWidth = 18;
                worksheet.Columns[5].columnWidth = 20;
                worksheet.Columns[6].columnWidth = 20;
                worksheet.Columns[7].columnWidth = 20;

                Excel.Range rng2 = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[8, 1]];

                rng2.Font.Bold = true;

                worksheet.Cells[4, 2] = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                worksheet.Cells[4, 3] = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                worksheet.Cells[4, 4] = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                worksheet.Cells[4, 5] = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                worksheet.Cells[4, 6] = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                worksheet.Cells[4, 7] = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                worksheet.Cells[5, 7] = dateTimePicker1.Value.ToString("MM-dd-yyyy");


                Excel.Range rng3 = worksheet.Range[worksheet.Cells[3, 2], worksheet.Cells[4, 7]];

                rng3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                rng3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                rng3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
                rng3.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
                rng3.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                rng3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

                Excel.Range rng4 = worksheet.Range[worksheet.Cells[2, 2], worksheet.Cells[2, 7]];

                rng4.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                rng4.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                rng4.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
                rng4.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

                excelApp.Visible = true;

                excelApp.UserControl = true;
            }
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
        }
    }
}
