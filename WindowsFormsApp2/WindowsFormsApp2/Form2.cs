using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Form2 : Form
    {
        readonly string connect = "Data Source=sql1c;Initial Catalog=Biblioteka;User ID=stud;Password=123456789";
        public Form2()
        {
            InitializeComponent();
            Meta();
            MetaChitat();
            MetaKnigi();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            Cb();
            CbKn();
        }
        private void Cb()
        {
            SqlDataAdapter adapter = new SqlDataAdapter("Select Fam+' '+Ima+' '+Otch from Chitateli", connect);
            DataTable table = new DataTable();
            adapter.Fill(table);

            try
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    comboBox1.Items.Add(table.Rows[i][0].ToString());
                }
            }
            catch { }
        }
        private void CbKn()
        {
            SqlDataAdapter adapter = new SqlDataAdapter("Select KodKnigi, NomerKnigi+' '+Nazvanie from Knigi", connect);
            DataTable table = new DataTable();
            adapter.Fill(table);

            comboBox2.DataSource = table.AsDataView();
            comboBox2.DisplayMember = "Nazvanie";
            comboBox2.ValueMember = "KodKnigi";
        }

        private void Meta()
        {
            string sqlText = "Select Knigovidacha.KodKV as [ID], Chitateli.Fam+' '+Chitateli.Ima+' '+' '+Chitateli.Otch, Knigi.Nazvanie, Knigovidacha.DataVzatiya, Knigovidacha.DataVozvrata from Knigovidacha, Chitateli, Knigi " +
                "where Chitateli.KodChitatela = Knigovidacha.FKkodChitatela and Knigi.KodKnigi = Knigovidacha.FKkodKnigi ";

            SqlDataAdapter adapter = new SqlDataAdapter(sqlText, connect);
            DataTable table = new DataTable();
            adapter.Fill(table);
            dataGridView1.DataSource = table;
            dataGridView1.Columns[1].HeaderText = "ФИО читателя";
            dataGridView1.Columns[2].HeaderText = "Название книги";
            dataGridView1.Columns[3].HeaderText = "Дата взятия";
            dataGridView1.Columns[4].HeaderText = "Дата возврата";
        }
        private void MetaChitat()
        {
            string sqlText = "Select KodChitatela, Fam, Ima, Otch from Chitateli ";

            SqlDataAdapter adapter = new SqlDataAdapter(sqlText, connect);
            DataTable table = new DataTable();
            adapter.Fill(table);
            dataGridView2.DataSource = table;
            dataGridView2.Columns[0].HeaderText = "Код читателя";
            dataGridView2.Columns[1].HeaderText = "Фамилия";
            dataGridView2.Columns[2].HeaderText = "Имя";
            dataGridView2.Columns[3].HeaderText = "Отчество";
        }
        private void MetaKnigi()
        {
            string sqlText = "Select KodKnigi, Nazvanie from Knigi";

            SqlDataAdapter adapter = new SqlDataAdapter(sqlText, connect);
            DataTable table = new DataTable();
            adapter.Fill(table);
            dataGridView3.DataSource = table;
            dataGridView3.Columns[0].HeaderText = "Код книги";
            dataGridView3.Columns[1].HeaderText = "Название";

        }

        private void DataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Добавить запись?",
                   "Сообщение",
                   MessageBoxButtons.OKCancel,
                   MessageBoxIcon.Question,
                   MessageBoxDefaultButton.Button1);
            if (result == DialogResult.OK)
            {
                string sqlText = "INSERT INTO Knigovidacha VALUES" +
                    "('" + comboBox1.Text.ToString() + "','"
                    + comboBox2.Text.ToString() + "',"
                    + "'" + dateTimePicker1.Text.ToString() + "',"
                    + "'" + dateTimePicker2.Text.ToString() + "')";
                using (SqlConnection con = new SqlConnection(connect))
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand(sqlText, con);
                    int kol = cmd.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Записей добавлено: " + kol.ToString(),
                        "Сообщение",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                Meta();
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Изменить запись?",
                "Сообщение",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1);
            if (result == DialogResult.OK)
            {
                string sqlText = "UPDATE Knigavidacha SET "
                    + "FKKodchitatela = '" + comboBox1.Text.ToString() + "',"
                    + "FKKodKnigi = '" + comboBox2.Text.ToString() + "',"
                    + "DataVzatiya = '" + dateTimePicker1.Value.Date.ToShortDateString() + "',"
                    + "DataVozvrata = '" + dateTimePicker2.Value.Date.ToShortDateString() + "'"
                    + "Where KodKV = '" + dataGridView1.CurrentRow.Cells[0].Value.ToString() + "'";

                using (SqlConnection con = new SqlConnection(connect))
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand(sqlText, con);
                    int kol = cmd.ExecuteNonQuery();
                    con.Close();
                }
                Meta();
            }

        }
        private void DataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            comboBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            comboBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            dateTimePicker1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            dateTimePicker2.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите удалить эти записи?", "Удаление записей", MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                string id;
                id = dataGridView1.SelectedRows[0].Cells["ID"].Value.ToString();

                string sqlText = "delete from Knigovidacha where KodKV =" + id;
                SqlDataAdapter adapter = new SqlDataAdapter(sqlText, connect);
                DataTable table = new DataTable();
                adapter.Fill(table);
                dataGridView1.DataSource = table;

                Meta();
            }
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
