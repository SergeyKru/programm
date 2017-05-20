using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace _14_6_
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OleDbDataAdapter adapter;

        DataSet dataset;
        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бДDataSet.Владельцы". При необходимости она может быть перемещена или удалена.
            this.владельцыTableAdapter.Fill(this.бДDataSet.Владельцы);
            adapter = new OleDbDataAdapter("Select Код, Фамилия, Имя, Отчество, Телефон from Владельцы", new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"БД.mdb"));

            dataset = new DataSet();

            adapter.Fill(dataset);

            dataGridView1.DataSource = dataset.Tables[0];
        }
        private void Добавить()
        {
            string conn_param = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Сергей\Desktop\ПИС\14\14(6)\14(6)\bin\Debug\БД.mdb";
            OleDbConnection connection = new OleDbConnection(conn_param);
            OleDbCommand command1 = connection.CreateCommand();
            string insSURNAME = textBox1.Text;
            string insNAME = textBox2.Text;
            string insPATRO = textBox3.Text;
            string insTel = textBox4.Text;
            connection.Open();
            command1.CommandText = "select max(Код) from Владельцы";
            OleDbDataReader reader = command1.ExecuteReader();
            reader.Read();
            int id = reader.GetInt32(0) + 1;
            reader.Close();

            command1.CommandText = "insert into Владельцы (Код, Фамилия, Имя, Отчество, Телефон) values (" + id + ",'" + insSURNAME + "', '" + insNAME + "', '" + insPATRO + "','" + insTel + "')";


            command1.ExecuteNonQuery();
            connection.Close();
            Refresh();
        }
        private void Обновить()
        {
            adapter = new OleDbDataAdapter("Select Код, Фамилия, Имя, Отчество, Телефон from Владельцы", new
            OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"C:\Users\Сергей\Desktop\ПИС\14\14(6)\14(6)\bin\Debug\БД.mdb"));
            dataset = new DataSet();
            adapter.Fill(dataset);
            dataGridView1.DataSource = dataset.Tables[0];
            string conn_param = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Сергей\Desktop\ПИС\14\14(6)\14(6)\bin\Debug\БД.mdb";
            OleDbConnection connection = new OleDbConnection(conn_param);

            OleDbCommand command1 = connection.CreateCommand();
            command1.CommandText = "select Код, Фамилия, Имя, Отчество, Телефон from Владельцы";

            connection.Open();
            command1.ExecuteNonQuery();
            connection.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Добавить();
            Обновить();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int id = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            string surname = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            string name = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            string patro = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            string telephone = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();

            string conn_param = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Сергей\Desktop\ПИС\14\14(6)\14(6)\bin\Debug\БД.mdb";
            OleDbConnection connection = new OleDbConnection(conn_param);
            OleDbCommand command1 = connection.CreateCommand();

            command1.CommandText = "UPDATE Владельцы SET Фамилия = '" + surname + "', Имя='" + name + "', Отчество='" + patro +  "', Телефон='" + telephone + "' WHERE Код=" + id + "";

            connection.Open();
            command1.ExecuteNonQuery();
            connection.Close();
            Refresh();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int id = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());

            string conn_param = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Сергей\Desktop\ПИС\14\14(6)\14(6)\bin\Debug\БД.mdb";
            OleDbConnection connection = new OleDbConnection(conn_param);
            OleDbCommand command1 = connection.CreateCommand();

            command1.CommandText = "DELETE FROM Владельцы WHERE Код=" + id + "";

            connection.Open();
            command1.ExecuteNonQuery();
            connection.Close();
            Обновить();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            string name = textBox6.Text;
            string second_name = textBox5.Text;
            string surname = textBox7.Text;
            string phone = textBox8.Text;
            string conn_param = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source = БД.mdb";
            OleDbConnection connection = new OleDbConnection(conn_param);
            OleDbCommand command = connection.CreateCommand();

            command.CommandText = "UPDATE Владельцы SET Имя=\"" + name + "\",Фамилия=\"" + second_name + "\",Отчество=\"" + surname + "\",Телефон=\"" + phone + "\" WHERE Код =" + comboBox1.Text + "";

            connection.Open();

            command.ExecuteNonQuery();

            connection.Close();

            Form1_Load(null, null);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string conn_param = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source = БД.mdb";

            OleDbConnection connection = new OleDbConnection(conn_param);
            int l = Convert.ToInt32(comboBox1.Text);
            OleDbCommand command = connection.CreateCommand();
            OleDbCommand command1 = connection.CreateCommand();
            OleDbCommand command2 = connection.CreateCommand();
            OleDbCommand command3 = connection.CreateCommand();
            command.CommandText = "SELECT Имя FROM Владельцы WHERE код=" + l + "";
            command1.CommandText = "SELECT Фамилия FROM Владельцы WHERE код=" + l + "";
            command2.CommandText = "SELECT Отчество FROM Владельцы WHERE код=" + l + "";
            command3.CommandText = "SELECT Телефон FROM Владельцы WHERE код=" + l + "";
            connection.Open();
            textBox6.Text = Convert.ToString(command.ExecuteScalar());
            textBox5.Text = Convert.ToString(command1.ExecuteScalar());
            textBox7.Text = Convert.ToString(command2.ExecuteScalar());
            textBox8.Text = Convert.ToString(command3.ExecuteScalar());
            command.ExecuteNonQuery();
            connection.Close();
        }
    }
}