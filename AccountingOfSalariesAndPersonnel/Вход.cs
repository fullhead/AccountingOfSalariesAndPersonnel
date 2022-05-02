using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;

namespace AccountingOfSalariesAndPersonnel
{
    public partial class Вход : Form
    {
        private SqlConnection sqlConnection = null;
        public Вход()
        {
            InitializeComponent();
        }

        private void Вход_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void ВходButton_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlDataAdapter Tablet = new SqlDataAdapter("Select Count (*) Login From Пользователи Where Логин ='" + textBox1.Text + "'and Пароль = '" + textBox2.Text + "'", sqlConnection);
            DataTable dt = new DataTable();
            Tablet.Fill(dt);
            if (dt.Rows[0][0].ToString() == "1")
            {
                Учёт_зарплат_и_кадров учёт_зарплат_и_кадров = new Учёт_зарплат_и_кадров();
                учёт_зарплат_и_кадров.Show();
                this.Hide();
            }
            else
            {
                label3.Show();
            }

        }
        private void TextBox2_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void tableLayoutPanel1_MouseMove(object sender, MouseEventArgs e)
        {
            label3.Hide();
        }
    }
}
