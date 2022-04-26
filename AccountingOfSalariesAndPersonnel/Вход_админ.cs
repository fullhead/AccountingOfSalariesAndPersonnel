using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;


namespace AccountingOfSalariesAndPersonnel
{
    public partial class Вход_админ : Form
    {
        private SqlConnection sqlConnection = null;
        public Вход_админ()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlDataAdapter Tablet = new SqlDataAdapter("Select Count (*) Login From Администраторы Where Логин ='" + textBox1.Text + "'and Пароль = '" + textBox2.Text + "'", sqlConnection);
            DataTable dt = new DataTable();
            Tablet.Fill(dt);
            if (dt.Rows[0][0].ToString() == "1")
            {
               
            }
            else
            {
                label3.Show();
            }
        }
    }
}
