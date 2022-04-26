using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using Tulpep.NotificationWindow;
using System.IO;

namespace AccountingOfSalariesAndPersonnel
{
    public partial class Учёт_зарплат_и_кадров : Form
    {
        private SqlConnection sqlConnection = null;
        private PopupNotifier popup = null;
        private SqlDataAdapter adapter = null;
        private DataTable table = null;
        public Учёт_зарплат_и_кадров()
        {
            InitializeComponent();
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView3.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView4.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView5.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView6.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView7.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView8.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        }

        //STATUS DB
        private void Учёт_зарплат_и_кадров_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfSalariesAndPersonnelDataSet.Пользователи". При необходимости она может быть перемещена или удалена.
            this.пользователиTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Пользователи);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfSalariesAndPersonnelDataSet.Пользователи". При необходимости она может быть перемещена или удалена.
            this.пользователиTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Пользователи);
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            if (sqlConnection.State == ConnectionState.Open)
                pictureBox1.Image = Properties.Resources.connected;

            else
                pictureBox1.Image = Properties.Resources.disconnect;
        }

        private void Учёт_зарплат_и_кадров_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void панельАдминистратораToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Вход_админ вход_админ = new Вход_админ();
            вход_админ.Show();
        }
    }
}
