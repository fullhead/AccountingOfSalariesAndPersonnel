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
            tabPage29.Parent = null;
        }

        //EXIT .EXE
        private void Учёт_зарплат_и_кадров_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        //ACCESS FOR ADMIN
        private void ВходToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            string result = Microsoft.VisualBasic.Interaction.InputBox("Введите пароль администратора:");
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlDataAdapter Tablet = new SqlDataAdapter("Select Count (*) Login From Администраторы Where Пароль = '" + result + "'", sqlConnection);
            DataTable dt = new DataTable();
            Tablet.Fill(dt);
            if (dt.Rows[0][0].ToString() == "1")
            {
                tabPage29.Parent = tabControl1;
            }
            else
            {
                MessageBox.Show("Неправильный пароль!");
            }
        }
        private void ВыходToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            tabPage29.Parent = null;
        }

        //STATUS DB
        private void Учёт_зарплат_и_кадров_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfSalariesAndPersonnelDataSet.Пользователи". При необходимости она может быть перемещена или удалена.
            this.пользователиTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Пользователи);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfSalariesAndPersonnelDataSet.Штатное_расписание". При необходимости она может быть перемещена или удалена.
            this.штатное_расписаниеTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Штатное_расписание);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfSalariesAndPersonnelDataSet.Трудовые_договора". При необходимости она может быть перемещена или удалена.
            this.трудовые_договораTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Трудовые_договора);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfSalariesAndPersonnelDataSet.Сотрудники". При необходимости она может быть перемещена или удалена.
            this.сотрудникиTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Сотрудники);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfSalariesAndPersonnelDataSet.Отпуски". При необходимости она может быть перемещена или удалена.
            this.отпускиTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Отпуски);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfSalariesAndPersonnelDataSet.Начисление_ЗП". При необходимости она может быть перемещена или удалена.
            this.начисление_ЗПTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Начисление_ЗП);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfSalariesAndPersonnelDataSet.Командировки". При необходимости она может быть перемещена или удалена.
            this.командировкиTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Командировки);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfSalariesAndPersonnelDataSet.Должности". При необходимости она может быть перемещена или удалена.
            this.должностиTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Должности);
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            if (sqlConnection.State == ConnectionState.Open)
                pictureBox1.Image = Properties.Resources.connected;

            else
                pictureBox1.Image = Properties.Resources.disconnect;
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            О_программе о_программе = new О_программе();
            о_программе.Show();
        }
    }
}
