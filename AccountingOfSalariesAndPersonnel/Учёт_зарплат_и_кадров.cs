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
            this.ДолжностиDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.КомандировкиDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.Начисление_зпDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.ОтпускиDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.СотрудникиDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.Трудовые_договораDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.Штатное_расписаниеDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.ПользователиDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            ПользователиPage.Parent = null;
            
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
                ПользователиPage.Parent = DB_pages;
                пользователиToolStripMenuItem.Visible = true;
                пользователиToolStripMenuItem1.Visible = true;
            }
            else
            {
                MessageBox.Show("Неправильный пароль!");
            }
        }
        private void ВыходToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            ПользователиPage.Parent = null;
            пользователиToolStripMenuItem.Visible = false;
            пользователиToolStripMenuItem1.Visible = false;
        }

        //STATUS AND LOAD DB
        private void Учёт_зарплат_и_кадров_Load(object sender, EventArgs e)
        {
            this.пользователиTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Пользователи);
            this.штатное_расписаниеTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Штатное_расписание);
            this.трудовые_договораTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Трудовые_договора);
            this.сотрудникиTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Сотрудники);
            this.отпускиTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Отпуски);
            this.начисление_ЗПTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Начисление_ЗП);
            this.командировкиTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Командировки);
            this.должностиTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Должности);

            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            if (sqlConnection.State == ConnectionState.Open)
                pictureBox1.Image = Properties.Resources.connected;

            else
                pictureBox1.Image = Properties.Resources.disconnect;
        }
        
        //SEARCH
        private void Должности_SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                adapter = new SqlDataAdapter("SELECT * from Должности where Наименование like'%" + Должности_SearchTextBox.Text + "%'", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                ДолжностиDataGrid.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
        }
        private void Командировки_SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                adapter = new SqlDataAdapter("SELECT * from Командировки where Место like'%" + Командировки_SearchTextBox.Text + "%'", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                КомандировкиDataGrid.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }

        }
        private void Начисление_ЗП_SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                adapter = new SqlDataAdapter("SELECT * from Начисление_ЗП where Дата_выплаты like'%" + Начисление_ЗП_SearchTextBox.Text + "%'", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                Начисление_зпDataGrid.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
        }

        private void Отпуски_SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                adapter = new SqlDataAdapter("SELECT * from Отпуски where Длительность like'%" + Отпуски_SearchTextBox.Text + "%'", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                ОтпускиDataGrid.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
        }

        private void Сотрудники_SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                adapter = new SqlDataAdapter("SELECT * from Сотрудники where ФИО like'%" + Сотрудники_SearchTextBox.Text + "%'", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                СотрудникиDataGrid.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
        }

        private void Трудовые_договора_SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                adapter = new SqlDataAdapter("SELECT * from Трудовые_договора where Дата_заключения like'%" + Трудовые_договора_SearchTextBox.Text + "%'", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                Трудовые_договораDataGrid.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
        }

        private void Штатное_расписание_SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                adapter = new SqlDataAdapter("SELECT * from Штатное_расписание where Наименование_структурного_подразделения like'%" + Штатное_расписание_SearchTextBox.Text + "%'", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                Штатное_расписаниеDataGrid.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
        }

        private void Пользователи_SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfCourtCases.Properties.Settings.AccountingOfCourtCasesConnectionString"].ConnectionString);
                sqlConnection.Open();
                adapter = new SqlDataAdapter("SELECT * from Пользователи where ФИО like'%" + Должности_SearchTextBox.Text + "%'", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                ПользователиDataGrid.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
        }

        //PRINT
        private void ДолжностиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Class_Print _Class_Print = new Class_Print(ДолжностиDataGrid, "Таблица [Должности]");
            _Class_Print.PrintForm();
        }

        private void КомандировкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Class_Print _Class_Print = new Class_Print(КомандировкиDataGrid, "Таблица [Командировки]");
            _Class_Print.PrintForm();
        }

        private void НачислениеЗпToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Class_Print _Class_Print = new Class_Print(Начисление_зпDataGrid, "Таблица [Начисление]");
            _Class_Print.PrintForm();
        }

        private void ОтпускиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Class_Print _Class_Print = new Class_Print(ОтпускиDataGrid, "Таблица [Отпуски]");
            _Class_Print.PrintForm();
        }

        private void СотрудникиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Class_Print _Class_Print = new Class_Print(СотрудникиDataGrid, "Таблица [Сотрудники]");
            _Class_Print.PrintForm();
        }

        private void ТрудовыеДоговораToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Class_Print _Class_Print = new Class_Print(Трудовые_договораDataGrid, "Таблица [Трудовые договора]");
            _Class_Print.PrintForm();
        }

        private void ШтатноеРасписаниеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Class_Print _Class_Print = new Class_Print(Штатное_расписаниеDataGrid, "Таблица [Штатное расписание]");
            _Class_Print.PrintForm();
        }

        private void ПользователиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Class_Print _Class_Print = new Class_Print(ПользователиDataGrid, "Таблица [Пользователи]");
            _Class_Print.PrintForm();

        }

        //SAVE FOR .CSV
        private void ДолжностиToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            try
            {
                adapter = new SqlDataAdapter("SELECT * FROM Должности", sqlConnection);
                adapter.Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
            string path = "";
            using (var path_dialog = new FolderBrowserDialog())
                if (path_dialog.ShowDialog() == DialogResult.OK)
                {
                    //Путь к директории
                    path = path_dialog.SelectedPath;
                };
            sqlConnection.Close();
            ToCSVДолжности(dt, path + @"\" + @"Отчёт_Должности.csv");
        }
        public static void ToCSVДолжности(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false, Encoding.UTF8);
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(';'))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        private void КомандировкиToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            try
            {
                adapter = new SqlDataAdapter("SELECT * FROM Командировки", sqlConnection);
                adapter.Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
            string path = "";
            using (var path_dialog = new FolderBrowserDialog())
                if (path_dialog.ShowDialog() == DialogResult.OK)
                {
                    //Путь к директории
                    path = path_dialog.SelectedPath;
                };
            sqlConnection.Close();
            ToCSVКомандировки(dt, path + @"\" + @"Отчёт_Командировки.csv");
        }
        public static void ToCSVКомандировки(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false, Encoding.UTF8);
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(';'))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        private void НачислениеЗПToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            try
            {
                adapter = new SqlDataAdapter("SELECT * FROM Начисление_ЗП", sqlConnection);
                adapter.Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
            string path = "";
            using (var path_dialog = new FolderBrowserDialog())
                if (path_dialog.ShowDialog() == DialogResult.OK)
                {
                    //Путь к директории
                    path = path_dialog.SelectedPath;
                };
            sqlConnection.Close();
            ToCSVНачисление_ЗП(dt, path + @"\" + @"Отчёт_Начисление_ЗП.csv");
        }
        public static void ToCSVНачисление_ЗП(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false, Encoding.UTF8);
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(';'))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        private void ОтпускиToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            try
            {
                adapter = new SqlDataAdapter("SELECT * FROM Отпуски", sqlConnection);
                adapter.Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
            string path = "";
            using (var path_dialog = new FolderBrowserDialog())
                if (path_dialog.ShowDialog() == DialogResult.OK)
                {
                    //Путь к директории
                    path = path_dialog.SelectedPath;
                };
            sqlConnection.Close();
            ToCSVОтпуски(dt, path + @"\" + @"Отчёт_Отпуски.csv");
        }
        public static void ToCSVОтпуски(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false, Encoding.UTF8);
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(';'))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        private void СотрудникиToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            try
            {
                adapter = new SqlDataAdapter("SELECT * FROM Сотрудники", sqlConnection);
                adapter.Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
            string path = "";
            using (var path_dialog = new FolderBrowserDialog())
                if (path_dialog.ShowDialog() == DialogResult.OK)
                {
                    //Путь к директории
                    path = path_dialog.SelectedPath;
                };
            sqlConnection.Close();
            ToCSVСотрудники(dt, path + @"\" + @"Отчёт_Сотрудники.csv");
        }
        public static void ToCSVСотрудники(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false, Encoding.UTF8);
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(';'))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        private void ТрудовыеДоговораToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            try
            {
                adapter = new SqlDataAdapter("SELECT * FROM Трудовые_договора", sqlConnection);
                adapter.Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
            string path = "";
            using (var path_dialog = new FolderBrowserDialog())
                if (path_dialog.ShowDialog() == DialogResult.OK)
                {
                    //Путь к директории
                    path = path_dialog.SelectedPath;
                };
            sqlConnection.Close();
            ToCSVТрудовые_договора(dt, path + @"\" + @"Отчёт_Трудовые_договора.csv");
        }
        public static void ToCSVТрудовые_договора(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false, Encoding.UTF8);
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(';'))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        private void ШтатноеРасписаниеToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            try
            {
                adapter = new SqlDataAdapter("SELECT * FROM Штатное_расписание", sqlConnection);
                adapter.Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
            string path = "";
            using (var path_dialog = new FolderBrowserDialog())
                if (path_dialog.ShowDialog() == DialogResult.OK)
                {
                    //Путь к директории
                    path = path_dialog.SelectedPath;
                };
            sqlConnection.Close();
            ToCSVШтатное_расписание(dt, path + @"\" + @"Отчёт_Штатное_расписание.csv");
        }
        public static void ToCSVШтатное_расписание(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false, Encoding.UTF8);
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(';'))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        private void ПользователиToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            try
            {
                adapter = new SqlDataAdapter("SELECT * FROM Пользователи", sqlConnection);
                adapter.Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sqlConnection.Close();
            }
            string path = "";
            using (var path_dialog = new FolderBrowserDialog())
                if (path_dialog.ShowDialog() == DialogResult.OK)
                {
                    //Путь к директории
                    path = path_dialog.SelectedPath;
                };
            sqlConnection.Close();
            ToCSVПользователи(dt, path + @"\" + @"Отчёт_Пользователи.csv");
        }
        public static void ToCSVПользователи(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false, Encoding.UTF8);
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(';'))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();

        }

        //
        private void ОПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            О_программе о_программе = new О_программе();
            о_программе.Show();
        }


        private void ОбновитьButton_Click(object sender, EventArgs e)
        {
            this.ДолжностиDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

        }


    }
}
