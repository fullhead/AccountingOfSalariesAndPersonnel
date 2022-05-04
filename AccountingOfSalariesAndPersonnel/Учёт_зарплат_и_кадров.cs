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
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
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
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
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
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
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
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
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
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
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
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
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
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
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
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                adapter = new SqlDataAdapter("SELECT * from Пользователи where ФИО like'%" + Пользователи_SearchTextBox.Text + "%'", sqlConnection);
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
                }
                else
                {
                    return;
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
                }
                else
                {
                    return;
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
                }
                else
                {
                    return;
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
                }
                else
                {
                    return;
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
                }
                else
                {
                    return;
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
                }
                else
                {
                    return;
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
                }
                else
                {
                    return;
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
                }
                else
                {
                    return;
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

        ////ДОЛЖНОСТИ
        //Долж_Обн_TabPage
        private async void ОбновитьButton_Click(object sender, EventArgs e)
        {

            if (наименованиеTextBox.Text == "" || this.окладTextBox.Text == "" || this.обязанностиTextBox.Text == "")
            {
                Долж_Обн_Обяз_зап_label.Show();
                Долж_Обн_Обяз_зап_label1.Show();
                Долж_Обн_Обяз_зап_label2.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("UPDATE Должности SET Наименование=@Наименование, Оклад=@Оклад, Обязанности=@Обязанности, Примечание=@Примечание WHERE Код_должности=@Код_должности", sqlConnection);
                command.Parameters.AddWithValue("Код_должности", код_должностиComboBox.Text);
                command.Parameters.AddWithValue("Наименование", наименованиеTextBox.Text);
                command.Parameters.AddWithValue("Оклад", окладTextBox.Text);
                command.Parameters.AddWithValue("Обязанности", обязанностиTextBox.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Должности",
                    ContentText = "Данные успешно обновлены!"
                };
                popup.Popup();
                adapter = new SqlDataAdapter("SELECT * FROM Должности", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                ДолжностиDataGrid.DataSource = table;
                await command.ExecuteNonQueryAsync();
                
            }

        }
        private void TableLayoutPanel2_MouseMove_1(object sender, MouseEventArgs e)
        {
            Долж_Обн_Обяз_зап_label.Hide();
            Долж_Обн_Обяз_зап_label1.Hide();
            Долж_Обн_Обяз_зап_label2.Hide();
        }
        private void НаименованиеTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && (l < '0' || l > '9') && l != '\b' && l != '.' && l != ',' && l != ' ' && l != '"')
            {
                e.Handled = true;
            }
        }
        private void ОкладTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        //Долж_Доб_TabPage
        private async void ДобавитьButton_Click(object sender, EventArgs e)
        {
            if (наименованиеTextBox1.Text == "" || this.окладTextBox1.Text == "" || this.обязанностиTextBox1.Text == "")
            {
                Долж_Доб_Обяз_зап_label.Show();
                Долж_Доб_Обяз_зап_label1.Show();
                Долж_Доб_Обяз_зап_label2.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Должности (Наименование, Оклад, Обязанности, Примечание) VALUES (@Наименование, @Оклад, @Обязанности, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("Наименование", наименованиеTextBox1.Text);
                command.Parameters.AddWithValue("Оклад", окладTextBox1.Text);
                command.Parameters.AddWithValue("Обязанности", обязанностиTextBox1.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox1.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Должности",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Должности", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                ДолжностиDataGrid.DataSource = table;
                код_должностиComboBox.DataSource = table;
                код_должностиComboBox1.DataSource = table;
            }
            
        }
        private void TableLayoutPanel3_MouseMove(object sender, MouseEventArgs e)
        {
            Долж_Доб_Обяз_зап_label.Hide();
            Долж_Доб_Обяз_зап_label1.Hide();
            Долж_Доб_Обяз_зап_label2.Hide();
        }
        private void ОкладTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }
        private void НаименованиеTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && (l < '0' || l > '9') && l != '\b' && l != '.' && l != ',' && l != ' ' && l != '"')
            {
                e.Handled = true;
            }
        }

        //Долж_Удал_TabPage
        private async void УдалитьButton_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM Должности WHERE Код_должности=@Код_должности", sqlConnection);
            command.Parameters.AddWithValue("Код_должности", код_должностиComboBox1.Text);
            popup = new PopupNotifier
            {
                Image = Properties.Resources.connected,
                ImageSize = new Size(96, 96),
                TitleText = "Должности",
                ContentText = "Данные успешно удалены!"
            };
            popup.Popup();
            await command.ExecuteNonQueryAsync();
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            adapter = new SqlDataAdapter("SELECT * FROM Должности", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            ДолжностиDataGrid.DataSource = table;
            код_должностиComboBox.DataSource = table;
            код_должностиComboBox1.DataSource = table;
        }

        ////КОМАНДИРОВКИ
        //Ком_Обн_TabPage
        private async void ОбновитьButton1_Click(object sender, EventArgs e)
        {
            if (дата_командировкиDateTimePicker.Text == "" || this.длительностьTextBox.Text == "" || this.местоComboBox.Text == "")
            {
                Ком_Обн_Обяз_зап_label.Show();
                Ком_Обн_Обяз_зап_label1.Show();
                Ком_Обн_Обяз_зап_label2.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("UPDATE Командировки SET Код_сотрудника=@Код_сотрудника, Дата_командировки=@Дата_командировки, Длительность=@Длительность, Место=@Место, Цель=@Цель, Оплата=@Оплата, Примечание=@Примечание WHERE Код_командировки=@Код_командировки", sqlConnection);
                command.Parameters.AddWithValue("Код_командировки", Код_командировкиComboBox.Text);
                command.Parameters.AddWithValue("Код_сотрудника", код_сотрудникаComboBox.Text);
                command.Parameters.AddWithValue("Дата_командировки", дата_командировкиDateTimePicker.Text);
                command.Parameters.AddWithValue("Длительность", длительностьTextBox.Text);
                command.Parameters.AddWithValue("Место", местоComboBox.Text);
                command.Parameters.AddWithValue("Цель", цельTextBox.Text);
                command.Parameters.AddWithValue("Оплата", оплатаTextBox.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox2.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Командировки",
                    ContentText = "Данные успешно обновлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Командировки", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                КомандировкиDataGrid.DataSource = table;
            }
        }
        private void TableLayoutPanel5_MouseMove(object sender, MouseEventArgs e)
        {
            Ком_Обн_Обяз_зап_label.Hide();
            Ком_Обн_Обяз_зап_label1.Hide();
            Ком_Обн_Обяз_зап_label2.Hide();
        }
        private void ДлительностьTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && (l < '0' || l > '9') && l != '\b' && l != '.' && l != ',' && l != ' ' && l != '"')
            {
                e.Handled = true;
            }
        }

        //Ком_Доб_TabPage
        private async void ДобавитьButton1_Click(object sender, EventArgs e)
        {
            if (дата_командировкиDateTimePicker1.Text == "" || this.длительностьTextBox1.Text == "" || this.местоComboBox1.Text == "")
            {
                Ком_Доб_Обяз_зап_label.Show();
                Ком_Доб_Обяз_зап_label1.Show();
                Ком_Доб_Обяз_зап_label2.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Командировки (Код_сотрудника, Дата_командировки, Длительность, Место, Цель, Оплата, Примечание) VALUES (@Код_сотрудника, @Дата_командировки, @Длительность, @Место, @Цель, @Оплата, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("Код_сотрудника", код_сотрудникаComboBox1.Text);
                command.Parameters.AddWithValue("Дата_командировки", дата_командировкиDateTimePicker1.Text);
                command.Parameters.AddWithValue("Длительность", длительностьTextBox1.Text);
                command.Parameters.AddWithValue("Место", местоComboBox1.Text);
                command.Parameters.AddWithValue("Цель", цельTextBox1.Text);
                command.Parameters.AddWithValue("Оплата", оплатаTextBox1.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox3.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Командировки",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Командировки", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                КомандировкиDataGrid.DataSource = table;
                Код_командировкиComboBox.DataSource = table;
                Код_командировкиComboBox1.DataSource = table;
            }
        }
        private void TableLayoutPanel6_MouseMove(object sender, MouseEventArgs e)
        {
            Ком_Доб_Обяз_зап_label.Hide();
            Ком_Доб_Обяз_зап_label1.Hide();
            Ком_Доб_Обяз_зап_label2.Hide();
        }
        private void ДлительностьTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && (l < '0' || l > '9') && l != '\b' && l != '.' && l != ',' && l != ' ' && l != '"')
            {
                e.Handled = true;
            }
        }

        //Ком_Удал_TabPage
        private async void УдалитьButton1_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM Командировки WHERE Код_командировки=@Код_командировки", sqlConnection);
            command.Parameters.AddWithValue("Код_командировки", Код_командировкиComboBox1.Text);
            popup = new PopupNotifier
            {
                Image = Properties.Resources.connected,
                ImageSize = new Size(96, 96),
                TitleText = "Командировки",
                ContentText = "Данные успешно удалены!"
            };
            popup.Popup();
            await command.ExecuteNonQueryAsync();
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            adapter = new SqlDataAdapter("SELECT * FROM Командировки", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            КомандировкиDataGrid.DataSource = table;
            Код_командировкиComboBox.DataSource = table;
            Код_командировкиComboBox1.DataSource = table;
        }

        ////СОТРУДНИКИ
        //Сот_Обн_TabPage
        private async void ОбновитьButton5_Click(object sender, EventArgs e)
        {
            if (фИОTextBox.Text == "" || this.полComboBox.Text == "" || this.возрастComboBox.Text == "" || this.паспортные_данныеTextBox.Text == "")
            {
                Сот_Обн_Обяз_зап_label.Show();
                Сот_Обн_Обяз_зап_label1.Show();
                Сот_Обн_Обяз_зап_label2.Show();
                Сот_Обн_Обяз_зап_label3.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("UPDATE Сотрудники SET Код_должности=@Код_должности, ФИО=@ФИО, Пол=@Пол, Возраст=@Возраст, Адрес=@Адрес, Паспортные_данные=@Паспортные_данные, Примечание=@Примечание WHERE Код_сотрудника=@Код_сотрудника", sqlConnection);
                command.Parameters.AddWithValue("Код_сотрудника", код_сотрудникаComboBox6.Text);
                command.Parameters.AddWithValue("Код_должности", код_должностиComboBox2.Text);
                command.Parameters.AddWithValue("ФИО", фИОTextBox.Text);
                command.Parameters.AddWithValue("Пол", полComboBox.Text);
                command.Parameters.AddWithValue("Возраст", возрастComboBox.Text);
                command.Parameters.AddWithValue("Адрес", адресTextBox.Text);
                command.Parameters.AddWithValue("Паспортные_данные", паспортные_данныеTextBox.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox8.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Сотрудники",
                    ContentText = "Данные успешно обновлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Сотрудники", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                СотрудникиDataGrid.DataSource = table;
            }
        }
        private void TableLayoutPanel14_MouseMove(object sender, MouseEventArgs e)
        {
            Сот_Обн_Обяз_зап_label.Hide();
            Сот_Обн_Обяз_зап_label1.Hide();
            Сот_Обн_Обяз_зап_label2.Hide();
            Сот_Обн_Обяз_зап_label3.Hide();
        }
        private void ФИОTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != ' ')
            {
                e.Handled = true;
            }
        }
        private void Паспортные_данныеTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        //Сот_Доб_TabPage
        private async void ДобавитьButton5_Click(object sender, EventArgs e)
        {
            if (фИОTextBox1.Text == "" || this.полComboBox1.Text == "" || this.возрастComboBox1.Text == "" || this.паспортные_данныеTextBox1.Text == "")
            {
                Сот_Доб_Обяз_зап_label.Show();
                Сот_Доб_Обяз_зап_label1.Show();
                Сот_Доб_Обяз_зап_label2.Show();
                Сот_Доб_Обяз_зап_label3.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Сотрудники (Код_должности, ФИО, Пол, Возраст, Адрес, Паспортные_данные, Примечание) VALUES (@Код_должности, @ФИО, @Пол, @Возраст, @Адрес, @Паспортные_данные, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("Код_должности", код_должностиComboBox3.Text);
                command.Parameters.AddWithValue("ФИО", фИОTextBox1.Text);
                command.Parameters.AddWithValue("Пол", полComboBox1.Text);
                command.Parameters.AddWithValue("Возраст", возрастComboBox1.Text);
                command.Parameters.AddWithValue("Адрес", адресTextBox1.Text);
                command.Parameters.AddWithValue("Паспортные_данные", паспортные_данныеTextBox1.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox9.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Сотрудники",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Сотрудники", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                СотрудникиDataGrid.DataSource = table;
                код_сотрудникаComboBox6.DataSource = table;
                код_сотрудникаComboBox7.DataSource = table;
            }
        }
        private void TableLayoutPanel15_MouseMove(object sender, MouseEventArgs e)
        {
            Сот_Доб_Обяз_зап_label.Hide();
            Сот_Доб_Обяз_зап_label1.Hide();
            Сот_Доб_Обяз_зап_label2.Hide();
            Сот_Доб_Обяз_зап_label3.Hide();
        }
        private void ФИОTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != ' ')
            {
                e.Handled = true;
            }

        }

        private void Паспортные_данныеTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        //Сот_Удал_TabPage
        private async void УдалитьButton5_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM Сотрудники WHERE Код_сотрудника=@Код_сотрудника", sqlConnection);
            command.Parameters.AddWithValue("Код_сотрудника", код_сотрудникаComboBox7.Text);
            popup = new PopupNotifier
            {
                Image = Properties.Resources.connected,
                ImageSize = new Size(96, 96),
                TitleText = "Сотрудники",
                ContentText = "Данные успешно удалены!"
            };
            popup.Popup();
            await command.ExecuteNonQueryAsync();
            adapter = new SqlDataAdapter("SELECT * FROM Сотрудники", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            СотрудникиDataGrid.DataSource = table;
            код_сотрудникаComboBox6.DataSource = table;
            код_сотрудникаComboBox7.DataSource = table;
        }

        ////НАЧИСЛЕНИЕ ЗП
        //Нач_Обн_TabPage
        private async void ОбновитьButton2_Click(object sender, EventArgs e)
        {
            if (сумма_выплатыTextBox.Text == "" || this.размер_премииTextBox.Text == "" || this.дата_выплатыDateTimePicker.Text == "")
            {
                Нач_Обн_Обяз_зап_label.Show();
                Нач_Обн_Обяз_зап_label1.Show();
                Нач_Обн_Обяз_зап_label2.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("UPDATE Начисление_ЗП SET Код_сотрудника=@Код_сотрудника, Сумма_выплаты=@Сумма_выплаты, Размер_премии=@Размер_премии, Дата_выплаты=@Дата_выплаты, Статус=@Статус, Примечание=@Примечание WHERE Код_начисления=@Код_начисления", sqlConnection);
                command.Parameters.AddWithValue("Код_начисления", код_начисленияComboBox.Text);
                command.Parameters.AddWithValue("Код_сотрудника", код_сотрудникаComboBox2.Text);
                command.Parameters.AddWithValue("Сумма_выплаты", сумма_выплатыTextBox.Text);
                command.Parameters.AddWithValue("Размер_премии", размер_премииTextBox.Text);
                command.Parameters.AddWithValue("Дата_выплаты", дата_выплатыDateTimePicker.Text);
                command.Parameters.AddWithValue("Статус", статусComboBox.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox4.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Начисление ЗП",
                    ContentText = "Данные успешно обновлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Начисление_ЗП", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                Начисление_зпDataGrid.DataSource = table;
            }
        }
        private void TableLayoutPanel8_MouseMove(object sender, MouseEventArgs e)
        {
            Нач_Обн_Обяз_зап_label.Hide();
            Нач_Обн_Обяз_зап_label1.Hide();
            Нач_Обн_Обяз_зап_label2.Hide();
        }
        private void Сумма_выплатыTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void Размер_премииTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }
        //Нач_Доб_TabPage
        private async void ДобавитьButton2_Click(object sender, EventArgs e)
        {
            if (сумма_выплатыTextBox1.Text == "" || this.размер_премииTextBox1.Text == "" || this.дата_выплатыDateTimePicker1.Text == "")
            {
                Нач_Доб_Обяз_зап_label.Show();
                Нач_Доб_Обяз_зап_label1.Show();
                Нач_Доб_Обяз_зап_label2.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Начисление_ЗП (Код_сотрудника, Сумма_выплаты, Размер_премии, Дата_выплаты, Статус, Примечание) VALUES (@Код_сотрудника, @Сумма_выплаты, @Размер_премии, @Дата_выплаты, @Статус, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("Код_сотрудника", код_сотрудникаComboBox3.Text);
                command.Parameters.AddWithValue("Сумма_выплаты", сумма_выплатыTextBox1.Text);
                command.Parameters.AddWithValue("Размер_премии", размер_премииTextBox1.Text);
                command.Parameters.AddWithValue("Дата_выплаты", дата_выплатыDateTimePicker1.Text);
                command.Parameters.AddWithValue("Статус", статусComboBox1.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox5.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Начисление ЗП",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Начисление_ЗП", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                Начисление_зпDataGrid.DataSource = table;
                код_начисленияComboBox.DataSource = table;
                код_начисленияComboBox1.DataSource = table;
            }
        }
        private void TableLayoutPanel9_MouseMove(object sender, MouseEventArgs e)
        {
            Нач_Доб_Обяз_зап_label.Hide();
            Нач_Доб_Обяз_зап_label1.Hide();
            Нач_Доб_Обяз_зап_label2.Hide();
        }
        private void Сумма_выплатыTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void Размер_премииTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        //Нач_Доб_TabPage
        private async void УдалитьButton2_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM Начисление_ЗП WHERE Код_начисления=@Код_начисления", sqlConnection);
            command.Parameters.AddWithValue("Код_начисления", код_начисленияComboBox1.Text);
            popup = new PopupNotifier
            {
                Image = Properties.Resources.connected,
                ImageSize = new Size(96, 96),
                TitleText = "Начисление ЗП",
                ContentText = "Данные успешно удалены!"
            };
            popup.Popup();
            await command.ExecuteNonQueryAsync();
            adapter = new SqlDataAdapter("SELECT * FROM Начисление_ЗП", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            Начисление_зпDataGrid.DataSource = table;
            код_начисленияComboBox.DataSource = table;
            код_начисленияComboBox1.DataSource = table;
        }

        ////ОТПУСКИ
        //Отп_Обн_TabPage
        private async void ОбновитьButton3_Click(object sender, EventArgs e)
        {
            if (дата_начала_отпускаDateTimePicker.Text == "" || this.дата_окончания_отпускаDateTimePicker.Text == "" || this.длительностьTextBox2.Text == "" || this.видComboBox.Text == "")
            {
                Отп_Обн_Обяз_зап_label.Show();
                Отп_Обн_Обяз_зап_label1.Show();
                Отп_Обн_Обяз_зап_label2.Show();
                Отп_Обн_Обяз_зап_label3.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("UPDATE Отпуски SET Код_сотрудника=@Код_сотрудника, Дата_начала_отпуска=@Дата_начала_отпуска, Дата_окончания_отпуска=@Дата_окончания_отпуска, Длительность=@Длительность, Вид=@Вид, Выплата=@Выплата, Примечание=@Примечание WHERE Код_отпуска=@Код_отпуска", sqlConnection);
                command.Parameters.AddWithValue("Код_отпуска", код_отпускаComboBox.Text);
                command.Parameters.AddWithValue("Код_сотрудника", код_сотрудникаСomboBox4.Text);
                command.Parameters.AddWithValue("Дата_начала_отпуска", дата_начала_отпускаDateTimePicker.Text);
                command.Parameters.AddWithValue("Дата_окончания_отпуска", дата_окончания_отпускаDateTimePicker.Text);
                command.Parameters.AddWithValue("Длительность", длительностьTextBox2.Text);
                command.Parameters.AddWithValue("Вид", видComboBox.Text);
                command.Parameters.AddWithValue("Выплата", выплатаTextBox.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox6.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Отпуски",
                    ContentText = "Данные успешно обновлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Отпуски", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                ОтпускиDataGrid.DataSource = table;
            }
        }
        private void TableLayoutPanel11_MouseMove(object sender, MouseEventArgs e)
        {
            Отп_Обн_Обяз_зап_label.Hide();
            Отп_Обн_Обяз_зап_label1.Hide();
            Отп_Обн_Обяз_зап_label2.Hide();
            Отп_Обн_Обяз_зап_label3.Hide();
        }
        private void Дата_окончания_отпускаDateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt1 = дата_начала_отпускаDateTimePicker1.Value;
            DateTime dt2 = дата_окончания_отпускаDateTimePicker1.Value;
            TimeSpan x = dt2 - dt1;
            длительностьTextBox3.Text = ((int)x.TotalDays).ToString() + " дней";
        }
        private void ВыплатаTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < '0' || l > '9') && l != '\b' && l != ',' && l != ' ')
            {
                e.Handled = true;
            }
        }
        //Отп_Доб_TabPage
        private async void ДобавитьButton3_Click(object sender, EventArgs e)
        {
            if (дата_начала_отпускаDateTimePicker1.Text == "" || this.дата_окончания_отпускаDateTimePicker1.Text == "" || this.длительностьTextBox3.Text == "" || this.видComboBox1.Text == "")
            {
                Отп_Доб_Обяз_зап_label.Show();
                Отп_Доб_Обяз_зап_label1.Show();
                Отп_Доб_Обяз_зап_label2.Show();
                Отп_Доб_Обяз_зап_label3.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Отпуски (Код_сотрудника, Дата_начала_отпуска, Дата_окончания_отпуска, Длительность, Вид, Выплата, Примечание) VALUES (@Код_сотрудника, @Дата_начала_отпуска, @Дата_окончания_отпуска, @Длительность, @Вид, @Выплата, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("Код_сотрудника", код_сотрудникаComboBox5.Text);
                command.Parameters.AddWithValue("Дата_начала_отпуска", дата_начала_отпускаDateTimePicker1.Text);
                command.Parameters.AddWithValue("Дата_окончания_отпуска", дата_окончания_отпускаDateTimePicker1.Text);
                command.Parameters.AddWithValue("Длительность", длительностьTextBox3.Text);
                command.Parameters.AddWithValue("Вид", видComboBox1.Text);
                command.Parameters.AddWithValue("Выплата", выплатаTextBox1.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox7.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Отпуски",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Отпуски", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                ОтпускиDataGrid.DataSource = table;
                код_отпускаComboBox.DataSource = table;
                код_отпускаComboBox1.DataSource = table;
            }
        }
        private void TableLayoutPanel12_MouseMove(object sender, MouseEventArgs e)
        {
            Отп_Доб_Обяз_зап_label.Hide();
            Отп_Доб_Обяз_зап_label1.Hide();
            Отп_Доб_Обяз_зап_label2.Hide();
            Отп_Доб_Обяз_зап_label3.Hide();
        }
        private void Дата_окончания_отпускаDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt1 = дата_начала_отпускаDateTimePicker.Value;
            DateTime dt2 = дата_окончания_отпускаDateTimePicker.Value;
            TimeSpan x = dt2 - dt1;
            длительностьTextBox2.Text = ((int)x.TotalDays).ToString() + " дней";
        }
        private void ВыплатаTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < '0' || l > '9') && l != '\b' && l != ',' && l != ' ')
            {
                e.Handled = true;
            }
        }
        //Отп_Удал_TabPage
        private async void УдалитьButton3_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM Отпуски WHERE Код_отпуска=@Код_отпуска", sqlConnection);
            command.Parameters.AddWithValue("Код_отпуска", код_отпускаComboBox.Text);
            popup = new PopupNotifier
            {
                Image = Properties.Resources.connected,
                ImageSize = new Size(96, 96),
                TitleText = "Отпуски",
                ContentText = "Данные успешно удалены!"
            };
            popup.Popup();
            await command.ExecuteNonQueryAsync();
            adapter = new SqlDataAdapter("SELECT * FROM Отпуски", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            ОтпускиDataGrid.DataSource = table;
            код_отпускаComboBox.DataSource = table;
            код_отпускаComboBox1.DataSource = table;
        }

        ////ТРУДОВЫЕ ДОГОВОРА
        //Труд_Обн_TabPage
        private async void ОбновитьButton6_Click(object sender, EventArgs e)
        {
            if (дата_заключенияDateTimePicker.Text == "" || this.длительностьComboBox4.Text == "")
            {
                Труд_Обн_Обяз_зап_label.Show();
                Труд_Обн_Обяз_зап_label1.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("UPDATE Трудовые_договора SET Код_сотрудника=@Код_сотрудника, Дата_заключения=@Дата_заключения, Длительность=@Длительность, Примечание=@Примечание WHERE Код_договора=@Код_договора", sqlConnection);
                command.Parameters.AddWithValue("Код_договора", код_договораComboBox.Text);
                command.Parameters.AddWithValue("Код_сотрудника", код_сотрудникаComboBox8.Text);
                command.Parameters.AddWithValue("Дата_заключения", дата_заключенияDateTimePicker.Text);
                command.Parameters.AddWithValue("Длительность", длительностьComboBox4.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox10.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Трудовые договора",
                    ContentText = "Данные успешно обновлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Трудовые_договора", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                Трудовые_договораDataGrid.DataSource = table;
            }
        }
        private void TableLayoutPanel17_MouseMove(object sender, MouseEventArgs e)
        {
            Труд_Обн_Обяз_зап_label.Hide();
            Труд_Обн_Обяз_зап_label1.Hide();
        }

        //Труд_Доб_TabPage
        private async void ДобавитьButton6_Click(object sender, EventArgs e)
        {
            if (дата_заключенияDateTimePicker1.Text == "" || this.длительностьComboBox5.Text == "")
            {
                Труд_Доб_Обяз_зап_label.Show();
                Труд_Доб_Обяз_зап_label1.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Трудовые_договора (Код_сотрудника, Дата_заключения, Длительность, Примечание) VALUES (@Код_сотрудника, @Дата_заключения, @Длительность, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("Код_сотрудника", код_сотрудникаComboBox9.Text);
                command.Parameters.AddWithValue("Дата_заключения", дата_заключенияDateTimePicker1.Text);
                command.Parameters.AddWithValue("Длительность", длительностьComboBox5.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox11.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Трудовые договора",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Трудовые_договора", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                Трудовые_договораDataGrid.DataSource = table;
                код_договораComboBox.DataSource = table;
                код_договораComboBox1.DataSource = table;
            }
        }
        private void TableLayoutPanel18_MouseMove(object sender, MouseEventArgs e)
        {
            Труд_Доб_Обяз_зап_label.Hide();
            Труд_Доб_Обяз_зап_label1.Hide();
        }

        //Труд_Удал_TabPage
        private async void УдалитьButton6_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM Трудовые_договора WHERE Код_договора=@Код_договора", sqlConnection);
            command.Parameters.AddWithValue("Код_договора", код_договораComboBox1.Text);
            popup = new PopupNotifier
            {
                Image = Properties.Resources.connected,
                ImageSize = new Size(96, 96),
                TitleText = "Трудовые договора",
                ContentText = "Данные успешно удалены!"
            };
            popup.Popup();
            await command.ExecuteNonQueryAsync();
            adapter = new SqlDataAdapter("SELECT * FROM Трудовые_договора", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            Трудовые_договораDataGrid.DataSource = table;
            код_договораComboBox.DataSource = table;
            код_договораComboBox1.DataSource = table;
        }

        ////ШТАТНОЕ РАСПИСАНИЕ
        //Штат_Обн_TabPage
        private async void ОбновитьButton7_Click(object sender, EventArgs e)
        {
            if (наименование_структурного_подразделенияTextBox.Text == "" || this.количество_штатных_единицTextBox.Text == "")
            {
                Шт_Обн_Обяз_зап_label.Show();
                Шт_Обн_Обяз_зап_label1.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("UPDATE Штатное_расписание SET Код_должности=@Код_должности, Наименование_структурного_подразделения=@Наименование_структурного_подразделения, Количество_штатных_единиц=@Количество_штатных_единиц, Примечание=@Примечание WHERE Код_расписания=@Код_расписания", sqlConnection);
                command.Parameters.AddWithValue("Код_расписания", код_расписанияComboBox.Text);
                command.Parameters.AddWithValue("Код_должности", код_должностиComboBox4.Text);
                command.Parameters.AddWithValue("Наименование_структурного_подразделения", наименование_структурного_подразделенияTextBox.Text);
                command.Parameters.AddWithValue("Количество_штатных_единиц", количество_штатных_единицTextBox.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox12.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Штатное расписание",
                    ContentText = "Данные успешно обновлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Штатное_расписание", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                Штатное_расписаниеDataGrid.DataSource = table;
            }
        }

        private void TableLayoutPanel20_MouseMove(object sender, MouseEventArgs e)
        {
            Шт_Обн_Обяз_зап_label.Hide();
            Шт_Обн_Обяз_зап_label1.Hide();
        }

        private void Наименование_структурного_подразделенияTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && (l < '0' || l > '9') && l != '\b' && l != '.' && l != ',' && l != ' ' && l != '"')
            {
                e.Handled = true;
            }
        }

        private void Количество_штатных_единицTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        //Штат_Доб_TabPage
        private async void ДобавитьButton7_Click(object sender, EventArgs e)
        {
            if (наименование_структурного_подразделенияTextBox1.Text == "" || this.количество_штатных_единицTextBox1.Text == "")
            {
                Шт_Доб_Обяз_зап_label.Show();
                Шт_Доб_Обяз_зап_label1.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Штатное_расписание (Код_должности, Наименование_структурного_подразделения, Количество_штатных_единиц, Примечание) VALUES (@Код_должности, @Наименование_структурного_подразделения, @Количество_штатных_единиц, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("Код_должности", код_должностиComboBox5.Text);
                command.Parameters.AddWithValue("Наименование_структурного_подразделения", наименование_структурного_подразделенияTextBox1.Text);
                command.Parameters.AddWithValue("Количество_штатных_единиц", количество_штатных_единицTextBox1.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox13.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Штатное расписание",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Штатное_расписание", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                Штатное_расписаниеDataGrid.DataSource = table;
                код_расписанияComboBox.DataSource = table;
                код_расписанияComboBox1.DataSource = table;
            }
        }

        private void TableLayoutPanel21_MouseMove(object sender, MouseEventArgs e)
        {
            Шт_Доб_Обяз_зап_label.Hide();
            Шт_Доб_Обяз_зап_label1.Hide();
        }

        private void Наименование_структурного_подразделенияTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && (l < '0' || l > '9') && l != '\b' && l != '.' && l != ',' && l != ' ' && l != '"')
            {
                e.Handled = true;
            }
        }

        private void Количество_штатных_единицTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _ = e.KeyChar;
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        //Штат_Удал_TabPage
        private async void УдалитьButton7_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM Штатное_расписание WHERE Код_расписания=@Код_расписания", sqlConnection);
            command.Parameters.AddWithValue("Код_расписания", код_расписанияComboBox1.Text);
            popup = new PopupNotifier
            {
                Image = Properties.Resources.connected,
                ImageSize = new Size(96, 96),
                TitleText = "Штатное расписание",
                ContentText = "Данные успешно удалены!"
            };
            popup.Popup();
            await command.ExecuteNonQueryAsync();
            adapter = new SqlDataAdapter("SELECT * FROM Штатное_расписание", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            Штатное_расписаниеDataGrid.DataSource = table;
            код_расписанияComboBox.DataSource = table;
            код_расписанияComboBox1.DataSource = table;
        }

        ////ПОЛЬЗОВАТЕЛИ
        //Пол_Обн_TabPage
        private async void ОбновитьButton8_Click(object sender, EventArgs e)
        {
            if (логинTextBox.Text == "" || this.парольTextBox.Text == "")
            {
                Пол_Обн_Обяз_зап_label.Show();
                Пол_Обн_Обяз_зап_label1.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("UPDATE Пользователи SET Логин=@Логин, ФИО=@ФИО, Роль=@Роль, Пароль=@Пароль," +
                    " Email=@Email, Телефон=@Телефон, Примечание=@Примечание WHERE Код_пользователя=@Код_пользователя", sqlConnection);
                command.Parameters.AddWithValue("Код_пользователя", код_пользователяComboBox.Text);
                command.Parameters.AddWithValue("Логин", логинTextBox.Text);
                command.Parameters.AddWithValue("ФИО", фИОTextBox2.Text);
                command.Parameters.AddWithValue("Роль", рольTextBox.Text);
                command.Parameters.AddWithValue("Пароль", парольTextBox.Text);
                command.Parameters.AddWithValue("Email", emailTextBox.Text);
                command.Parameters.AddWithValue("Телефон", телефонTextBox.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox14.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Пользователи",
                    ContentText = "Данные успешно обновлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();
                adapter = new SqlDataAdapter("SELECT * FROM Пользователи", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                ПользователиDataGrid.DataSource = table;
            }
        }

        private void TableLayoutPanel23_MouseMove(object sender, MouseEventArgs e)
        {
            Пол_Обн_Обяз_зап_label.Hide();
            Пол_Обн_Обяз_зап_label1.Hide();
        }

        private void ЛогинTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'A' || l > 'z') && (l < '0' || l > '9') && l != '\b' && l != '@')
            {
                e.Handled = true;
            }
        }

        private void ПарольTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'A' || l > 'z') && (l < '0' || l > '9') && l != '\b' && l != '@' && l != '%' && l != '$' && l != '&')
            {
                e.Handled = true;
            }
        }

        //Пол_Доб_TabPage
        private async void ДобавитьButton8_Click(object sender, EventArgs e)
        {
            if (this.логинTextBox1.Text == "" || this.парольTextBox1.Text == "")
            {
                Пол_Доб_Обяз_зап_label.Show();
                Пол_Доб_Обяз_зап_label1.Show();
            }
            else
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
                sqlConnection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Пользователи (Логин, ФИО, Роль, Пароль, Email, Телефон, Примечание) VALUES (@Логин, @ФИО, @Роль, @Пароль, @Email, @Телефон, @Примечание)", sqlConnection);
                command.Parameters.AddWithValue("Логин", логинTextBox1.Text);
                command.Parameters.AddWithValue("ФИО", фИОTextBox3.Text);
                command.Parameters.AddWithValue("Роль", рольTextBox1.Text);
                command.Parameters.AddWithValue("Пароль", парольTextBox1.Text);
                command.Parameters.AddWithValue("Email", emailTextBox1.Text);
                command.Parameters.AddWithValue("Телефон", телефонTextBox1.Text);
                command.Parameters.AddWithValue("Примечание", примечаниеTextBox15.Text);
                popup = new PopupNotifier
                {
                    Image = Properties.Resources.connected,
                    ImageSize = new Size(96, 96),
                    TitleText = "Пользователи",
                    ContentText = "Данные успешно добавлены!"
                };
                popup.Popup();
                await command.ExecuteNonQueryAsync();

                adapter = new SqlDataAdapter("SELECT * FROM Пользователи", sqlConnection);
                table = new DataTable();
                adapter.Fill(table);
                ПользователиDataGrid.DataSource = table;
                код_пользователяComboBox.DataSource = table;
                код_пользователяComboBox1.DataSource = table;
            }
        }

        private void TableLayoutPanel24_MouseMove(object sender, MouseEventArgs e)
        {
            Пол_Доб_Обяз_зап_label.Hide();
            Пол_Доб_Обяз_зап_label1.Hide();
        }

        private void ЛогинTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'A' || l > 'z') && (l < '0' || l > '9') && l != '\b' && l != '@')
            {
                e.Handled = true;
            }
        }

        private void ПарольTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'A' || l > 'z') && (l < '0' || l > '9') && l != '\b' && l != '@' && l != '%' && l != '$' && l != '&')
            {
                e.Handled = true;
            }
        }

        //Пол_Удал_TabPage
        private async void УдалитьButton8_Click(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["AccountingOfSalariesAndPersonnel.Properties.Settings.AccountingOfSalariesAndPersonnelConnectionString"].ConnectionString);
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM Пользователи WHERE Код_пользователя=@Код_пользователя", sqlConnection);
            command.Parameters.AddWithValue("Код_пользователя", код_пользователяComboBox1.Text);
            popup = new PopupNotifier
            {
                Image = Properties.Resources.connected,
                ImageSize = new Size(96, 96),
                TitleText = "Пользователи",
                ContentText = "Данные успешно удалены!"
            };
            popup.Popup();
            await command.ExecuteNonQueryAsync();
            adapter = new SqlDataAdapter("SELECT * FROM Пользователи", sqlConnection);
            table = new DataTable();
            adapter.Fill(table);
            ПользователиDataGrid.DataSource = table;
            код_пользователяComboBox.DataSource = table;
            код_пользователяComboBox1.DataSource = table;
        }

        private void ОПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            О_программе о_программе = new О_программе();
            о_программе.Show();
        }
    }
}
