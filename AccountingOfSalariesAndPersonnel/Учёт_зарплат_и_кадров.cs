using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AccountingOfSalariesAndPersonnel
{
    public partial class Учёт_зарплат_и_кадров : Form
    {
        public Учёт_зарплат_и_кадров()
        {
            InitializeComponent();
        }

        private void Учёт_зарплат_и_кадров_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "accountingOfSalariesAndPersonnelDataSet.Должности". При необходимости она может быть перемещена или удалена.
            this.должностиTableAdapter.Fill(this.accountingOfSalariesAndPersonnelDataSet.Должности);

        }
    }
}
