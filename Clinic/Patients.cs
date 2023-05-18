using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.Remoting.Contexts;

namespace Clinic
{
    public partial class Patients : Form
    {
        private string Query = "Select * from PatientTbl";
        public Patients()
        {
            InitializeComponent();
            DisplayPat(Query);
            if (Login.Role == "Dcotor")
            {
                LabLbl.Enabled = false;
                DocLbl.Enabled = false;
                RecLbl.Enabled = false;
            }else if (Login.Role == "Receptionist")
            {
                LabLbl.Enabled = false;
                DocLbl.Enabled = false;
                RecLbl.Enabled = false;
            }
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;

        }
   
        int Key = 0;
        SqlConnection Con = new SqlConnection(@"workstation id=clinicdanyadb.mssql.somee.com;packet size=4096;user id=danyafrontend_SQLLogin_2;pwd=77s976d7og;data source=clinicdanyadb.mssql.somee.com;persist security info=False;initial catalog=clinicdanyadb");
        private void DisplayPat(string Query)
        {
            // получаем данные в DataSet через SqlDataAdapter
            Con.Open();
    
            SqlDataAdapter sda = new SqlDataAdapter(Query, Con);
            SqlCommandBuilder builder = new SqlCommandBuilder(sda);
            //создаем объект DataSet
            var ds = new DataSet();
            //заполняем DataSet
            sda.Fill(ds);
            //отображаем данные
            PatientsDGV.DataSource = ds.Tables[0];
            Con.Close();
        }
        private void Clear()
        {
            //очищаем все значения
            PatNameTB.Text = "";
            PatGenCB.SelectedIndex = 0;
            PatAddTb.Text = "";
            PatPhone.Text = "";
            PatAlTb.Text = "";
            PatHIVCB.SelectedIndex = 0;
            Key = 0;
        }
        private void AddBtn_Click(object sender, EventArgs e)
        {

            //проверка на пустые значения
            if (PatNameTB.Text == "" || PatPhone.Text == "" || PatAlTb.Text == "" || PatAddTb.Text == "" || PatGenCB.SelectedIndex == -1 || PatHIVCB.SelectedIndex == -1)
            {
                MessageBox.Show("Заполните все поля");
            }
            else
            {
                //отлавливаем ошибку
                try
                {
                    //открываем подключение
                    Con.Open();
                    //создаем команду для добавления
                    SqlCommand cmd = new SqlCommand("insert into PatientTbl(PatName,PatGen,PatDOB,PatAdd,PatPhone,PatHIV,PatAll)values(@PN,@PG,@PD,@PA,@PP,@PH,@PAl)", Con);
                    //добавляем значения в строку
                    cmd.Parameters.AddWithValue("@PN", PatNameTB.Text);
                    cmd.Parameters.AddWithValue("@PG", PatGenCB.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@PD", PatDOB.Value.Date);
                    cmd.Parameters.AddWithValue("@PA", PatAddTb.Text);
                    cmd.Parameters.AddWithValue("@PP", PatPhone.Text);
                    cmd.Parameters.AddWithValue("@PH", PatHIVCB.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@PAl", PatAlTb.Text);

                    cmd.ExecuteNonQuery();
                    // выполняет sql-выражение и возвращает количество измененных записей.
                    MessageBox.Show("Пациент добавлен");
                    //закрываем подключение
                    Con.Close();
                    DisplayPat(Query);
                    Clear();
                }
                catch (Exception err)
                {
                    //показываем ошибку
                    MessageBox.Show(err.Message);
                }
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

            Close();
        }

        private void PatientsDGV_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //заполняем данные
            PatNameTB.Text = PatientsDGV.SelectedRows[0].Cells[1].Value.ToString();
            PatGenCB.SelectedItem = PatientsDGV.SelectedRows[0].Cells[2].Value.ToString();
            PatDOB.Text = PatientsDGV.SelectedRows[0].Cells[3].Value.ToString();
            PatAddTb.Text = PatientsDGV.SelectedRows[0].Cells[4].Value.ToString();
            PatPhone.Text = PatientsDGV.SelectedRows[0].Cells[5].Value.ToString();
            PatHIVCB.SelectedItem = PatientsDGV.SelectedRows[0].Cells[6].Value.ToString();
            PatAlTb.Text = PatientsDGV.SelectedRows[0].Cells[7].Value.ToString();
    
          
            //Сохраняем ключ для обновления(id секретаря)
            if (PatNameTB.Text == "")
            {
                Key = 0;
            }
            else
            {
                Key = Convert.ToInt32(PatientsDGV.SelectedRows[0].Cells[0].Value.ToString());
            }
        }

        private void EditBtn_Click(object sender, EventArgs e)
        {
            //проверка на пустые значения
            if (PatNameTB.Text == "" || PatPhone.Text == "" || PatAlTb.Text == "" || PatAddTb.Text == "" || PatGenCB.SelectedIndex == -1 || PatHIVCB.SelectedIndex == -1)
            {
                MessageBox.Show("Заполните все поля");
            }
            else
            {
                //отлавливаем ошибку
                try
                {
                    //открываем подключение
                    Con.Open();
                    //создаем команду для добавления
                    SqlCommand cmd = new SqlCommand("Update PatientTbl set PatName=@PN,PatGen=@PG,PatDOB=@PD,PatAdd=@PA,PatPhone=@PP,PatHIV=@PH,PatAll=@PAl where PatId=@PKey", Con);
                    //добавляем значения в строку
                    cmd.Parameters.AddWithValue("@PN", PatNameTB.Text);
                    cmd.Parameters.AddWithValue("@PG", PatGenCB.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@PD", PatDOB.Value.Date);
                    cmd.Parameters.AddWithValue("@PA", PatAddTb.Text);
                    cmd.Parameters.AddWithValue("@PP", PatPhone.Text);
                    cmd.Parameters.AddWithValue("@PH", PatHIVCB.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@PAl", PatAddTb.Text);
                    cmd.Parameters.AddWithValue("@PKey", Key);
                    cmd.ExecuteNonQuery();
                    // выполняет sql-выражение и возвращает количество измененных записей.
                    MessageBox.Show("Пациент обновлен");
                    //закрываем подключение
                    Con.Close();
                    DisplayPat(Query);
                    Clear();
                }
                catch (Exception err)
                {
                    //показываем ошибку
                    MessageBox.Show(err.Message);
                }
            }
        }

        private void DelBtn_Click(object sender, EventArgs e)
        {
            //если не выбран никто
            if (Key == 0)
            {
                MessageBox.Show("Выберите пациента");
            }
            else
            {
                //отлавливаем ошибку
                try
                {
                    //открываем подключение
                    Con.Open();
                    //создаем команду для редактирования
                    SqlCommand cmd = new SqlCommand("Delete from PatientTbl where PatId=@PKey", Con);
                    //добавляем значения в строку
                    cmd.Parameters.AddWithValue("@PKey", Key);
                    cmd.ExecuteNonQuery();
                    // выполняет sql-выражение и возвращает количество измененных записей.
                    MessageBox.Show("Пациент удален");
                    //закрываем подключение
                    Con.Close();
                    DisplayPat(Query);
                    Clear();
                }
                catch (Exception err)
                {
                    //показываем ошибку
                    MessageBox.Show(err.Message);
                }
            }
        }

        private void label10_Click(object sender, EventArgs e)
        {
            Homes obj = new Homes();
            obj.Show();
            this.Hide();
        }

        private void label14_Click(object sender, EventArgs e)
        {
            Login obj = new Login();
            obj.Show();
            this.Hide();
        }

        private void DocLbl_Click(object sender, EventArgs e)
        {
            Doctors obj = new Doctors();
            obj.Show();
            this.Hide();
        }

        private void LabLbl_Click(object sender, EventArgs e)
        {
            LabTests obj = new LabTests();
            obj.Show();
     
        }

        private void RecLbl_Click(object sender, EventArgs e)
        {
            Receptionists obj = new Receptionists();
            obj.Show();
            this.Hide();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < PatientsDGV.RowCount; i++)
            {
                for (int j = 0; j < PatientsDGV.ColumnCount; j++)
                {
                    PatientsDGV[j, i].Style.BackColor = Color.White;
                    PatientsDGV[j, i].Style.ForeColor = Color.Black;
                }
            }

            if (!string.IsNullOrWhiteSpace(textBox1.Text))
            {
                PatientsDGV.ClearSelection();
                for (int i = 0; i < PatientsDGV.RowCount - 1; i++)
                {
                    for (int j = 0; j < PatientsDGV.ColumnCount ; j++)
                    {
                        if (PatientsDGV[j, i].Value.ToString().ToLower().Contains(textBox1.Text.ToLower()))
                        {
                            PatientsDGV[j, i].Style.BackColor = Color.Black;
                            PatientsDGV[j, i].Style.ForeColor = Color.White;
                        }
                    }
                }
            }
        }
   

        private Excel.Application GetExcel()
        {
            Excel.Application xlApp;
            Worksheet xlSheet;
            xlApp = new Excel.Application();
            Excel.Workbook wBook;
            wBook = xlApp.Workbooks.Add();
            xlApp.Columns.ColumnWidth = 15;
            xlSheet = wBook.Sheets[1];
            xlSheet.Name = "Покупатели";
            xlSheet.Cells.HorizontalAlignment = 3;
            for (int j = 1; j < PatientsDGV.Columns.Count + 1; j++)
            {
                xlApp.Cells[1, j] = PatientsDGV.Columns[j - 1].HeaderText;
            }
            return xlApp;
        }

        private void ExpBtn_Click(object sender, EventArgs e)
        {
            switch (MessageBox.Show("Экспортировать все?", "Справка", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question))
            {
                case DialogResult.Yes:
                    {
                        var xlApp = GetExcel();
                        for (int i = 0; i < PatientsDGV.Rows.Count - 1; i++)
                        {
                            for (int j = 0; j < PatientsDGV.Columns.Count ; j++)
                            {
                                xlApp.Cells[i + 2, j + 1] = PatientsDGV.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                        xlApp.Visible = true;
                        break;
                    };

                case DialogResult.No:
                    {
                        var xlApp = GetExcel();
                        for (int i = 0; i < PatientsDGV.SelectedRows.Count -1; i++)
                        {
                            for (int j = 0; j < PatientsDGV.Columns.Count ; j++)
                            {
                                xlApp.Cells[i + 2, j + 1] = PatientsDGV.SelectedRows[i].Cells[j].Value.ToString();
                            }
                        }
                        xlApp.Visible = true;
                        break;
                    };

                case DialogResult.Cancel:
                    return;
            }
        }

        private void SpLbl_Click(object sender, EventArgs e)
        {
            About obj = new About();
            obj.ShowDialog();
        }

        private void buttonSort_Click(object sender, EventArgs e)
        {

            var COL = new System.Windows.Forms.DataGridViewColumn();

            switch (comboBox1.SelectedItem.ToString())
            {
                case "Адрес":
                    COL = PatientsDGV.Columns["PatAdd"];
                    break;
                case "Дата":
                    COL = PatientsDGV.Columns["PatDOB"];
                    break;
                case "Имя":
                   
                    COL = PatientsDGV.Columns["PatName"];
                    break;
            }
            if (radioButtonUp.Checked)
            {
                PatientsDGV.Sort(COL, System.ComponentModel.ListSortDirection.Ascending);
            }
            else
            {
                PatientsDGV.Sort(COL, System.ComponentModel.ListSortDirection.Descending);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DisplayPat(Query);
            buttonSort.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
                    DisplayPat(Query + " where PatGen=N'" + comboBox2.SelectedItem.ToString() + "'");
            buttonSort.Enabled = false;
        }
    }
}
