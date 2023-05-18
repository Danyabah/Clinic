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
using System.Collections;

namespace Clinic
{
    public partial class Doctors : Form
    {
        private string Query = "Select * from DoctorTbl";
        public Doctors()
        {
            InitializeComponent();
            DisplayDoc(Query);
        }
        int Key = 0;
        SqlConnection Con = new SqlConnection(@"workstation id=clinicdanyadb.mssql.somee.com;packet size=4096;user id=danyafrontend_SQLLogin_2;pwd=77s976d7og;data source=clinicdanyadb.mssql.somee.com;persist security info=False;initial catalog=clinicdanyadb");
        private void DisplayDoc(string Query)
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
            DoctorsDGV.DataSource = ds.Tables[0];
            Con.Close();
        }
        private void Clear()
        {
            //очищаем все значения
            DocNameTb.Text = "";
            DocPhoneTb.Text = "";
            DocAddTb.Text = "";
            DocExpTb.Text = "";
            DocPassWordTb.Text = "";
            DocGenCb.SelectedIndex = 0;
            DocSpecCb.SelectedIndex = 0;
            Key = 0;
        }
        private void AddBtn_Click(object sender, EventArgs e)
        {
            //проверка на пустые значения
            if (DocNameTb.Text == "" || DocPassWordTb.Text == "" || DocPhoneTb.Text == "" || DocAddTb.Text == "" || DocGenCb.SelectedIndex == -1 || DocSpecCb.SelectedIndex == -1)
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
                    SqlCommand cmd = new SqlCommand("insert into DoctorTbl(DocName,DOCDOB,DOCGEN,DOCSPEC,DOCEXP,DOCPHONE,DOCADD,DOCPASS)values(@DN,@DD,@DG,@DS,@DE,@DP,@DA,@DPA)", Con);
                    //добавляем значения в строку
                    cmd.Parameters.AddWithValue("@DN", DocNameTb.Text);
                    cmd.Parameters.AddWithValue("@DD", DocDOB.Value.Date);
                    cmd.Parameters.AddWithValue("@DG", DocGenCb.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@DS", DocSpecCb.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@DE", DocExpTb.Text);
                    cmd.Parameters.AddWithValue("@DP", DocPhoneTb.Text);
                    cmd.Parameters.AddWithValue("@DA", DocAddTb.Text);
                    cmd.Parameters.AddWithValue("@DPA", DocPassWordTb.Text);
                    cmd.ExecuteNonQuery();
                    // выполняет sql-выражение и возвращает количество измененных записей.
                    MessageBox.Show("Доктор добавлен");
                    //закрываем подключение
                    Con.Close();
                    DisplayDoc(Query);
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

        private void DoctorsDGV_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //заполняем данные
            DocNameTb.Text = DoctorsDGV.SelectedRows[0].Cells[1].Value.ToString();
            DocDOB.Text = DoctorsDGV.SelectedRows[0].Cells[2].Value.ToString();
            DocGenCb.SelectedItem = DoctorsDGV.SelectedRows[0].Cells[3].Value.ToString();
            DocSpecCb.SelectedItem = DoctorsDGV.SelectedRows[0].Cells[4].Value.ToString();
            DocExpTb.Text = DoctorsDGV.SelectedRows[0].Cells[5].Value.ToString();
            DocPhoneTb.Text = DoctorsDGV.SelectedRows[0].Cells[6].Value.ToString();
            DocAddTb.Text = DoctorsDGV.SelectedRows[0].Cells[7].Value.ToString();
            DocPassWordTb.Text = DoctorsDGV.SelectedRows[0].Cells[8].Value.ToString();
            //Сохраняем ключ для обновления(id секретаря)
            if (DocNameTb.Text == "")
            {
                Key = 0;
            }
            else
            {
                Key = Convert.ToInt32(DoctorsDGV.SelectedRows[0].Cells[0].Value.ToString());
            }
        }

        private void EditBtn_Click(object sender, EventArgs e)
        {
            //проверка на пустые значения
            if (DocNameTb.Text == "" || DocPassWordTb.Text == "" || DocPhoneTb.Text == "" || DocAddTb.Text == "" || DocGenCb.SelectedIndex == -1 || DocSpecCb.SelectedIndex == -1)
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
                    //создаем команду для обновления
                    SqlCommand cmd = new SqlCommand("Update DoctorTbl set DocName=@DN,DOCDOB=@DD,DOCGEN=@DG,DOCSPEC=@DS,DOCEXP=@DE,DOCPHONE=@DP,DOCADD=@DA,DOCPASS=@DPA where DOCID=@DKey", Con);
                    //добавляем значения в строку
                    cmd.Parameters.AddWithValue("@DN", DocNameTb.Text);
                    cmd.Parameters.AddWithValue("@DD", DocDOB.Value.Date);
                    cmd.Parameters.AddWithValue("@DG", DocGenCb.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@DS", DocSpecCb.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@DE", DocExpTb.Text);
                    cmd.Parameters.AddWithValue("@DP", DocPhoneTb.Text);
                    cmd.Parameters.AddWithValue("@DA", DocAddTb.Text);
                    cmd.Parameters.AddWithValue("@DPA", DocPassWordTb.Text);
                    cmd.Parameters.AddWithValue("@DKey", Key);
                    cmd.ExecuteNonQuery();
                    // выполняет sql-выражение и возвращает количество измененных записей.
                    MessageBox.Show("Доктор изменен");
                    //закрываем подключение
                    Con.Close();
                    DisplayDoc(Query);
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
                MessageBox.Show("Выберите доктора");
            }
            else
            {
                //отлавливаем ошибку
                try
                {
                    //открываем подключение
                    Con.Open();
                    //создаем команду для редактирования
                    SqlCommand cmd = new SqlCommand("Delete from DoctorTbl where DocId=@DKey", Con);
                    //добавляем значения в строку
                    cmd.Parameters.AddWithValue("@DKey", Key);
                    cmd.ExecuteNonQuery();
                    // выполняет sql-выражение и возвращает количество измененных записей.
                    MessageBox.Show("Доктор удален");
                    //закрываем подключение
                    Con.Close();
                    DisplayDoc(Query);
                    Clear();
                }
                catch (Exception err)
                {
                    //показываем ошибку
                    MessageBox.Show(err.Message);
                }
            }
        }

        private void label11_Click(object sender, EventArgs e)
        {
            Homes obj = new Homes();
            obj.Show();
            this.Hide();
        }

        private void label12_Click(object sender, EventArgs e)
        {
            LabTests obj = new LabTests();
            obj.Show();
        }

        private void label13_Click(object sender, EventArgs e)
        {
            Receptionists obj = new Receptionists();
            obj.Show();
            this.Hide();
        }

        private void label10_Click(object sender, EventArgs e)
        {
            Patients obj = new Patients();
            obj.Show();
            this.Hide();
        }

        private void label14_Click(object sender, EventArgs e)
        {
            Login obj = new Login();
            obj.Show();
            this.Hide();
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
            xlSheet.Name = "Доктора";
            xlSheet.Cells.HorizontalAlignment = 3;
            for (int j = 1; j < DoctorsDGV.Columns.Count + 1; j++)
            {
                xlApp.Cells[1, j] = DoctorsDGV.Columns[j - 1].HeaderText;
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
                        for (int i = 0; i < DoctorsDGV.Rows.Count - 1; i++)
                        {
                            for (int j = 0; j < DoctorsDGV.Columns.Count; j++)
                            {
                                xlApp.Cells[i + 2, j + 1] = DoctorsDGV.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                        xlApp.Visible = true;
                        break;
                    };

                case DialogResult.No:
                    {
                        var xlApp = GetExcel();
                        for (int i = 0; i < DoctorsDGV.SelectedRows.Count - 1; i++)
                        {
                            for (int j = 0; j < DoctorsDGV.Columns.Count; j++)
                            {
                                xlApp.Cells[i + 2, j + 1] = DoctorsDGV.SelectedRows[i].Cells[j].Value.ToString();
                            }
                        }
                        xlApp.Visible = true;
                        break;
                    };

                case DialogResult.Cancel:
                    return;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < DoctorsDGV.RowCount; i++)
            {
                for (int j = 0; j < DoctorsDGV.ColumnCount; j++)
                {
                    DoctorsDGV[j, i].Style.BackColor = Color.White;
                    DoctorsDGV[j, i].Style.ForeColor = Color.Black;
                }
            }

            if (!string.IsNullOrWhiteSpace(textBox1.Text))
            {
                DoctorsDGV.ClearSelection();
                for (int i = 0; i < DoctorsDGV.RowCount - 1; i++)
                {
                    for (int j = 0; j < DoctorsDGV.ColumnCount; j++)
                    {
                        if (DoctorsDGV[j, i].Value.ToString().ToLower().Contains(textBox1.Text.ToLower()))
                        {
                            DoctorsDGV[j, i].Style.BackColor = Color.Black;
                            DoctorsDGV[j, i].Style.ForeColor = Color.White;
                        }
                    }
                }
            }
        }

        private void buttonSort_Click(object sender, EventArgs e)
        {
            var COL = new System.Windows.Forms.DataGridViewColumn();

            switch (comboBox1.SelectedItem.ToString())
            {
                case "Специальность":
                    COL = DoctorsDGV.Columns["DOCSPEC"];
                    break;
                case "Опыт":
                    COL = DoctorsDGV.Columns["DOCEXP"];
                    break;
                case "Имя":

                    COL = DoctorsDGV.Columns["DocName"];
                    break;
            }
            if (radioButtonUp.Checked)
            {
                DoctorsDGV.Sort(COL, System.ComponentModel.ListSortDirection.Ascending);
            }
            else
            {
                DoctorsDGV.Sort(COL, System.ComponentModel.ListSortDirection.Descending);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DisplayDoc(Query);
            buttonSort.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DisplayDoc(Query + " where DOCGEN=N'" + comboBox2.SelectedItem.ToString() + "'");
            buttonSort.Enabled = false;
        }

        private void SpLbl_Click(object sender, EventArgs e)
        {
            About obj = new About();
            obj.ShowDialog();
        }
    }
}
