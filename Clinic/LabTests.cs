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

namespace Clinic
{
    public partial class LabTests : Form
    {
        public LabTests()
        {
            InitializeComponent();
            DisplayTest();
        }
        int Key = 0;
        SqlConnection Con = new SqlConnection(@"workstation id=clinicdanyadb.mssql.somee.com;packet size=4096;user id=danyafrontend_SQLLogin_2;pwd=77s976d7og;data source=clinicdanyadb.mssql.somee.com;persist security info=False;initial catalog=clinicdanyadb");
        private void DisplayTest()
        {
            // получаем данные в DataSet через SqlDataAdapter
            Con.Open();
            string Query = "Select * from TestTbl";
            SqlDataAdapter sda = new SqlDataAdapter(Query, Con);
            SqlCommandBuilder builder = new SqlCommandBuilder(sda);
            //создаем объект DataSet
            var ds = new DataSet();
            //заполняем DataSet
            sda.Fill(ds);
            //отображаем данные
            LabTestDGV.DataSource = ds.Tables[0];
            Con.Close();
        }
        private void Clear()
        {
            //очищаем все значения
            LabTestTb.Text = "";
            LabCostTb.Text = "";
            Key = 0;
        }
        private void PatientsDGV_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //заполняем данные
            LabTestTb.Text = LabTestDGV.SelectedRows[0].Cells[1].Value.ToString();
            LabCostTb.Text = LabTestDGV.SelectedRows[0].Cells[2].Value.ToString();

            //Сохраняем ключ для обновления(id секретаря)
            if (LabTestTb.Text == "")
            {
                Key = 0;
            }
            else
            {
                Key = Convert.ToInt32(LabTestDGV.SelectedRows[0].Cells[0].Value.ToString());
            }
        }

        private void AddBtn_Click(object sender, EventArgs e)
        {
            //проверка на пустые значения
            if (LabTestTb.Text == "" || LabCostTb.Text == "")
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
                    SqlCommand cmd = new SqlCommand("insert into TestTbl(TestName,TestCost)values(@TN,@TC)", Con);
                    //добавляем значения в строку
                    cmd.Parameters.AddWithValue("@TN", LabTestTb.Text);
                    cmd.Parameters.AddWithValue("@TC", LabCostTb.Text);

                    cmd.ExecuteNonQuery();
                    // выполняет sql-выражение и возвращает количество измененных записей.
                    MessageBox.Show("Тест добавлен");
                    //закрываем подключение
                    Con.Close();
                    DisplayTest();
                    Clear();
                }
                catch (Exception err)
                {
                    //показываем ошибку
                    MessageBox.Show(err.Message);
                }
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void editBtn_Click(object sender, EventArgs e)
        {
            //проверка на пустые значения
            if (LabTestTb.Text == "" || LabCostTb.Text == "")
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
                    SqlCommand cmd = new SqlCommand("update TestTbl set TestName=@TN,TestCost=@TC where TestNum=@TKey", Con);
                    //добавляем значения в строку
                    cmd.Parameters.AddWithValue("@TN", LabTestTb.Text);
                    cmd.Parameters.AddWithValue("@TC", LabCostTb.Text);
                    cmd.Parameters.AddWithValue("@TKey", Key);

                    cmd.ExecuteNonQuery();
                    // выполняет sql-выражение и возвращает количество измененных записей.
                    MessageBox.Show("Тест обновлен");
                    //закрываем подключение
                    Con.Close();
                    DisplayTest();
                    Clear();
                }
                catch (Exception err)
                {
                    //показываем ошибку
                    MessageBox.Show(err.Message);
                }
            }
        }

        private void delBtn_Click(object sender, EventArgs e)
        {
            //если не выбран никто
            if (Key == 0)
            {
                MessageBox.Show("Выберите тест");
            }
            else
            {
                //отлавливаем ошибку
                try
                {
                    //открываем подключение
                    Con.Open();
                    //создаем команду для редактирования
                    SqlCommand cmd = new SqlCommand("Delete from TestTbl where TestNum=@TKey", Con);
                    //добавляем значения в строку
                    cmd.Parameters.AddWithValue("@TKey", Key);
                    cmd.ExecuteNonQuery();
                    // выполняет sql-выражение и возвращает количество измененных записей.
                    MessageBox.Show("Тест удален");
                    //закрываем подключение
                    Con.Close();
                    DisplayTest();
                    Clear();
                }
                catch (Exception err)
                {
                    //показываем ошибку
                    MessageBox.Show(err.Message);
                }
            }
        }
    }
}
