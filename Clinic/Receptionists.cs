using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Clinic
{
    public partial class Receptionists : Form
    {
        public Receptionists()
        {
            InitializeComponent();
            DisplayRec();
            if (Login.Role == "Dcotor")
            {
                PatLbl.Enabled = false;
                DocLbl.Enabled = false;
                RecLbl.Enabled = false;
            }
        }
        //подключаемся к sql
        SqlConnection Con = new SqlConnection(@"workstation id=clinicdanyadb.mssql.somee.com;packet size=4096;user id=danyafrontend_SQLLogin_2;pwd=77s976d7og;data source=clinicdanyadb.mssql.somee.com;persist security info=False;initial catalog=clinicdanyadb");
      
        private void DisplayRec()
        {
            // получаем данные в DataSet через SqlDataAdapter
            Con.Open();
            string Query = "Select * from ReceptionistTbl";
            SqlDataAdapter sda = new SqlDataAdapter(Query,Con); 
            SqlCommandBuilder builder = new SqlCommandBuilder(sda);
            //создаем объект DataSet
            var ds = new DataSet();
            //заполняем DataSet
            sda.Fill(ds);
            //отображаем данные
            ReceptionistDGV.DataSource = ds.Tables[0];
            Con.Close();
        }

        private void AddBtn_Click(object sender, EventArgs e)
        {
            //проверка на пустые значения
            if (RNameTb.Text == "" || RPasswordTb.Text == "" || RPhoneTb.Text == "" || RAddressTb.Text == "")
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
                    //string query = $"INSERT INTO ReceptionistTbl(RecepName,RecepPhone,RecepAdd,RecepPass)VALUES('{RNameTb.Text}','{RPhoneTb.Text}','{RAddressTb.Text}','{RPasswordTb.Text}')";
                     SqlCommand cmd = new SqlCommand("insert into ReceptionistTbl(RecepName,RecepPhone,RecepAdd,RecepPass)values(@RN,@RP,@RA,@RPA)", Con);
                    //SqlCommand cmd = new SqlCommand(query, Con);
                    //добавляем значения в строку
                    cmd.Parameters.AddWithValue("@RN", RNameTb.Text);
                    cmd.Parameters.AddWithValue("@RP", RPhoneTb.Text);
                    cmd.Parameters.AddWithValue("@RA", RAddressTb.Text);
                    cmd.Parameters.AddWithValue("@RPA", RPasswordTb.Text);
                    cmd.ExecuteNonQuery();
                    // выполняет sql-выражение и возвращает количество измененных записей.
                    MessageBox.Show("Секретарь добавлен");
                    //закрываем подключение
                    Con.Close();
                    DisplayRec();
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
                MessageBox.Show("Выберите секретаря");
            }
            else
            {
                //отлавливаем ошибку
                try
                {
                    //открываем подключение
                    Con.Open();
                    //создаем команду для редактирования
                    SqlCommand cmd = new SqlCommand("Delete from ReceptionistTbl where RecepId=@RKey", Con);
                    //добавляем значения в строку
                    cmd.Parameters.AddWithValue("@RKey", Key);
                    cmd.ExecuteNonQuery();
                    // выполняет sql-выражение и возвращает количество измененных записей.
                    MessageBox.Show("Секретарь удален");
                    //закрываем подключение
                    Con.Close();
                    DisplayRec();
                    Clear();
                }
                catch (Exception err)
                {
                    //показываем ошибку
                    MessageBox.Show(err.Message);
                }
            }
        }
        int Key = 0;
        private void ReceptionistDGV_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //заполняем данные
            RNameTb.Text = ReceptionistDGV.SelectedRows[0].Cells[1].Value.ToString();
            RPhoneTb.Text = ReceptionistDGV.SelectedRows[0].Cells[2].Value.ToString();
            RAddressTb.Text = ReceptionistDGV.SelectedRows[0].Cells[3].Value.ToString();
            RPasswordTb.Text = ReceptionistDGV.SelectedRows[0].Cells[4].Value.ToString();
            //Сохраняем ключ для обновления(id секретаря)
            if(RNameTb.Text == "")
            {
                Key = 0;
            }
            else
            {
                Key = Convert.ToInt32( ReceptionistDGV.SelectedRows[0].Cells[0].Value.ToString());
            }
        }

        private void EditBtn_Click(object sender, EventArgs e)
        {
            //проверка на пустые значения
            if (RNameTb.Text == "" || RPasswordTb.Text == "" || RPhoneTb.Text == "" || RAddressTb.Text == "")
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
                    //создаем команду для редактирования
                    SqlCommand cmd = new SqlCommand("update ReceptionistTbl set RecepName=@RN,RecepPhone=@RP,RecepAdd=@RA,RecepPass=@RPA where RecepId=@Rkey", Con);
                    //добавляем значения в строку
                    cmd.Parameters.AddWithValue("@RN", RNameTb.Text);
                    cmd.Parameters.AddWithValue("@RP", RPhoneTb.Text);
                    cmd.Parameters.AddWithValue("@RA", RAddressTb.Text);
                    cmd.Parameters.AddWithValue("@RPA", RPasswordTb.Text);
                    cmd.Parameters.AddWithValue("@RKey",Key);
                    cmd.ExecuteNonQuery();
                    // выполняет sql-выражение и возвращает количество измененных записей.
                    MessageBox.Show("Секретарь обновлен");
                    //закрываем подключение
                    Con.Close();
                    DisplayRec();
                    Clear();
                }
                catch (Exception err)
                {
                    //показываем ошибку
                    MessageBox.Show(err.Message);
                }
            }
        }

        private void Clear()
        {
            //очищаем все значения
            RNameTb.Text = "";
            RPhoneTb.Text = "";
            RPasswordTb.Text = "";
            RAddressTb.Text = "";
            Key = 0;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void label14_Click(object sender, EventArgs e)
        {
            Login obj = new Login();
            obj.Show();
            this.Hide();
        }

        private void RecLbl_Click(object sender, EventArgs e)
        {
            Homes obj = new Homes();
            obj.Show();
            this.Hide();
        }

        private void LabLbl_Click(object sender, EventArgs e)
        {
            LabTests obj = new LabTests();
            obj.Show();
 
        }

        private void DocLbl_Click(object sender, EventArgs e)
        {
            Doctors obj = new Doctors();
            obj.Show();
            this.Hide();
        }

        private void PatLbl_Click(object sender, EventArgs e)
        {
            Patients obj = new Patients();
            obj.Show();
            this.Hide();
        }

        private void SpLbl_Click(object sender, EventArgs e)
        {
            About obj = new About();
            obj.ShowDialog();
        }
    }
}
