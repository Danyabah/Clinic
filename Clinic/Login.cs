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
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void ResetBtn_Click(object sender, EventArgs e)
        {
            RoleCb.SelectedIndex = 0;
            uNameTb.Text = "";
            PassTb.Text = "";
        }
        SqlConnection Con = new SqlConnection(@"workstation id=clinicdanyadb.mssql.somee.com;packet size=4096;user id=danyafrontend_SQLLogin_2;pwd=77s976d7og;data source=clinicdanyadb.mssql.somee.com;persist security info=False;initial catalog=clinicdanyadb");
        public static string Role;
        private void LoginBtn_Click(object sender, EventArgs e)
        {
            if(RoleCb.SelectedIndex == -1)
            {
                MessageBox.Show("Выберите вашу должность");
            }
            else if (RoleCb.SelectedIndex == 0)
            {
               if(uNameTb.Text == "" || PassTb.Text == "")
                {
                    MessageBox.Show("Введите имя и пароль");
                }else if(uNameTb.Text == "Админ" && PassTb.Text == "12345678")
                {
                    Role = "Admin";
                    Patients Obj = new Patients();
                    Obj.Show();
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("Не верное имя или пароль");
                }
            }else if(RoleCb.SelectedIndex == 1)
            {
                if (uNameTb.Text == "" || PassTb.Text == "")
                {
                    MessageBox.Show("Введите имя и пароль");
                }
                else
                {
                    Con.Open();
                    SqlDataAdapter sda = new SqlDataAdapter("Select Count(*) from DoctorTbl where DocName=N'"+uNameTb.Text+"' and DocPass='" +PassTb.Text +"'",Con);
                    DataTable dt = new DataTable();
                    sda.Fill(dt);
                    if (dt.Rows[0][0].ToString() == "1")
                    {
                        Role = "Doctor";
                        Prescriptions Obj = new Prescriptions();
                        Obj.Show();
                        this.Hide();
                    }
                    else
                    {
                        MessageBox.Show("Доктор не найден");
                    }
                    Con.Close();
                }
            }
            else 
            {
                if (uNameTb.Text == "" || PassTb.Text == "")
                {
                    MessageBox.Show("Введите имя и пароль");
                }
                else
                {
                    Con.Open();
                    SqlDataAdapter sda = new SqlDataAdapter("Select Count(*) from ReceptionistTbl where RecepName=N'" + uNameTb.Text + "' and RecepPass='" + PassTb.Text + "'", Con);
                    DataTable dt = new DataTable();
                    sda.Fill(dt);
                    if (dt.Rows[0][0].ToString() == "1")
                    {
                        Role = "Receptionist";
                        Homes Obj = new Homes();
                        Obj.Show();
                        this.Hide();
                    }
                    else
                    {
                        MessageBox.Show("Секретарь не найден");
                    }
                    Con.Close();
                }
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
