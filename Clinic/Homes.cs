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
    public partial class Homes : Form
    {
        public Homes()
        {
            InitializeComponent();
            if(Login.Role == "Receptionist")
            {
                RecLbl.Enabled = false;
                DoctorLbl.Enabled = false;
                LabLbl.Enabled = false;
            }else if(Login.Role == "Doctor")
            {
                Prescriptions obj = new Prescriptions();
                obj.Show();
                this.Hide();
            }
            CountPatients();
            CountDoctors();
            CountTest();
            CountHIV();
        }
        SqlConnection Con = new SqlConnection(@"workstation id=clinicdanyadb.mssql.somee.com;packet size=4096;user id=danyafrontend_SQLLogin_2;pwd=77s976d7og;data source=clinicdanyadb.mssql.somee.com;persist security info=False;initial catalog=clinicdanyadb");
       private void CountPatients()
        {
            Con.Open();
            SqlDataAdapter sda = new SqlDataAdapter("Select count(*) from PatientTbl", Con);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            PatNumLbl.Text = dt.Rows[0][0].ToString();
            Con.Close();
   }
        private void CountDoctors()
        {
            Con.Open();
            SqlDataAdapter sda = new SqlDataAdapter("Select count(*) from DoctorTbl", Con);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            DocNumLbl.Text = dt.Rows[0][0].ToString();
            Con.Close();
        }
        private void CountTest()
        {
            Con.Open();
            SqlDataAdapter sda = new SqlDataAdapter("Select count(*) from TestTbl", Con);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            LabTestLbl.Text = dt.Rows[0][0].ToString();
            Con.Close();
        }
        private void CountHIV()
        {
            string status = "Положительный";
            Con.Open();
            SqlDataAdapter sda = new SqlDataAdapter("Select count(*) from PatientTbl where PatHIV=N'"+ status + "'", Con);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            HIVLbl.Text = dt.Rows[0][0].ToString();
            Con.Close();
        }
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void PatLbl_Click(object sender, EventArgs e)
        {
            Patients obj = new Patients();
            obj.Show();
            this.Hide();
        }

        private void DoctorLbl_Click(object sender, EventArgs e)
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

        private void label14_Click(object sender, EventArgs e)
        {
            Login obj = new Login();
            obj.Show();
            this.Hide();
        }
    }
}
