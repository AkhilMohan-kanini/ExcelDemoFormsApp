using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Configuration;

namespace ExcelDemoFormsApp
{
    public partial class Form1 : Form
    {
        //Global Variable - Connection String
        string oleConDbDetails = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
        OleDbConnection conn;

        public Form1()
        {
            InitializeComponent();
            //Creating Database Connection
            conn = new OleDbConnection(oleConDbDetails);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            try
            {
                //Opening  Database Connection
                conn.Open();
                TestDbLabel.Text = "Connected Successfully";

                //Command for fetching data from Excel
                OleDbDataAdapter cmd   = new OleDbDataAdapter("select * from [Sheet1$]", conn);


                // adding data from excel to  a dataset
                DataTable dtTable = new DataTable();
                cmd.Fill(dtTable);

                //Displaying data with GridView
                dataGridView1.DataSource = dtTable;

            }
            catch(Exception exc)
            {
                //Display in label if Connection fails
                TestDbLabel.Text = "Connection Failed : " + exc.Message;
            }
            finally
            {
                //Close the database Connection
                conn.Close();
            }

        }
    }
}
