using System;
using System.Windows.Forms;

namespace Assignment1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region Button: Open SQL Script File
        /// <summary>
        /// This button's function is to open a file dialog window for the user to choose a SQL script file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            label2.Visible = false;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Open SQL Query File";
            openFileDialog1.InitialDirectory = @"c:\";
            openFileDialog1.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }
        #endregion

        #region Button: Open Access Database File
        /// <summary>
        /// This button's function is to open a file dialog window for the user to choose an Access database file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            label2.Visible = false;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Open Microsoft Access Database File";
            openFileDialog1.InitialDirectory = @"c:\";
            openFileDialog1.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialog1.FileName;
            }
        }
        #endregion

        #region Button: Run SQL Interpreter
        /// <summary>
        /// This button's function is to run the user's SQL script against the Access database they chose.
        /// It will also print the completed SQL scripts in the listbox.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="a"></param>
        private void button3_Click(object sender, EventArgs a)
        {
            listBox1.Items.Clear(); // Clear the listbox every time you Run a new script.
            label2.Visible = false;
            pictureBox1.Visible = true;
            label1.Visible = true;

            bool invalid = false; // Used for the following error handling.

            if (textBox1.Text == "" || textBox2.Text == " ")
            {
                MessageBox.Show("Please choose a valid file location for your query.", "Invalid Query File Location", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                invalid = true;
            }

            if (textBox2.Text == "" || textBox2.Text == " ")
            {
                MessageBox.Show("Please choose a valid Microsoft Access Database File location.", "Invalid File Location", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                invalid = true;
            }

            if (invalid == true)
            {
                pictureBox1.Visible = false;
                label1.Visible = false;
                return;
            }

            // OLEDB Connection String
            string ConnectString = "provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + textBox2.Text;

            // OLEDB Connection Object
            System.Data.OleDb.OleDbConnection Connection = new System.Data.OleDb.OleDbConnection(ConnectString);

            try
            {
                Connection.Open();
            }

            catch (System.Exception e)
            {
                MessageBox.Show("Problems opening... " + e.Message, "Error Running Your Script", MessageBoxButtons.OK, MessageBoxIcon.Error);
                pictureBox1.Visible = false;
                label1.Visible = false;
                return;
            }

            string Buffer = System.IO.File.ReadAllText(textBox1.Text); // Create string object 'Buffer'. Use ReadAllText() and provide the file path that we stored in the textbox as a string.
            string[] sqlCommands = Buffer.Split(new char[] { ';' }); // delimit and separate the script file into individual query statements.

            // Since we just got rid of all the semicolons(;), we can add them back this way.
            int k = 0;
            while (k < sqlCommands.Length)
            {
                sqlCommands[k] += ";";
                k++;
            }

            for (int j = 0; j < sqlCommands.Length - 1; j++)
            {
                string[] line = sqlCommands[j].Split(new char[] { '\n' }); // For every query statement in the sql script, delimit and separate the statement into lines.

                try
                {
                    int i = 0;
                    while (i < line.Length)
                    {
                        listBox1.Items.Add(line[i]); // Populate the listbox with String items.
                        i++;
                    }
                    System.Data.OleDb.OleDbCommand SQLCommand = new System.Data.OleDb.OleDbCommand(sqlCommands[j], Connection); // Create OleDB command object using our query statement and connection string.
                    SQLCommand.ExecuteNonQuery(); // Execute OleDB command object.
                }

                catch (System.Exception e)
                {
                    listBox1.Items.Add(e.Message);
                    MessageBox.Show("Problems executing..." + e.Message, "Error Running Your Script", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Connection.Close();
                    pictureBox1.Visible = false;
                    label1.Visible = false;
                    return;
                }
            }

            pictureBox1.Visible = false;
            label1.Visible = false;
            label2.Visible = true;

            try
            {
                Connection.Close();
            }

            catch (System.Exception e)
            {
                MessageBox.Show("Problems closing..." + e.Message);
                return;
            }
        }
        #endregion

        #region Button: Close Application
        private void button4_Click(object sender, EventArgs e)
        {
            Application.Exit(); // End the Application Process.
        }
        #endregion
    }
}