using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Bank_Activity
{
    public partial class BankActitvityForm : Form
    {
        public BankActitvityForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string[] bankNames = { "Citi Bank", "Citizens Bank", "Community & Southern", "Pacific Western Bank", "PNC Bank", "Private Bank", "TD Bank", "Wells Fargo Bank" };


            BankActivityForm.DataSource = bankNames;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog1 = new OpenFileDialog();
            fileDialog1.Filter = "CSV files (*.csv)|*.csv";
            fileDialog1.FilterIndex = 2;
            fileDialog1.RestoreDirectory = true;

            fileDialog1.ShowDialog();

            string filePath = fileDialog1.FileName;
            if (filePath != "")
                textBox1.Text = filePath;

        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            MainForm frm = new MainForm();
            frm.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "EXCEL files (*.xlsx)|*.xlsx";

            saveFileDialog1.ShowDialog();

            string savePath = saveFileDialog1.FileName;
            if (savePath != "")
                textBox2.Text = savePath;
       
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filePath = textBox1.Text;
            string savePath = textBox2.Text;
            string bank = BankActivityForm.GetItemText(BankActivityForm.SelectedItem);

			if ( filePath != null && savePath != null)
			{
				//Change the cursor to a wait cursor
				Cursor.Current = Cursors.WaitCursor;
				this.label4.Text = "Please wait while the program does it's job...";

				//Run the program with a try clause for errors
				try
				{
					MainProgram.Start(filePath, savePath, bank);
				}
				catch
				{
					MessageBox.Show("There was an error running the program please try again.");
				}               
							
				MessageBox.Show("The program has completed it's task");
			}
			else
			{
                MessageBox.Show("You have not selected a file to open, or to save");
			}
			
            //Return to main form
            this.Hide();
            MainForm frm = new MainForm();
            frm.Show();

            //Switch back to default cursor
            Cursor.Current = Cursors.Default;
            this.label4.Text = "Done...";
        }
    }
}

