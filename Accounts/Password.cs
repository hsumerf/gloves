using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Accounts
{
    public partial class Password : Form
    {
        public Password()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "tabish12345karachi")

            {
                MessageBox.Show("Done, Thank you.","Correct",MessageBoxButtons.OK,MessageBoxIcon.Information);
                Properties.Settings.Default.first = false;
                Properties.Settings.Default.Save();
                this.Close();
            }

            else
                MessageBox.Show("Wrong passowrd, you can not use this software.","Incorrect",MessageBoxButtons.OK,MessageBoxIcon.Error);
        }

        private void Password_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (textBox1.Text != "tabish12345karachi")
            {
                Application.Exit();
            }
        }
    }
}
