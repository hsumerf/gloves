using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;

namespace Accounts
{
    public partial class Backup : Form
    {
        public Backup()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            textBox1.Text = folderBrowserDialog1.SelectedPath;
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            try
            {
                sb.Append(System.IO.File.ReadAllText("main.dbs"));
                sb.Append(("*File-Ends-Here*"));

                sb.Append(System.IO.File.ReadAllText("stocks.dbs"));
                sb.Append(("*File-Ends-Here*"));

                sb.Append(System.IO.File.ReadAllText("clients.dbs"));
                sb.Append(("*File-Ends-Here*"));

                sb.Append(System.IO.File.ReadAllText("category.dbs"));
            }
            catch (Exception)
            {
                MessageBox.Show("Unable to backup, data files are courpted!", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            try
            {
                System.IO.File.WriteAllText(textBox1.Text + "\\Backup.bck", sb.ToString());
                MessageBox.Show("Data backed up succesfully!", "Succeeded!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();

            }
            catch (Exception)
            {
                MessageBox.Show("Can not find backup location", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void button4_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Backup file (*.bck)|*.bck|All files (*.*)|*.*";
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
           textBox2.Text  = openFileDialog1.FileName;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string source = System.IO.File.ReadAllText(textBox2.Text);
                var files = source.Split(new string[] { "*File-Ends-Here*" }, StringSplitOptions.None);
                System.IO.File.WriteAllText("main.dbs", files[0]);
                System.IO.File.WriteAllText("stocks.dbs", files[1]);
                System.IO.File.WriteAllText("clients.dbs", files[2]);
                System.IO.File.WriteAllText("category.dbs", files[3]);
                MessageBox.Show("Data resoterd succesfully!", "Succeeded!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();

            }
            catch (Exception)
            {
                MessageBox.Show("Can not restore backup file is incorrect", "Failed!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            StringBuilder sb = new StringBuilder();

            try
            {
                sb.Append(System.IO.File.ReadAllText("main.dbs"));
                sb.Append(("*File-Ends-Here*"));

                sb.Append(System.IO.File.ReadAllText("stocks.dbs"));
                sb.Append(("*File-Ends-Here*"));

                sb.Append(System.IO.File.ReadAllText("clients.dbs"));
                sb.Append(("*File-Ends-Here*"));

                sb.Append(System.IO.File.ReadAllText("category.dbs"));
            }
            catch (Exception)
            {
                MessageBox.Show("Unable to backup, data files are courpted!", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            try
            {
                System.IO.File.WriteAllText(textBox1.Text + "\\Backup.bck", sb.ToString());
                MessageBox.Show("Data backed up succesfully!", "Succeeded!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();

            }
            catch (Exception)
            {
                MessageBox.Show("Can not find backup location", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            mail.From = new MailAddress("samitraders2018@gmail.com");
            mail.To.Add(textBox3.Text);
            mail.Subject = "Sami Traders - Backup";
            mail.Body = DateTime.Now.ToString();

            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment("Backup.bck");
            mail.Attachments.Add(attachment);

            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential("samitraders2018@gmail.com", "sami2018");
            SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);
            MessageBox.Show("Back has been sent to your email adress!");
        }

        private void Backup_Load(object sender, EventArgs e)
        {

        }
    }
}
