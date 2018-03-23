using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Accounts
{
    public partial class New_Client : Form
    {
        Clients clt;
        public New_Client(Clients clta)
        {
            InitializeComponent();
            clt = clta;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
                return;

            XmlDocument doc = new XmlDocument();
            doc.Load("clients.dbs");

            if (doc.SelectSingleNode("//Client[@name='" + textBox1.Text + "']") != null)
            {
                MessageBox.Show("Already");
                return;
            }

            XmlElement rootElement = doc.DocumentElement;
            XmlElement newClient = doc.CreateElement("Client");
            newClient.SetAttribute("name", textBox1.Text);
            XmlElement total = doc.CreateElement("total");
            total.InnerText = "0";
            XmlElement recieve = doc.CreateElement("recieve");
            recieve.InnerText = "0";
            XmlElement cate = doc.CreateElement("cate");
            cate.InnerText = comboBox1.Text ;
            XmlElement des = doc.CreateElement("des");
            des.InnerText = textBox2.Text;
            newClient.AppendChild(total);
            newClient.AppendChild(recieve);
            newClient.AppendChild(cate);
            newClient.AppendChild(des);
            rootElement.AppendChild(newClient);
            doc.Save("clients.dbs");
            textBox1.Text = "";
            clt.reload();
        }

        private void New_Client_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            var list1 = System.IO.File.ReadAllLines("category.dbs");
            for (int i = 0; i < list1.Length; i++)
            {
                comboBox1.Items.Add(list1[i]);
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Enter))
            {
                SendKeys.Send("{TAB}");
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
