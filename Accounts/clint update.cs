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
    public partial class clint_update : Form
    {
        Clients clnt;
        public clint_update(Clients clnta)
        {
            InitializeComponent();
            clnt = clnta;
        }

        private void clint_update_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            var list1 = System.IO.File.ReadAllLines("category.dbs");
            for (int i = 0; i < list1.Length; i++)
            {
                comboBox1.Items.Add(list1[i]);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            
            XmlDocument doc = new XmlDocument();
            doc.Load("clients.dbs");
            var nodetodelete = doc.SelectSingleNode("//Client[@name='" + textBox1.Text + "']");
            nodetodelete.ChildNodes[2].InnerText = comboBox1.Text;
            nodetodelete.ChildNodes[3].InnerText = textBox2.Text;
            doc.Save("clients.dbs");
            clnt.reload();
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
