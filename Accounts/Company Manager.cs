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
    public partial class Company_Manager : Form
    {
        stocks stocks;
        public Company_Manager(stocks stocksFrom)
        {
            InitializeComponent();
            stocks = stocksFrom;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
                return;

            XmlDocument doc = new XmlDocument();
            doc.Load("stocks.dbs");
            if (doc.SelectSingleNode("//company[@name='" + textBox1.Text + "']") != null)
            {
                MessageBox.Show("Already");
                return;
            }
            XmlElement rootElement = doc.DocumentElement;
            XmlElement newcomapny = doc.CreateElement("company");
            newcomapny.SetAttribute("name", textBox1.Text);
            rootElement.AppendChild(newcomapny);
            doc.Save("stocks.dbs");
            stocks.relode();
            stocks.comboBox1.SelectedItem = textBox1.Text;
            stocks.RefreshList();
            reload();
            textBox1.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
            {
                MessageBox.Show("No item selected, Please select item to delete.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DialogResult dialogResult = MessageBox.Show("Are you sure you want to delete?", "Confrimation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
            }
            else if (dialogResult == DialogResult.No)
            {
                return;
            }
            XmlDocument doc = new XmlDocument();
            doc.Load("stocks.dbs");
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                var nodetodelete = doc.SelectSingleNode("//company[@name='" + listView1.SelectedItems[i].SubItems[0].Text + "']");
                nodetodelete.ParentNode.RemoveChild(nodetodelete);
            }
            
            doc.Save("stocks.dbs");
            stocks.relode();
            reload();
        }

        private void Company_Manager_Load(object sender, EventArgs e)
        {
            reload();
        }


        public void reload() {
            listView1.Items.Clear();
            XmlDocument doc = new XmlDocument();
            doc.Load("stocks.dbs");
            XmlElement rootElement = doc.DocumentElement;
            var list = rootElement.ChildNodes;

            for (int i = 0; i < list.Count; i++)
            {

                listView1.Items.Add(list.Item(i).Attributes[0].InnerText);
            }
        }
    }
}
