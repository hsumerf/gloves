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
    public partial class stocks : Form
    {
        public stocks()
        {
            InitializeComponent();
        }

        public void RefreshList()
        {

            listView1.Items.Clear();
            XmlDocument doc = new XmlDocument();
            doc.Load("stocks.dbs");
            var node = doc.SelectSingleNode("//company[@name='" + comboBox1.Text + "']");
            var list = node.ChildNodes;

            for (int i = 0; i < list.Count; i++)
            {
                listView1.Items.Add(new ListViewItem(new[] { list.Item(i).Attributes[0].InnerText, list.Item(i).InnerText }));

            }
        }

        private void stocks_Load(object sender, EventArgs e)
        {
            relode();
        }

        public void relode()
        {
            comboBox1.Items.Clear();
            XmlDocument doc = new XmlDocument();
            doc.Load("stocks.dbs");
            XmlElement rootElement = doc.DocumentElement;
            var list = rootElement.ChildNodes;

            for (int i = 0; i < list.Count; i++)
            {

                comboBox1.Items.Add(list.Item(i).Attributes[0].InnerText);
            }
        
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //XmlDocument doc = new XmlDocument();
            //doc.Load("stocks.dbs");
            //XmlElement rootElement = doc.DocumentElement;
            //XmlElement newcomapny = doc.CreateElement(company.Text);
            //rootElement.AppendChild(newcomapny);
            //doc.Save("stocks.dbs");
            //relode();
            //comboBox1.SelectedIndex = comboBox1.Items.Count - 1;
            //RefreshList();
        }

        private void button4_Click(object sender, EventArgs e)
        {

            NewStock newsform = new NewStock(this);
            newsform.ShowDialog();
        }

        private void button2_Click_1(object sender, EventArgs e)
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
            for (int i = 0; i < listView1.SelectedItems.Count; i++)
            {
                var rootElement = doc.SelectSingleNode("//company[@name='" + comboBox1.Text + "']" + "//item[@name='" + listView1.SelectedItems[i].SubItems[0].Text + "']");
                rootElement.ParentNode.RemoveChild(rootElement);
              

            }
            doc.Save("stocks.dbs");
            RefreshList();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
            {
                MessageBox.Show("No item selected, Please select item to update.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Update updateform = new Update(this);
            updateform.itemLabel.Text = listView1.SelectedItems[0].SubItems[0].Text;
            updateform.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Company_Manager cmp_form = new Company_Manager(this);
            cmp_form.ShowDialog();
        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //ComboBox cb = (ComboBox)sender;
            ////       cb.DroppedDown = true;
            //listView1.Items.Clear();
            //string strFindStr = "";
            //if (e.KeyChar == (char)8)
            //{
            //    if (cb.SelectionStart <= 1)
            //    {
            //        cb.Text = "";
            //        return;
            //    }

            //    if (cb.SelectionLength == 0)
            //        strFindStr = cb.Text.Substring(0, cb.Text.Length - 1);
            //    else
            //        strFindStr = cb.Text.Substring(0, cb.SelectionStart - 1);
            //}
            //else
            //{
            //    if (cb.SelectionLength == 0)
            //        strFindStr = cb.Text + e.KeyChar;
            //    else
            //        strFindStr = cb.Text.Substring(0, cb.SelectionStart) + e.KeyChar;
            //}
            //int intIdx = -1;
            //// Search the string in the ComboBox list.
            //intIdx = cb.FindString(strFindStr);
            //if (intIdx != -1)
            //{
            //    cb.SelectedText = "";
            //    cb.SelectedIndex = intIdx;
            //    cb.SelectionStart = strFindStr.Length;
            //    cb.SelectionLength = cb.Text.Length;
            //    e.Handled = true;
            //}
            //else
            //    e.Handled = true;
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
           
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshList();
        }
    }
}
