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
    public partial class Clients : Form
    {
        public Clients()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        public void reload()
        {
            listView1.Items.Clear();
            XmlDocument doc = new XmlDocument();
            doc.Load("clients.dbs");
            XmlElement rootElement = doc.DocumentElement;
            var list = rootElement.ChildNodes;
            for (int i = 0; i < list.Count; i++)
                listView1.Items.Add(new ListViewItem(new[] { list.Item(i).Attributes[0].Value,
                                                             list.Item(i).ChildNodes[2].InnerText,
                                                             list.Item(i).ChildNodes[3].InnerText,
                                                             }));
        }
        public void ColumnSorter(int index)
        {
            sorter sorter = listView1.ListViewItemSorter as sorter;
            if (sorter == null)
            {
                sorter = new sorter(index);
                listView1.ListViewItemSorter = sorter;
            }
            else
            {
                sorter.Column = index;
            }
            listView1.Sort();
        }

        private void Clients_Load(object sender, EventArgs e)
        {
            reload();
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
            {
                MessageBox.Show("No item selected, Please select item to delete.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            doc.Load("clients.dbs");
            for (int i = 0; i < listView1.SelectedItems.Count; i++)
            {
                var nodetodelete = doc.SelectSingleNode("//Client[@name='" + listView1.SelectedItems[i].SubItems[0].Text + "']");
                nodetodelete.ParentNode.RemoveChild(nodetodelete);
            }
          
            doc.Save("clients.dbs");
            reload();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            category_manager cm = new category_manager(this);
            cm.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
            {
                MessageBox.Show("No item selected, Please select item to update.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            clint_update clupdt = new clint_update(this);
            clupdt.textBox1.Text = listView1.SelectedItems[0].SubItems[0].Text;
            clupdt.comboBox1.SelectedItem = listView1.SelectedItems[0].SubItems[1].Text;
            clupdt.textBox2.Text = listView1.SelectedItems[0].SubItems[2].Text;
            clupdt.ShowDialog();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            New_Client nc = new New_Client(this);
            nc.ShowDialog();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            category_manager cm = new category_manager(this);
            cm.ShowDialog();
        }

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            ColumnSorter(e.Column);
        }
    }
}
