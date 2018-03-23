using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace Accounts
{
    public partial class balSheet : Form
    {
        public balSheet()
        {
            InitializeComponent();
        }

        private void balSheet_Load(object sender, EventArgs e)
        {
            //XmlDocument doc = new XmlDocument();
            //doc.Load("clients.dbs");
            //var root = doc.DocumentElement;
            //var namelist = root.ChildNodes;

            //for (int i = 0; i < namelist.Count; i++)
            //{
            //    var item = (new ListViewItem(new[] { namelist.Item(i).Attributes[0].InnerText,
            //                                         namelist.Item(i).ChildNodes[0].InnerText,
            //                                         namelist.Item(i).ChildNodes[1].InnerText,
            //                                         (Convert.ToInt32(namelist.Item(i).ChildNodes[0].InnerText) - Convert.ToInt32(namelist.Item(i).ChildNodes[1].InnerText)).ToString()}));
            //    listView1.Items.Add(item);
            //}

            comboBox1.Items.Clear();
            var list1 = System.IO.File.ReadAllLines("category.dbs");
            for (int i = 0; i < list1.Length; i++)
            {
                comboBox1.Items.Add(list1[i]);
            }
            comboBox1.Items.Insert(0, "All");
            comboBox1.SelectedIndex = 0;

            ColumnSorter(0);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportToExcel("ok.csv", listView1);
        }

        private void ExportToExcel(string path, ListView listsource)
        {
            StringBuilder CVS = new StringBuilder();
            for (int i = 0; i < listsource.Columns.Count; i++)
            {
                CVS.Append(listsource.Columns[i].Text + ",");
            }
            CVS.Append(Environment.NewLine);
            for (int i = 0; i < listsource.Items.Count; i++)
            {
                for (int j = 0; j < listsource.Columns.Count; j++)
                {
                    CVS.Append(listsource.Items[i].SubItems[j].Text + ",");
                }
                CVS.Append(Environment.NewLine);
            }
            System.IO.File.WriteAllText(path, CVS.ToString());
            Process.Start(path);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double rec = 0;
            double crd = 0;
            XDocument X = XDocument.Load("clients.dbs");
            var FilterList = X.Element("clients").Elements("Client").Where
                (E => (((comboBox1.Text == "All" )? true : (E.Element("cate").Value == comboBox1.Text)) && (E.Element("des").Value.ToUpper().Contains(textBox1.Text.ToUpper()))));
           listView1.BeginUpdate();
           listView1.ListViewItemSorter = null;
            listView1.Items.Clear();
            foreach (var item in FilterList)
            {
                listView1.Items.Add(new ListViewItem(new[] { item.FirstAttribute.Value,
                                                             item.Element("total").Value,
                                                             item.Element ("recieve").Value,
                                                             (double.Parse(item.Element("total").Value) - double.Parse(item.Element("recieve").Value)).ToString(),
                                                      
                }));
                crd += double.Parse(item.Element("total").Value);
                rec += double.Parse(item.Element("recieve").Value);
            }
            ColumnSorter(0);
            listView1.EndUpdate();
            total_crd.Text = crd.ToString();
            total_Rec.Text = rec.ToString();
            label4.Text = (crd - rec).ToString();

        }

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            ColumnSorter(e.Column);
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
    }
}
