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

namespace Accounts
{
    public partial class StockSheet : Form
    {
        public StockSheet()
        {
            InitializeComponent();
        }

        private void StockSheet_Load(object sender, EventArgs e)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load("stocks.dbs");
            var root = doc.DocumentElement;

            for (int i = 0; i < root.ChildNodes.Count; i++)
            {
                for (int j = 0; j < root.ChildNodes[i].ChildNodes.Count; j++)
                {
                    var item = (new ListViewItem(new[] { root.ChildNodes[i].Attributes[0].InnerText,
                                                         root.ChildNodes[i].ChildNodes[j].Attributes[0].InnerText,
                                                        root.ChildNodes[i].ChildNodes[j].InnerText}));
                    listView1.Items.Add(item);

                }
            }
            ColumnSorter(1);
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

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            ColumnSorter(e.Column);
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

        private void button3_Click(object sender, EventArgs e)
        {
            ExportToExcel("ok.csv", listView1);
        }

        private void listView1_ColumnClick_1(object sender, ColumnClickEventArgs e)
        {
            ColumnSorter(e.Column);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listView1.BeginUpdate();
            listView1.Items.Clear();
            listView1.ListViewItemSorter = null;
            XmlDocument doc = new XmlDocument();
            doc.Load("stocks.dbs");
            var root = doc.DocumentElement;

            for (int i = 0; i < root.ChildNodes.Count; i++)
            {
                
                for (int j = 0; j < root.ChildNodes[i].ChildNodes.Count; j++)
                {
                    if (root.ChildNodes[i].Attributes[0].InnerText.ToUpper().Contains(companyTxtbox.Text.ToUpper()) && (root.ChildNodes[i].ChildNodes[j].Attributes[0].InnerText.ToUpper().Contains(itemTxtbox.Text.ToUpper())))
                    {
                        var item = (new ListViewItem(new[] { root.ChildNodes[i].Attributes[0].InnerText,
                                                         root.ChildNodes[i].ChildNodes[j].Attributes[0].InnerText,
                                                        root.ChildNodes[i].ChildNodes[j].InnerText}));
                        listView1.Items.Add(item);
                    }                 
                }
            }
          
            ColumnSorter(1);
            listView1.EndUpdate();
        }
    }
}
