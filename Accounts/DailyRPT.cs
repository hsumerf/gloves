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
    public partial class DailyRPT : Form
    {
        public DailyRPT()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            double rec = 0;
            double crd = 0;
            XDocument X = XDocument.Load("main.dbs");
         //   var dateparse = DateTime.ParseExact(E.Element("date").Value), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            var FilterList = X.Element("all").Elements("person").Where
                (E => (DateTime.ParseExact((E.Element("date").Value), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture) >= dateTimeFrom.Value.AddDays(-1) && (DateTime.ParseExact((E.Element("date").Value), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture) < dateTimeTo.Value)));
            listView1.BeginUpdate();
            listView1.ListViewItemSorter = null;
            listView1.Items.Clear();
            foreach (var item in FilterList)
            {
                listView1.Items.Add(new ListViewItem(new[] { item.FirstAttribute.Value,
                                                             item.Element("date").Value,
                                                             item.Element("name").Value,
                                                             item.Element ("company").Value,
                                                             item.Element("item").Value,
                                                             item.Element ("qty").Value,
                                                             item.Element("peritem").Value,
                                                             item.Element("total").Value,
                                                             item.Element("recieve").Value,
                                                             item.Element("details").Value}));
                crd += double.Parse(item.Element("total").Value);
                rec += double.Parse(item.Element("recieve").Value);
            }

            ColumnSorter(1);

            listView1.EndUpdate();
            total_crd.Text = crd.ToString();
            total_Rec.Text = rec.ToString();
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
            doc.Load("main.dbs");

            XmlDocument stockdoc = new XmlDocument();
            stockdoc.Load("stocks.dbs");

            XmlDocument clientdoc = new XmlDocument();
            clientdoc.Load("clients.dbs");

            for (int i = 0; i < listView1.SelectedItems.Count; i++)
            {
         
                try
                {
                    string qtyafter;
                    string qtybefore = stockdoc.SelectSingleNode("//company[@name='" + listView1.SelectedItems[i].SubItems[3].Text + "']" + "//item[@name='" + listView1.SelectedItems[i].SubItems[4].Text + "']").InnerText;
                    if (listView1.SelectedItems[i].SubItems[8].Text == "0")
                    {
                        qtyafter = (Convert.ToDouble(qtybefore) + Convert.ToDouble(listView1.SelectedItems[i].SubItems[5].Text)).ToString();

                    }
                    else
                    {
                        qtyafter = (Convert.ToDouble(qtybefore) - Convert.ToDouble(listView1.SelectedItems[i].SubItems[5].Text)).ToString();
                    }

                    stockdoc.SelectSingleNode("//company[@name='" + listView1.SelectedItems[i].SubItems[3].Text + "']" + "//item[@name='" + listView1.SelectedItems[i].SubItems[4].Text + "']").InnerText = qtyafter;
                }
                catch (Exception)
                {
                }

                try
                {
                    double Before = Convert.ToDouble(clientdoc.SelectSingleNode("//Client[@name='" + listView1.SelectedItems[i].SubItems[2].Text + "']" + "//total").InnerText);
                    double After = Before - Convert.ToDouble(listView1.SelectedItems[i].SubItems[7].Text);
                    clientdoc.SelectSingleNode("//Client[@name='" + listView1.SelectedItems[i].SubItems[2].Text + "']" + "//total").InnerText = After.ToString();
                }
                catch (Exception)
                {
                }

                try
                {
                    double Before = Convert.ToDouble(clientdoc.SelectSingleNode("//Client[@name='" + listView1.SelectedItems[i].SubItems[2].Text + "']" + "//recieve").InnerText);
                    double After = Before - Convert.ToDouble(listView1.SelectedItems[i].SubItems[8].Text);
                    clientdoc.SelectSingleNode("//Client[@name='" + listView1.SelectedItems[i].SubItems[2].Text + "']" + "//recieve").InnerText = After.ToString();
                }
                catch (Exception)
                {
                }

                var nodetodelete = doc.SelectSingleNode("//person[@id='" + listView1.SelectedItems[i].SubItems[0].Text + "']");
                nodetodelete.ParentNode.RemoveChild(nodetodelete);
            }

            doc.Save("main.dbs");
            clientdoc.Save("clients.dbs");
            stockdoc.Save("stocks.dbs");

           button3.PerformClick();
        }

        private void button1_Click(object sender, EventArgs e)
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

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            ColumnSorter(e.Column);
        }

        private void DailyRPT_Load(object sender, EventArgs e)
        {

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
