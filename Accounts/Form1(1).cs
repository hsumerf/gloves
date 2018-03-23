using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace Accounts
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //public void refreshList()
        //{
        //    XmlDocument doc = new XmlDocument();
        //    doc.Load("main.dbs");
            
        //    //Display all the book titles.
        //    XmlNodeList elemList = doc.GetElementsByTagName("person");
        //    //if (elemList.Item(0).Attributes[0].Value == "13")
        //    //{
        //    //    MessageBox.Show(elemList.Item(1).SelectSingleNode("//name").InnerText);
        //    //    MessageBox.Show("SAd");
        //    //}

        //    //  MessageBox.Show(doc.SelectSingleNode("//person[@id='14']//name").InnerText);
        //    //var ff = doc.GetElementById("12");
        //    // MessageBox.Show ( ff.OuterXml);

        //    // MessageBox.Show (doc.SelectSingleNode("//d221/name").InnerText);
        //    listView1.BeginUpdate();
        //    listView1.Items.Clear();
        //    for (int i = 0; i < elemList.Count; i++)
        //    {

        //        string id = elemList.Item(i).Attributes[0].Value;
        //        string date = elemList.Item(i).ChildNodes[0].InnerText;
        //        string name = elemList.Item(i).ChildNodes[1].InnerText;
        //        string company = elemList.Item(i).ChildNodes[2].InnerText;
        //        string _item = elemList.Item(i).ChildNodes[3].InnerText;
        //        string qty = elemList.Item(i).ChildNodes[4].InnerText;
        //        string peritem = elemList.Item(i).ChildNodes[5].InnerText;
        //        string recv = elemList.Item(i).ChildNodes[6].InnerText;
        //        string details = elemList.Item(i).ChildNodes[7].InnerText;

        //        var item = (new ListViewItem(new[] { id, date, name, company, _item, qty, peritem, (Convert.ToInt32(qty) * Convert.ToInt32(peritem)).ToString(), recv, details }));
        //        listView1.Items.Add(item);
        //    }
        //    listView1.EndUpdate();
        //}

       

        private void button4_Click_1(object sender, EventArgs e)
        {
            AddDeal adddeal = new AddDeal();
            adddeal.ShowDialog();
        }


        private void button5_Click(object sender, EventArgs e)
        {
            stocks stocks = new stocks();
            stocks.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            recieve recv = new recieve();
            recv.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Lesasure les = new Lesasure();
            les.ShowDialog();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Lesasure ledger = new Lesasure();
            ledger.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            balSheet balDialoge = new balSheet();
            balDialoge.ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            StockSheet sksheet = new StockSheet();
            sksheet.ShowDialog();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            DailyRPT dailyrpt = new DailyRPT();
            dailyrpt.ShowDialog();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (Accounts.Properties.Settings.Default.first == true)
            {
                Password pas = new Password();
                pas.ShowDialog();
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            AddDeal addnew = new AddDeal();
            addnew.Show(this);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            recieve recv = new recieve();
            recv.Show(this);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            stocks stockmanager = new stocks();
            stockmanager.Show(this);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            StockSheet SS = new StockSheet();
            SS.Show(this);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            DailyRPT DailyRPT = new DailyRPT();
            DailyRPT.Show(this);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            balSheet  balSheet = new balSheet();
            balSheet.Show(this);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            Clients clint = new Clients();
            clint.Show(this);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            Backup bckup = new Backup();
            bckup.Show(this);
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
                b1.PerformClick();
            else if (e.KeyCode == Keys.F2)
                b2.PerformClick();
            else if (e.KeyCode == Keys.F3)
                b3.PerformClick();
            else if (e.KeyCode == Keys.F4)
                b4.PerformClick();
            else if (e.KeyCode == Keys.F5)
                b5.PerformClick();
            else if (e.KeyCode == Keys.F6)
                b6.PerformClick();
            else if (e.KeyCode == Keys.F7)
                b7.PerformClick();
            else if (e.KeyCode == Keys.F8)
                b8.PerformClick();
            else if (e.KeyCode == Keys.F9)
                b9.PerformClick();
            else if (e.KeyCode == Keys.F10)
               Process.Start("calc.exe");

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
