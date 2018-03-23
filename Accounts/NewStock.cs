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
    public partial class NewStock : Form
    {
        stocks stocks;
        public NewStock(stocks stocksform)
        {
            InitializeComponent();
            stocks = stocksform;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    XmlConvert.VerifyName(textBox1.Text.Replace(" ", "_"));
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("This item name is not valid, Use Alphabates only.");
            //    return;

            //}
            if (textBox1.Text == "")
                return;

            XmlDocument doc = new XmlDocument();
            doc.Load("stocks.dbs");
            if (doc.SelectSingleNode("//company[@name='" + stocks.comboBox1.Text + "']" + "//item[@name='" + textBox1.Text + "']") != null)
            {
                MessageBox.Show("Already");
                return;
            }
            if (ValPer.Text == "")
            {
                ValPer.Value = 0;
            }
            var rootElement = doc.SelectSingleNode("//company[@name='" + stocks.comboBox1.Text+ "']");
            XmlElement newItem = doc.CreateElement("item");
            newItem.SetAttribute("name", textBox1.Text);
            newItem.InnerText = ValPer.Value.ToString();
            rootElement.AppendChild(newItem);
            doc.Save("stocks.dbs");
            stocks.RefreshList();
            textBox1.Text = "";
            ValPer.Text = "";
        }

        private void NewStock_Load(object sender, EventArgs e)
        {
            ValPer.Text = "";
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Enter))
            {
                SendKeys.Send("{TAB}");
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void ValPer_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Decimal)
                e.SuppressKeyPress = true;
        }
    }
}
