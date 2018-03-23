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
    public partial class Update : Form
    {
        stocks stocks;
        public Update(stocks stocksFrom)
        {
            InitializeComponent();
            stocks = stocksFrom;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (ValPer.Text == "")
            {
                ValPer.Value = 0;
            }
            XmlDocument stockdoc = new XmlDocument();
            stockdoc.Load("stocks.dbs");
            string qtybefore = stockdoc.SelectSingleNode("//company[@name='" + stocks.comboBox1.Text + "']" + "//item[@name='" + itemLabel.Text + "']").InnerText;
            string qtyafter = (Convert.ToInt32(qtybefore) + ValPer.Value).ToString();
            stockdoc.SelectSingleNode("//company[@name='" + stocks.comboBox1.Text + "']" + "//item[@name='" + itemLabel.Text + "']").InnerText = qtyafter;
            stockdoc.Save("stocks.dbs");
            stocks.RefreshList();
            this.Close();
        }

        private void Update_Load(object sender, EventArgs e)
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
