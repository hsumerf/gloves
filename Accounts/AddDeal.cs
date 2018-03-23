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
    public partial class AddDeal : Form
    {
        
        public AddDeal()
        {
           
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
            if ( nameBox.Text == "" || valQty.Value == 0 || ValPer.Value == 0 || valQty.Text == ""|| ValPer.Text == "")
            {
                MessageBox.Show("Few fields are missiing","Incomplete", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (labelitems.Text !="")
            if (valQty.Value > Convert.ToInt32 (labelitems.Text))
            {
                MessageBox.Show("Out of stock!","No stock",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }

            //if (valQty.Text == "")
            //{
            //    valQty.Value = 0;
            //}
            //if (ValPer.Text == "")
            //{
            //    ValPer.Value = 0;
            //}


            XmlDocument doc = new XmlDocument();
            doc.Load("main.dbs");
            string lastId; 
            try
            {
               lastId = doc.DocumentElement.LastChild.Attributes[0].Value;
            }
            catch (Exception)
            {

                lastId = "-1";
            }            

            XmlElement rootElement = doc.DocumentElement;
            XmlElement newitem = doc.CreateElement("person");
            newitem.SetAttribute("id", (Convert.ToInt32(lastId) + 1).ToString());
            XmlElement date = doc.CreateElement("date");
            date.InnerText = dateTimePicker1.Text;
            XmlElement name = doc.CreateElement("name");
            name.InnerText = nameBox.Text;
            XmlElement company = doc.CreateElement("company");
            company.InnerText = companyBox.Text;
            XmlElement item = doc.CreateElement("item");
            item.InnerText = itemBox.Text;
            XmlElement qty = doc.CreateElement("qty");
            qty.InnerText = valQty.Value.ToString();
            XmlElement peritem = doc.CreateElement("peritem");
            peritem.InnerText = ValPer.Value.ToString();
            XmlElement Total = doc.CreateElement("total");
            Total.InnerText = (double.Parse(valQty.Text) * double.Parse(ValPer.Value.ToString())).ToString();
            XmlElement recv = doc.CreateElement("recieve");
            recv.InnerText = "0";
            XmlElement details = doc.CreateElement("details");
            details.InnerText = textBox1.Text;

            newitem.AppendChild(date);
            newitem.AppendChild(name);
            newitem.AppendChild(company);
            newitem.AppendChild(item);
            newitem.AppendChild(qty);
            newitem.AppendChild(peritem);
            newitem.AppendChild(Total);
            newitem.AppendChild(recv);
            newitem.AppendChild(details);

            rootElement.AppendChild(newitem);
            doc.Save("main.dbs");
            subtractStock();
            RefreshStockLabel();
            MessageBox.Show("Data saved successfully", "Saved!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            valQty.Text = "";
            ValPer.Text = "";
            textBox1.Text = "";
            total.Text = "Total: ";
        }

        private void subtractStock()
        {
          
            try
            {
                int Before;
                int After;
                XmlDocument doc = new XmlDocument();
                doc.Load("stocks.dbs");

                Before = Convert.ToInt32(doc.SelectSingleNode("//company[@name='" + companyBox.Text + "']" + "//item[@name='" + itemBox.Text + "']").InnerText);
                After = Before - (int)valQty.Value;
                doc.SelectSingleNode("//company[@name='" + companyBox.Text + "']" + "//item[@name='" + itemBox.Text + "']").InnerText = After.ToString();
                doc.Save("stocks.dbs");
            }
            catch (Exception)
            {
            }
            try
            {
                double Before;
                double After;
                XmlDocument docC = new XmlDocument();
                docC.Load("clients.dbs");
                Before = Convert.ToDouble(docC.SelectSingleNode("//Client[@name='" + nameBox.Text + "']" + "//total").InnerText);
                After = Before + (double)valQty.Value * (double)ValPer.Value;
                docC.SelectSingleNode("//Client[@name='" + nameBox.Text + "']" + "//total").InnerText = After.ToString();
                docC.Save("clients.dbs");
            }
            catch (Exception)
            {
            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
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

        private void AddDeal_Load(object sender, EventArgs e)
        {
            valQty.Text = "";
            ValPer.Text = "";
            XmlDocument doc = new XmlDocument();
            doc.Load("stocks.dbs");
            XmlElement rootElement = doc.DocumentElement;
            var list = rootElement.ChildNodes;

            for (int i = 0; i < list.Count; i++)
            {

                companyBox.Items.Add(list.Item(i).Attributes[0].InnerText);
            }
            companyBox.Items.Add("");

            try
            {
                nameBox.Items.Clear();
                doc.Load("clients.dbs");
                rootElement = doc.DocumentElement;
                list = rootElement.ChildNodes;
                for (int i = 0; i < list.Count; i++)
                    nameBox.Items.Add(list.Item(i).Attributes[0].InnerText);
            }
            catch (Exception)
            {

            }
          
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshStockLabel();
        }

        private void RefreshStockLabel() {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load("stocks.dbs");
                labelitems.Text = doc.SelectSingleNode("//company[@name='" + companyBox.Text + "']" + "//item[@name='" + itemBox.Text + "']").InnerText;
            }
            catch (Exception)
            {

                labelitems.Text = "";
            }

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void ValPer_ValueChanged(object sender, EventArgs e)
        {
            total.Text = "Total: "+(valQty.Value * ValPer.Value).ToString();
        }

        private void valQty_ValueChanged(object sender, EventArgs e)
        {
            total.Text = "Total: " + (valQty.Value * ValPer.Value).ToString();
        }

        private void valQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Decimal)
                e.SuppressKeyPress = true;
        }

        private void ValPer_KeyDown(object sender, KeyEventArgs e)
        {
        //    if (e.KeyCode == Keys.Decimal)
        //        e.SuppressKeyPress = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void nameBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            string strFindStr = "";
            if (e.KeyChar == (char)8)
            {
                if (cb.SelectionStart <= 1)
                {
                    cb.Text = "";
                    return;
                }

                if (cb.SelectionLength == 0)
                    strFindStr = cb.Text.Substring(0, cb.Text.Length - 1);
                else
                    strFindStr = cb.Text.Substring(0, cb.SelectionStart - 1);
            }
            else
            {
                if (cb.SelectionLength == 0)
                    strFindStr = cb.Text + e.KeyChar;
                else
                    strFindStr = cb.Text.Substring(0, cb.SelectionStart) + e.KeyChar;
            }
            int intIdx = -1;
            // Search the string in the ComboBox list.
            intIdx = cb.FindString(strFindStr);
            if (intIdx != -1)
            {
                cb.SelectedText = "";
                cb.SelectedIndex = intIdx;
                cb.SelectionStart = strFindStr.Length;
                cb.SelectionLength = cb.Text.Length;
                e.Handled = true;
            }
            else
                e.Handled = true;
        }

        private void itemBox_DropDown(object sender, EventArgs e)
        {
            
        }

        private void nameBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void companyBox_KeyPress(object sender, KeyPressEventArgs e)
        {

            ComboBox cb = (ComboBox)sender;
            string strFindStr = "";
            if (e.KeyChar == (char)8)
            {
                if (cb.SelectionStart <= 1)
                {
                    cb.Text = "";
                    return;
                }

                if (cb.SelectionLength == 0)
                    strFindStr = cb.Text.Substring(0, cb.Text.Length - 1);
                else
                    strFindStr = cb.Text.Substring(0, cb.SelectionStart - 1);
            }
            else
            {
                if (cb.SelectionLength == 0)
                    strFindStr = cb.Text + e.KeyChar;
                else
                    strFindStr = cb.Text.Substring(0, cb.SelectionStart) + e.KeyChar;
            }
            int intIdx = -1;
            // Search the string in the ComboBox list.
            intIdx = cb.FindString(strFindStr);
            if (intIdx != -1)
            {
                cb.SelectedText = "";
                cb.SelectedIndex = intIdx;
                cb.SelectionStart = strFindStr.Length;
                cb.SelectionLength = cb.Text.Length;
                e.Handled = true;
            }
            else
                e.Handled = true;
        }

        private void itemBox_KeyPress(object sender, KeyPressEventArgs e)
        {

            ComboBox cb = (ComboBox)sender;
            string strFindStr = "";
            if (e.KeyChar == (char)8)
            {
                if (cb.SelectionStart <= 1)
                {
                    cb.Text = "";
                    return;
                }

                if (cb.SelectionLength == 0)
                    strFindStr = cb.Text.Substring(0, cb.Text.Length - 1);
                else
                    strFindStr = cb.Text.Substring(0, cb.SelectionStart - 1);
            }
            else
            {
                if (cb.SelectionLength == 0)
                    strFindStr = cb.Text + e.KeyChar;
                else
                    strFindStr = cb.Text.Substring(0, cb.SelectionStart) + e.KeyChar;
            }
            int intIdx = -1;
            // Search the string in the ComboBox list.
            intIdx = cb.FindString(strFindStr);
            if (intIdx != -1)
            {
                cb.SelectedText = "";
                cb.SelectedIndex = intIdx;
                cb.SelectionStart = strFindStr.Length;
                cb.SelectionLength = cb.Text.Length;
                e.Handled = true;
            }
            else
                e.Handled = true;
        }

        private void itemBox_Enter(object sender, EventArgs e)
        {
            itemBox.Items.Clear();
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load("stocks.dbs");
                var node = doc.SelectSingleNode("//company[@name='" + companyBox.Text + "']");
                var list = node.ChildNodes;

                for (int i = 0; i < list.Count; i++)
                {
                    itemBox.Items.Add(list.Item(i).Attributes[0].InnerText);

                }
            }
            catch (Exception)
            {
            }
            itemBox.Items.Add("");
        }
    }
}
