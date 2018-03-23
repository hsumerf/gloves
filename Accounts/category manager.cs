using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Accounts
{
    public partial class category_manager : Form
    {
        Clients clints;
        public category_manager(Clients clnt)
        {
            InitializeComponent();
            clints = clnt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                return;
            }
            System.IO.File.AppendAllText("category.dbs",textBox1.Text + Environment.NewLine);
            clints.reload();
            reload();
        }

        private void category_manager_Load(object sender, EventArgs e)
        {
            reload();
        }
        private void reload()
        {
            listView1.Items.Clear();
            var list = System.IO.File.ReadAllLines("category.dbs");
            for (int i = 0; i < list.Length; i++)
            {
                listView1.Items.Add(list[i]);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < listView1.SelectedItems.Count; i++)
            {
                if (listView1.SelectedItems.Count == 0)
                {
                    MessageBox.Show("No item selected, Please select item to delete.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                listView1.SelectedItems[i].Remove();
                System.IO.File.WriteAllLines("category.dbs", listView1.Items.Cast<ListViewItem>().Select(x => x.Text).ToList());

            }
        }
    }
}
