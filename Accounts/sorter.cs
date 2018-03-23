using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Accounts
{
    public class sorter : IComparer
    {
        public int Column { get; set; }
    //    public SortOrder Order { get; set; }
        public sorter (int colIndex)
        {
            Column = colIndex;
     //       Order = SortOrder.None;
        }

        public int Compare(object a, object b)
        {
            int result;
            ListViewItem itemA = a as ListViewItem;
            ListViewItem itemB = b as ListViewItem;
            if (itemA == null && itemB == null)
                result = 0;
            else if (itemA == null)
                result = -1;
            else if (itemB == null)
                result = 1;
            if (itemA == itemB)
                result = 0;
            try
            {
                result = DateTime.Compare((DateTime.ParseExact(itemA.SubItems[Column].Text, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture)), DateTime.ParseExact(itemB.SubItems[Column].Text, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture));
            }
            catch (Exception)
            {
                try
                {
                    result = (Convert.ToInt32(itemA.SubItems[Column].Text).CompareTo( Convert.ToInt32(itemB.SubItems[Column].Text)));
                }
                catch (Exception)
                {
                    result = String.Compare(itemA.SubItems[Column].Text, itemB.SubItems[Column].Text);
                }
            }

            //if (Order == SortOrder.Ascending)
            //    // Invert the value returned by Compare.
            //    result *= -1;
            return result;
        }
    }
}
