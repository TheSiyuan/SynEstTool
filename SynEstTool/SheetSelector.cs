using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SynEstTool
{
    public partial class SheetSelector : Form
    {
        public SheetSelector(String[] Alist)
        {
            InitializeComponent();
            Select_No.Items.Clear();
            foreach (var item in Alist)
            {
                Select_No.Items.Add(item);
            }
        }
        private void MoveListBoxItems(ListBox source, ListBox destination)
        {
            ListBox.SelectedObjectCollection sourceItems = source.SelectedItems;
            foreach (var item in sourceItems)
            {
                destination.Items.Add(item);
            }
            while (source.SelectedItems.Count > 0)
            {
                source.Items.Remove(source.SelectedItems[0]);
            }
        }
        private void BtnItemMoveRight_Click(object sender, EventArgs e)
        {
            MoveListBoxItems(Select_No, Select_Yes);
        }

        private void BtnItemMoveLeft_Click(object sender, EventArgs e) => MoveListBoxItems(Select_Yes, Select_No);

        private void BtnConsolidate_Click(object sender, EventArgs e) => this.Select_No.Items.Add(1);

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Dispose();           
            this.Close();
        }

        private void SheetSelector_Load(object sender, EventArgs e)
        {

        }
    }
}
