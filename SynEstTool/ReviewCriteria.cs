using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SynEstTool
{
    public partial class ReviewCriteria : Form
    {
        public ReviewCriteria()
        {
            InitializeComponent();
            foreach (Control control in this.Controls)
            {
                if (control is System.Windows.Forms.TextBox)
                {
                    control.Text = ConfigurationManager.AppSettings.Get(control.Name);
                }
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            Configuration config = configuration;
            foreach (Control control in this.Controls)
            {
                if (control is System.Windows.Forms.TextBox)
                {
                    config.AppSettings.Settings[control.Name].Value = control.Text;
                }
            }
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
            this.Dispose();
            this.Close();
        }
    }
}
