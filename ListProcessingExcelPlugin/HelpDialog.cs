using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ListProcessingExcelPlugin
{
    public partial class HelpDialog : Form
    {
        public HelpDialog()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void HelpDialog_Load(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void easterEggButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("The column numbers don't have to be equal (e.g. [a] and [a,b] \n" +
                            "which can be used to compare comma-separated items vs not. (i.e. John Doe v.s. Doe,John)");
        }

    }
}



