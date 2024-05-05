using PP_Project.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PP_Project
{
    public partial class Form1 : Form
    {
        string destination = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrWhiteSpace(destination))
                GetWordData.InitWord(destination);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "Word | *.docx" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                destination = ofd.FileName;
                label3.Text = destination;
            }
        }

    }
}
