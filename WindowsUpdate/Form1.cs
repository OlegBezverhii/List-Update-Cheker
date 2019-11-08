using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsUpdate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            List<string> result = new List<string>(200);
            if (radioButton1.Checked) result = WindowsUpdate.Update.listUpdateHistory();
            if (radioButton2.Checked) result = WindowsUpdate.Update.DISMlist();
            if (radioButton3.Checked) result = WindowsUpdate.Update.XMLlist();
            if (radioButton4.Checked) result = WindowsUpdate.Update.Sessionlist("localhost");

            for (int i = 0; i < result.Count; i++)
                textBox1.Text += result[i].ToString() + Environment.NewLine;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox4.Clear();
            List<string> result = new List<string>(200);

            string login = loginBox.Text.ToString();
            string password = passwordBox.Text.ToString();
            string namepc = textBox5.Text.ToString();

            if (radioButton5.Checked) result = WindowsUpdate.Update.GetWMIlist(namepc, login, password);
            if (radioButton6.Checked) result = WindowsUpdate.Update.GetWSUSlist(namepc, login, password);

            for (int i = 0; i < result.Count; i++)
                textBox4.Text += result[i].ToString() + Environment.NewLine;
        }
    }
}
