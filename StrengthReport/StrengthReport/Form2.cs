using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace StrengthReport
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        #region === Members ===

        System.Collections.ArrayList m_list = null;

        #endregion

        #region === Public ===

        public void SetData(ref System.Collections.ArrayList list)
        {
            m_list = list;
        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            if (m_list != null)
            {
                foreach (string item in m_list)
                {
                    listBox1.Items.Add(item);
                }
            }
        }
    }
}