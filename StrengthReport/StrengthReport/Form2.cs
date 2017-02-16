using System;
using System.Windows.Forms;

namespace StrengthReport
{
    public partial class Form2 : Form
    {
        public Form2(System.Collections.ArrayList list)
        {
            InitializeComponent();
            m_list = list;
        }

        #region === members ===

        System.Collections.ArrayList m_list = null;

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
                listBox1.Items.AddRange(m_list.ToArray());
            }

            this.Text = string.Format("{0}: {1}", SReport_Utility.KitConstant.companyName, SReport_Utility.KitConstant.softwareName);
        }
    }
}