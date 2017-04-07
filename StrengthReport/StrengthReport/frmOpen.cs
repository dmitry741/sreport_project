using System;
using System.Data;
using System.Windows.Forms;
using SReport_Utility;

namespace StrengthReport
{
    public partial class frmOpen : Form
    {
        SReportDataBase m_db = new SReportDataBase();
        string m_path = "";
        string m_ReportName = "";
        string m_db_path = Application.StartupPath + "\\" + Environment.UserName + "\\db.xml";
        bool m_bUpdate = false;

        public frmOpen()
        {
            InitializeComponent();
        }

        public string Path => m_path;

        public string ReportName => m_ReportName;

        private void button1_Click(object sender, EventArgs e)
        {
            int index = listBox1.SelectedIndex;

            if (index >= 0)
            {
                DataRow dr = m_db.GetDataRow(index);
                m_path = (string)dr["BD_Path"];
                m_ReportName = (string)dr["BD_Name"];

                // close with ok
                DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("Расчет не выбран. Чтобы открыть ранее созданный расчет, выберите его в списке и снова нажмите Открыть.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int index = listBox1.SelectedIndex;

            if (index >= 0)
            {
                DataRow dr = m_db.GetDataRow(index);
                string Name = (string)dr["BD_Name"];

                if (MessageBox.Show("Вы действительно хотите удалить расчет: " + Name, KitConstant.softwareName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    string pathToRemove = (string)dr["BD_Path"];

                    if (System.IO.File.Exists(pathToRemove))
                    {
                        System.IO.File.Delete(pathToRemove);
                        listBox1.Items.RemoveAt(index);
                        m_db.RemoveAt(index);
                        textBox2.Text = "";
                        m_bUpdate = true;
                    }
                    else
                    {
                        MessageBox.Show("Невозможно удалить расчет. Файл отсутствует", KitConstant.softwareName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                }
            }
            else
            {
                MessageBox.Show("Расчет не выбран. Чтобы удалить ранее созданный расчет, выберите его в списке и снова нажмите Удалить.");
            }
        }

        private void frmOpen_Load(object sender, EventArgs e)
        {            
            if (System.IO.File.Exists(m_db_path))
            {
                m_db.Open(m_db_path);

                for (int i = 0; i < m_db.count; i++)
                {
                    DataRow dr = m_db.GetDataRow(i);
                    listBox1.Items.Add((string)dr["BD_Name"]);
                }
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = listBox1.SelectedIndex;

            if (index >= 0)
            {
                DataRow dr = m_db.GetDataRow(index);
                textBox2.Text = (string)dr["BD_Comment"];
            }
        }

        private void frmOpen_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (m_bUpdate)
            {
                m_db.Save(m_db_path);
            }
        }

        private void frmOpen_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                int index = listBox1.SelectedIndex;

                if (index >= 0)
                {
                    DataRow dr = m_db.GetDataRow(index);
                    m_path = (string)dr["BD_Path"];
                    m_ReportName = (string)dr["BD_Name"];

                    // close with ok
                    DialogResult = DialogResult.OK;
                }
            }
        }
    }
}