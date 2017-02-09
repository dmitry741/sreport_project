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
        bool m_bUpdate = false;
        string m_source_path = Application.StartupPath;

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
                MessageBox.Show("������ �� ������. ����� ������� ����� ��������� ������, �������� ��� � ������ � ����� ������� �������.");
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

                if (MessageBox.Show("�� ������������� ������ ������� ������: " + Name, "������ �� ���������", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
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
                        MessageBox.Show("���������� ������� ������. ���� ������ ��������� ��� �����������", "������ �� ���������", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                }
            }
            else
            {
                MessageBox.Show("������ �� ������. ����� ������� ����� ��������� ������, �������� ��� � ������ � ����� ������� �������.");
            }
        }

        private void frmOpen_Load(object sender, EventArgs e)
        {            
            string startupPath = m_source_path + "\\";

            if (System.IO.File.Exists(startupPath + "db.xml"))
            {
                m_db.Open(startupPath + "db.xml");

                for (int i = 0; i < m_db.Count; i++)
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
                string startupPath = m_source_path + "\\";
                m_db.Save(startupPath + "db.xml");
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