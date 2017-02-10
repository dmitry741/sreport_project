using System;
using System.Windows.Forms;
using SReport_Utility;
using System.Data;

namespace StrengthReport
{
    public partial class frmSave : Form
    {
        public frmSave()
        {
            InitializeComponent();
        }

        string m_name = string.Empty;
        string m_comment = string.Empty;
        string m_lastError = string.Empty;

        public string ReportName => m_name;

        public string Comment => m_comment;

        #region === private ===

        bool IsValid(string reportName)
        {
            bool result = false;

            do
            {
                if (reportName.Length < 3)
                {
                    m_lastError = "Название отчета слишком короткое.";
                    break;
                }

                SReportDataBase db = new SReportDataBase();
                string startupPath = Application.StartupPath + "\\db.xml";

                if (System.IO.File.Exists(startupPath))
                {
                    db.Open(startupPath);
                    bool bEqual = false;

                    for (int i = 0; i < db.count; i++)
                    {
                        DataRow dr = db.GetDataRow(i);
                        string name = (string)dr["BD_Name"];

                        if (name == reportName)
                        {
                            m_lastError = "Расчет с таким названием уже существует. Введите другое имя.";
                            bEqual = true;
                            break;
                        }
                    }

                    if (bEqual)
                        break;

                    result = true;
                }
            } while (false);

            return result;
        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            m_name = textBox1.Text;
            m_comment = textBox2.Text;

            if (!IsValid(m_name))
            {
                MessageBox.Show(m_lastError, "Расчет на прочность", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            DialogResult = DialogResult.OK;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void frmSave_Load(object sender, EventArgs e)
        {
            textBox2.Text = string.Format("// --- {0} ---", DateTime.Now.ToShortDateString());
        }
    }
}