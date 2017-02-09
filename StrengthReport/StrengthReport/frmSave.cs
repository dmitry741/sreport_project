using System;
using System.Windows.Forms;

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

        public string ReportName => m_name;

        public string Comment => m_comment;

        private void button1_Click(object sender, EventArgs e)
        {
            m_name = textBox1.Text;
            m_comment = textBox2.Text;

            if (m_name.Length < 3)
            {
                MessageBox.Show("Название отчета должно содержать как минимум 3 символа.", "Расчет на прочность", MessageBoxButtons.OK, MessageBoxIcon.Stop);
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