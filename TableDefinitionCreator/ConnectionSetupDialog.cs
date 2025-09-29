using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using TableDefinitionCreator.utils;

namespace TableDefinitionCreator
{
    public partial class ConnectionSetupDialog : Form
    {
        public string IP => txtIP.Text.Trim();
        public string Database => txtDatabase.Text.Trim();
        public string UserID => txtID.Text.Trim();
        public string Password => txtPW.Text.Trim();

        public ConnectionSetupDialog()
        {
            InitializeComponent();
        }
        private void ConnectionSetupDialog_Load(object sender, EventArgs e)
        {
            string conStr = ConfigManager.GetConnectionString();
            if (string.IsNullOrEmpty(conStr) == false)
            {
                var builder = new SqlConnectionStringBuilder(conStr);
                txtIP.Text = builder.DataSource;
                txtDatabase.Text = builder.InitialCatalog;
                txtID.Text = builder.UserID;
                txtPW.Text = builder.Password;
            }

            txtIP.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(AllFieldsFilled() == false)
            {
                MessageBox.Show("모든 필드를 입력해주세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ConfigManager.SetConnectionString(IP, Database, UserID, Password);

            bool isConnected = DbAccess.IsConn(out string errMsg);
            if(isConnected == false)
            {
                MessageBox.Show($"데이터베이스에 연결할 수 없습니다.\n입력한 정보를 다시 확인해주세요.\n\n{errMsg}", "연결 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DialogResult = DialogResult.OK;
        }

        private bool AllFieldsFilled()
        {
            return !string.IsNullOrEmpty(IP)
                && !string.IsNullOrEmpty(Database)
                && !string.IsNullOrEmpty(UserID)
                && !string.IsNullOrEmpty(Password);
        }

        private void txtPW_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }
    }
}
