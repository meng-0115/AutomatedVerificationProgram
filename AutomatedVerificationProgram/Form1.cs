using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutomatedVerificationProgram
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        

        private void btnAutomatedBatchVerificationPage_Click(object sender, EventArgs e)
        {
            #region 頁面切換
            AutomatedBatchVerificationPage page = new AutomatedBatchVerificationPage();
            page.Location = new Point(0, 0);
            page.Size = pnlMainContent.ClientSize;
            page.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            pnlMainContent.Controls.Clear();
            pnlMainContent.Controls.Add(page);
            #endregion

        }

        private void btnManualSingleVerificationPage_Click(object sender, EventArgs e)
        {
            #region 頁面切換
            ManualSingleVerificationPage page = new ManualSingleVerificationPage();
            page.Location = new Point(0, 0);
            page.Size = pnlMainContent.ClientSize;
            page.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            pnlMainContent.Controls.Clear();
            pnlMainContent.Controls.Add(page);
            #endregion
        }

        private void btnManualBatchVerificationPage_Click(object sender, EventArgs e)
        {
            #region 頁面切換
            ManualBatchVerificationPage page = new ManualBatchVerificationPage();
            page.Location = new Point(0, 0);
            page.Size = pnlMainContent.ClientSize;
            page.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            pnlMainContent.Controls.Clear();
            pnlMainContent.Controls.Add(page);
            #endregion
        }

        private void btnSystemSettingsPage_Click(object sender, EventArgs e)
        {

            #region 頁面切換
            SystemSettingsPage page = new SystemSettingsPage();
            page.Location = new Point(0, 0);
            page.Size = pnlMainContent.ClientSize;
            page.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            pnlMainContent.Controls.Clear();
            pnlMainContent.Controls.Add(page);
            #endregion/**/
        }
    }
}
