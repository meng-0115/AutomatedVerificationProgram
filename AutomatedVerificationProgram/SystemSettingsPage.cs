using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;
using static AutomatedVerificationProgram.McdCustomVerificationSequencePage;

namespace AutomatedVerificationProgram
{
    public partial class SystemSettingsPage : UserControl
    {
        public SystemSettingsPage()
        {
            InitializeComponent();
        }
               
        private void btnOutdatedCoordinatesSetting_Click(object sender, EventArgs e)
        {
            OutdatedCoordinatesSetting OutdatedCoordinatesSetting = new OutdatedCoordinatesSetting();
            OutdatedCoordinatesSetting.ShowDialog();
        }

        private void btnmcdCoordinatesSetting_Click(object sender, EventArgs e)
        {
            McdCoordinatesSettingcs McdCoordinatesSetting = new McdCoordinatesSettingcs();
            McdCoordinatesSetting.ShowDialog();
        }

        private void btnOcrSetting_Click(object sender, EventArgs e)
        {
            OcrSettingPage ocrSettingPage = new OcrSettingPage();
            ocrSettingPage.ShowDialog();
        }
    }
}
