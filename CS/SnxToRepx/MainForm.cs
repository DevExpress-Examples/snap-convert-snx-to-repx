using DevExpress.XtraReports.UI;
using System;
using System.Windows.Forms;
using SnxToRepx.Converter;

namespace SnxToRepx
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }        

        private void bedtLoadSnx_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "SNX files|*.snx";
                if (dialog.ShowDialog() == DialogResult.OK)
                    bedtLoadSnx.Text = dialog.FileName;
            }
        }

        private void bedtSaveRepx_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "REPX files|*.repx";
                if (dialog.ShowDialog() == DialogResult.OK)
                    bedtSaveRepx.Text = dialog.FileName;
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(bedtLoadSnx.Text))
            {
                XtraReport report = null;
                using (SnapReportConverter converter = new SnapReportConverter())
                {
                    converter.ConfigureDataConnection += Converter_ConfigureDataConnection;
                    report = converter.Convert(bedtLoadSnx.Text);
                }

                if (report != null)
                {
                    if (!string.IsNullOrEmpty(bedtSaveRepx.Text))
                        report.SaveLayout(bedtSaveRepx.Text);

                    if (chkDesigner.Checked)
                    {
                        using (ReportDesignTool designTool = new ReportDesignTool(report))
                        {
                            designTool.ShowRibbonDesignerDialog();
                        }
                    }
                }
            }
        }

        private void Converter_ConfigureDataConnection(object sender, DevExpress.DataAccess.Sql.ConfigureDataConnectionEventArgs e)
        {
            //Insert your code to establish the data connection.
            //For more information, review the following topic:
            //https://docs.devexpress.com/CoreLibraries/DevExpress.DataAccess.Sql.SqlDataSource.ConfigureDataConnection
        }
    }
}
