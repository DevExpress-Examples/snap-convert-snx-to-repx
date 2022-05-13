Imports DevExpress.XtraReports.UI
Imports System
Imports System.Windows.Forms
Imports SnxToRepx.Converter

Namespace SnxToRepx

    Public Partial Class MainForm
        Inherits Form

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub bedtLoadSnx_ButtonClick(ByVal sender As Object, ByVal e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs)
            Using dialog As OpenFileDialog = New OpenFileDialog()
                dialog.Filter = "SNX files|*.snx"
                If dialog.ShowDialog() = DialogResult.OK Then bedtLoadSnx.Text = dialog.FileName
            End Using
        End Sub

        Private Sub bedtSaveRepx_ButtonClick(ByVal sender As Object, ByVal e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs)
            Using dialog As SaveFileDialog = New SaveFileDialog()
                dialog.Filter = "REPX files|*.repx"
                If dialog.ShowDialog() = DialogResult.OK Then bedtSaveRepx.Text = dialog.FileName
            End Using
        End Sub

        Private Sub btnConvert_Click(ByVal sender As Object, ByVal e As EventArgs)
            If IO.File.Exists(bedtLoadSnx.Text) Then
                Dim report As XtraReport = Nothing
                Using converter As SnapReportConverter = New SnapReportConverter()
                    AddHandler converter.ConfigureDataConnection, AddressOf Converter_ConfigureDataConnection
                    report = converter.Convert(bedtLoadSnx.Text)
                End Using

                If report IsNot Nothing Then
                    If Not String.IsNullOrEmpty(bedtSaveRepx.Text) Then report.SaveLayout(bedtSaveRepx.Text)
                    If chkDesigner.Checked Then
                        Using designTool As ReportDesignTool = New ReportDesignTool(report)
                            designTool.ShowRibbonDesignerDialog()
                        End Using
                    End If
                End If
            End If
        End Sub

        Private Sub Converter_ConfigureDataConnection(ByVal sender As Object, ByVal e As DevExpress.DataAccess.Sql.ConfigureDataConnectionEventArgs)
        'Insert your code to establish the data connection.
        'For more information, review the following topic:
        'https://docs.devexpress.com/CoreLibraries/DevExpress.DataAccess.Sql.SqlDataSource.ConfigureDataConnection
        End Sub
    End Class
End Namespace
