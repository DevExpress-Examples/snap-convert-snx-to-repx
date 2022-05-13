Imports System
Imports System.Windows.Forms

Namespace SnxToRepx

    Friend Module Program

        ''' <summary>
        ''' The main entry point for the application.
        ''' </summary>
        <STAThread>
        Sub Main()
            Threading.Thread.CurrentThread.CurrentCulture = New Globalization.CultureInfo("en-US")
            Threading.Thread.CurrentThread.CurrentUICulture = New Globalization.CultureInfo("en-US")
            Call Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)
            Call Application.Run(New MainForm())
        End Sub
    End Module
End Namespace
