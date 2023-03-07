Public Class frmShowHtml
    Public Sub New(strHtml As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        wb.ScriptErrorsSuppressed = True
        wb.DocumentText = strHtml
    End Sub
End Class