Option Explicit On 

Imports System
Imports System.IO
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.Design
Imports System.Drawing


Friend Class clsDesigner

    Inherits System.Web.UI.Design.ControlDesigner

    Friend Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Overrides Function GetDesignTimeHtml() As String
        Dim ctl As ActiveGanttVBACtl = CType(Me.Component, ActiveGanttVBACtl)
        Dim sw As New StringWriter()
        Dim tw As New HtmlTextWriter(sw)
        Dim placeholderlink As New HyperLink()
        placeholderlink.Width = ctl.Width
        placeholderlink.Height = ctl.Height
        placeholderlink.BorderStyle = BorderStyle.Solid
        placeholderlink.BorderColor = Color.Gray
        placeholderlink.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(2)
        placeholderlink.Text = "<p align=""left""><b>ActiveGantt Scheduler Component for ASP.Net</b></p><p align=""left"">Visual Basic .Net Version " & ctl.Version & "<p>"
        'placeholderlink.NavigateUrl = "http://www.sourcecodestore.com"
        placeholderlink.RenderControl(tw)
        Return sw.ToString()
    End Function


End Class
