Imports System.Web

Friend Class clsAGVBARenderStream
    Implements System.Web.IHttpModule

    Public Const ImageHandlerRequestFilename As String = "image_stream1000.aspx"
    Public Const ImageNamePrefix As String = "i_m_g"


    Public Sub Dispose() Implements System.Web.IHttpModule.Dispose

    End Sub

    Public Sub Init(ByVal context As System.Web.HttpApplication) Implements System.Web.IHttpModule.Init
        AddHandler context.BeginRequest, AddressOf Me.OnBeginRequest
    End Sub

    Public Sub OnBeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        Dim httpApp As HttpApplication = sender

        Dim oCtrl As ActiveGanttVBACtl = Nothing

        If (httpApp.Request.Path.ToLower().IndexOf(ImageHandlerRequestFilename) <> -1) Then

            oCtrl = httpApp.Application(ImageNamePrefix & httpApp.Request.QueryString("id"))
            If (oCtrl Is Nothing) Then
                Return
            Else
                Dim memStream As System.IO.MemoryStream = oCtrl.OnPaint()
                memStream.WriteTo(httpApp.Context.Response.OutputStream)
                memStream.Close()

                httpApp.Context.ClearError()
                httpApp.Context.Response.ContentType = "image/png"
                httpApp.Response.StatusCode = 200
                httpApp.Application.Remove(ImageNamePrefix & httpApp.Request.QueryString("id"))
                'httpApp.Response.End()
                httpApp.Context.ApplicationInstance.CompleteRequest()
            End If
        End If
    End Sub

End Class
