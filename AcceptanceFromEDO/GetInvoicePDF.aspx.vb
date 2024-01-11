Imports Diadoc.Api

Public Class GetInvoicePDF
    Inherits System.Web.UI.Page

    Public MyApi As DiadocApi
    Public MyToken As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка страницы
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyInvoice As Diadoc.Api.Proto.Documents.Document
        Dim MI As Object
        Dim MyBytes As Byte()
        Dim FileName As String

        If Session("Login") = "" Then
            Response.Redirect("Login.aspx", True)
        End If

        If IsNothing(Session("MyApi")) = False Then
            MyApi = Session("MyApi")
        End If

        If IsNothing(Session("MyToken")) = False Then
            MyToken = Session("MyToken")
        End If

        If IsNothing(Session("InvoiceToPDFDoc")) = False Then
            MyInvoice = Session("InvoiceToPDFDoc")
        End If

        If Not IsPostBack Then
            For i As Integer = 0 To 50
                MI = MyApi.GeneratePrintForm(MyToken, Declarations.MyboxId, MyInvoice.MessageId, MyInvoice.EntityId)
                If MI.HasContent = False Then
                    System.Threading.Thread.Sleep((MI.RetryAfter + 1) * 1000)
                Else
                    Exit For
                End If
            Next

            Try
                MyBytes = MI.Content.Bytes
                FileName = MyInvoice.DocumentNumber + ".pdf"
                Response.Clear()
                Response.ClearHeaders()
                Response.ClearContent()
                Response.AddHeader("content-disposition", "filename=" + FileName)
                Response.ContentType = "application/pdf"
                Response.AddHeader("Content-Length", MyBytes.Length)
                Response.BinaryWrite(MyBytes)
                Response.End()
            Catch ex As Exception
                If My.Settings.MyDebug = "YES" Then
                    EventLog.WriteEntry("AcceptanceFromEDO", ex.Message)
                End If
            End Try
        End If
    End Sub

End Class