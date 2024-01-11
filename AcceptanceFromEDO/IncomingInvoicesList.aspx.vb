Imports System.Data.SqlClient
Imports Diadoc.Api
Imports System.Xml

Public Class IncomingInvoicesList
    Inherits System.Web.UI.Page
    'Public Shared GridView1_SelRow As Integer = 0
    'Public Shared GridView1_SelRowOld As Integer = 0
    Public Docdt As DataTable
    Public Suppdt As DataTable
    Public MyApi As DiadocApi
    Public MyToken As String = ""
    Public POrderNum As String = ""
    Public MyErrString As String = ""

    Public Sub New()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// конструктор
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Try
            Docdt = New DataTable()
            Docdt.Columns.Add("CompanyName", Type.GetType("System.String"))
            Docdt.Columns.Add("INN", Type.GetType("System.String"))
            Docdt.Columns.Add("KPP", Type.GetType("System.String"))
            Docdt.Columns.Add("DocNum", Type.GetType("System.String"))
            Docdt.Columns.Add("DocDate", Type.GetType("System.String"))
            Docdt.Columns.Add("DocSumm", Type.GetType("System.String"))
            Docdt.Columns.Add("DocVAT", Type.GetType("System.String"))
            Docdt.Columns.Add("MyDoc", Type.GetType("System.Object"))
            Docdt.Columns.Add("DocState", Type.GetType("System.String"))
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --1--> " & ex.Message)
            End If
        End Try

        Try
            Suppdt = New DataTable()
            Suppdt.Columns.Add("INN", Type.GetType("System.String"))
            Suppdt.Columns.Add("KPP", Type.GetType("System.String"))
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --2--> " & ex.Message)
            End If
            Exit Sub
        End Try
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка страницы
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Session("Login") = "" Then
            Response.Redirect("Login.aspx", True)
        End If

        If IsNothing(Session("Suppdt")) = False Then
            Suppdt = Session("Suppdt")
        End If

        If IsNothing(Session("Docdt")) = False Then
            Docdt = Session("Docdt")
        End If

        If IsNothing(Session("MyApi")) = False Then
            MyApi = Session("MyApi")
        End If

        If IsNothing(Session("MyToken")) = False Then
            MyToken = Session("MyToken")
        End If

        Page.ClientScript.GetPostBackEventReference(GridView1, "")

        If Not IsPostBack Then
            GetSupplierList()
            GetSupplierInfo()
            CheckButtonState()
            CheckButtonState2()
        End If
    End Sub

    Private Sub GetSupplierList()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение списка поставщиков с автоматической приемкой товаров и занесение 
        '// его в DropDownList с ID="SupplierList"
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                          '--рабочая строка
        Dim ds As New DataSet()

        Try
            MySQLStr = "SELECT '---' AS Code, ' Все поставщики' AS Name "
            MySQLStr = MySQLStr & "UNION ALL "
            MySQLStr = MySQLStr & "SELECT Ltrim(Rtrim(PL010300.PL01001)) AS Code, Ltrim(Rtrim(PL010300.PL01001)) + '   ' + Ltrim(Rtrim(PL010300.PL01002)) AS Name "
            MySQLStr = MySQLStr & "FROM PL010300 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_SupplierCard0300 ON PL010300.PL01001 = tbl_SupplierCard0300.PL01001 "
            MySQLStr = MySQLStr & "WHERE(tbl_SupplierCard0300.EDO = 2) "
            MySQLStr = MySQLStr & "ORDER BY Code "
            Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                Try
                    Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                        cmd.CommandType = CommandType.Text
                        Using da As New SqlDataAdapter()
                            da.SelectCommand = cmd
                            da.Fill(ds)
                            SupplierList.DataSource = ds
                            SupplierList.DataTextField = "Name"
                            SupplierList.DataValueField = "Code"
                            SupplierList.DataBind()
                            SupplierList.SelectedValue = "---"
                        End Using
                    End Using
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --3--> " & ex.Message)
                    End If
                Finally
                    MyConn.Close()
                End Try
            End Using
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --4--> " & ex.Message)
            End If
        End Try
    End Sub

    Private Sub GetInvoiceList()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получене списка инвойсов от работающих в режиме автоприемки поставщиков 
        '// (выбранного в DropDownList с ID="SupplierList" поставщика)
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDF As Diadoc.Api.DocumentsFilter
        Dim MyDocList As Diadoc.Api.Proto.Documents.DocumentList
        Dim MyDocument As Diadoc.Api.Proto.Documents.Document
        Dim MyContragent As Diadoc.Api.Proto.Organization
        Dim MyAfterIndexKey As String
        Dim MyDepartment As String

        Declarations.MyboxId = My.Settings.MyboxId
        MyDepartment = My.Settings.AcceptanceGroup

        MyDF = New DocumentsFilter
        MyDF.FilterCategory = "Invoice.InboundWaitingForResolution"
        MyDF.BoxId = Declarations.MyboxId

        '----получение списка документов на дату
        Try
            MyDocList = MyApi.GetDocuments(MyToken, MyDF)
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --5--> " & ex.Message)
            End If
        End Try

        Docdt.Rows.Clear()
        While MyDocList.Documents.Count <> 0
            For i As Integer = 0 To MyDocList.Documents.Count - 1
                MyDocument = MyDocList.Documents(i)
                MyContragent = MyApi.GetOrganizationByBoxId(MyDocument.CounteragentBoxId)
                MyAfterIndexKey = MyDocument.IndexKey

                If AutoEDOPresent(MyContragent.Inn, MyContragent.Kpp) = True _
                    And MyDocument.ResolutionStatus.Target.Department = MyDepartment Then
                    Docdt.Rows.Add()
                    '---название компании
                    Docdt.Rows(Docdt.Rows.Count - 1)("CompanyName") = MyContragent.FullName
                    '---ИНН
                    Docdt.Rows(Docdt.Rows.Count - 1)("INN") = MyContragent.Inn
                    '---КПП
                    Docdt.Rows(Docdt.Rows.Count - 1)("KPP") = MyContragent.Kpp
                    '---Номер документа
                    Docdt.Rows(Docdt.Rows.Count - 1)("DocNum") = MyDocument.DocumentNumber
                    '---Дата документа
                    Docdt.Rows(Docdt.Rows.Count - 1)("DocDate") = MyDocument.DocumentDate
                    '---Сумма документа
                    Docdt.Rows(Docdt.Rows.Count - 1)("DocSumm") = MyDocument.InvoiceMetadata.Total
                    '---НДС документа
                    Docdt.Rows(Docdt.Rows.Count - 1)("DocVAT") = MyDocument.InvoiceMetadata.Vat
                    '---Документ
                    Docdt.Rows(Docdt.Rows.Count - 1)("MyDoc") = MyDocument
                    '---Состояние документа
                    Docdt.Rows(Docdt.Rows.Count - 1)("DocState") = ""
                End If
            Next
            MyDF.AfterIndexKey = MyAfterIndexKey
            Try
                MyDocList = MyApi.GetDocuments(MyToken, MyDF)
            Catch ex As Exception
                If My.Settings.MyDebug = "YES" Then
                    EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --6--> " & ex.Message)
                End If
            End Try
        End While
        Session("Docdt") = Docdt

        SetNumberPages1()
    End Sub

    Private Sub SetNumberPages1()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Формирование списка страниц документов и выбор первого  
        '//
        '////////////////////////////////////////////////////////////////////////////////

        PagesList1.Items.Clear()
        For i As Integer = 0 To System.Math.Ceiling(Docdt.Rows.Count / QTYOnPageList1.SelectedValue) - 1
            PagesList1.Items.Insert(i, (i + 1).ToString())
        Next
        LabelQTYPages1.Text = "из " + System.Math.Ceiling(Docdt.Rows.Count / QTYOnPageList1.SelectedValue).ToString
    End Sub

    Private Function AutoEDOPresent(MyInn As String, MyKpp As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка поставщика с ИНН и КПП  
        '// работает ли он с автоматической приемкой товаров через ЭДО
        '//
        '////////////////////////////////////////////////////////////////////////////////

        For i As Integer = 0 To Suppdt.Rows.Count - 1
            If MyInn = Suppdt.Rows(i).Item(0) And MyKpp = Suppdt.Rows(i).Item(1) Then
                AutoEDOPresent = True
                Exit Function
            End If
        Next
        AutoEDOPresent = False
    End Function

    Private Sub ShowInvoiceList()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// вывод списка инвойсов в таблицу с разбиением по страницам 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim dt As DataTable                                   '--Datatable для полученных документов

        dt = New DataTable()
        dt.Columns.Add("CompanyName", Type.GetType("System.String"))
        dt.Columns.Add("INN", Type.GetType("System.String"))
        dt.Columns.Add("KPP", Type.GetType("System.String"))
        dt.Columns.Add("DocNum", Type.GetType("System.String"))
        dt.Columns.Add("DocDate", Type.GetType("System.String"))
        dt.Columns.Add("DocSumm", Type.GetType("System.String"))
        dt.Columns.Add("DocVAT", Type.GetType("System.String"))
        dt.Columns.Add("MyDoc", Type.GetType("System.Object"))
        dt.Columns.Add("DocState", Type.GetType("System.String"))
        For i As Integer = 0 To QTYOnPageList1.SelectedValue - 1
            dt.Rows.Add()
            Try
                For j As Integer = 0 To 8
                    dt.Rows(i).Item(j) = Docdt((PagesList1.SelectedValue - 1) * QTYOnPageList1.SelectedValue + i).Item(j)
                Next
            Catch ex As Exception
                dt.Rows(i).Delete()
                Exit For
            End Try
        Next
        GridView1.DataSource = dt
        GridView1.DataBind()
    End Sub

    Private Sub ButtonRun_Click(sender As Object, e As EventArgs) Handles ButtonRun.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка документов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        GetInvoiceList()
        ShowInvoiceList()
        GetPurchaseOrders()
        CheckButtonState()
    End Sub

    Private Sub SupplierList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SupplierList.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор поставщика (поставщиков) для загрузки документов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        GetSupplierInfo()
    End Sub

    Private Sub GetSupplierInfo()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение информации по выбраным поставщикам (поставщику) для загрузки документов
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                          '--рабочая строка
        Dim ds As New DataSet()

        Try
            Suppdt.Rows.Clear()
            If SupplierList.SelectedValue = "---" Then
                MySQLStr = "SELECT PL010300.PL01025 AS INN, PL010300.PL01048 AS KPP "
                MySQLStr = MySQLStr & "FROM PL010300 INNER JOIN "
                MySQLStr = MySQLStr & "tbl_SupplierCard0300 ON PL010300.PL01001 = tbl_SupplierCard0300.PL01001 "
                MySQLStr = MySQLStr & "WHERE (tbl_SupplierCard0300.EDO = 2) "
            Else
                MySQLStr = "SELECT PL01025 AS INN, PL01048 AS KPP "
                MySQLStr = MySQLStr & "FROM PL010300 "
                MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(SupplierList.SelectedValue.ToString) & "') "
            End If
            Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                Try
                    Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                        cmd.CommandType = CommandType.Text
                        Using da As New SqlDataAdapter()
                            da.SelectCommand = cmd
                            da.Fill(ds)
                            For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
                                Suppdt.Rows.Add()
                                Suppdt.Rows(Suppdt.Rows.Count - 1)("INN") = ds.Tables(0).Rows(i).Item(0)
                                Suppdt.Rows(Suppdt.Rows.Count - 1)("KPP") = ds.Tables(0).Rows(i).Item(1)
                            Next
                            Session("Suppdt") = Suppdt
                        End Using
                    End Using
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --7--> " & ex.Message)
                    End If
                Finally
                    MyConn.Close()
                End Try
            End Using
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --8--> " & ex.Message)
            End If
        End Try
    End Sub

    Private Sub PagesList1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PagesList1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена показываемой страницы списка инвойсов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ShowInvoiceList()
        Session("GridView1_SelRowOld") = Session("GridView1_SelRow")
        Session("GridView1_SelRow") = 0
        GridView1.SelectedIndex = Session("GridView1_SelRow")
        ChangeSelRow()
        GetPurchaseOrders()
        CheckButtonState()
    End Sub

    Private Sub QTYOnPageList1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles QTYOnPageList1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена количества инвойсов, выводимых на странице
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SetNumberPages1()
        ShowInvoiceList()
        Session("GridView1_SelRowOld") = Session("GridView1_SelRow")
        Session("GridView1_SelRow") = 0
        GridView1.SelectedIndex = Session("GridView1_SelRow")
        ChangeSelRow()
        GetPurchaseOrders()
        CheckButtonState()
    End Sub

    Private Sub GridView1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles GridView1.RowDataBound
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсвечивание строк
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.Row.RowType = DataControlRowType.DataRow Then

            e.Row.Attributes("onclick") = Page.ClientScript.GetPostBackClientHyperlink(GridView1, "Select$" & e.Row.RowIndex)
            e.Row.Attributes.Add("OnMouseOver", "this.style.cursor = 'pointer'")

            If e.Row.RowIndex = GridView1.SelectedIndex Then
                e.Row.BackColor = Drawing.Color.FromArgb(68, 68, 255)
            Else
                If Trim(e.Row.Cells(7).Text) = "&nbsp;" Then
                    e.Row.BackColor = Drawing.Color.White
                Else
                    e.Row.BackColor = Drawing.Color.LightGreen
                End If
            End If
        End If
    End Sub

    Private Sub GridView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GridView1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// изменение выделенной строки после того, как событие свершилось
        '//
        '////////////////////////////////////////////////////////////////////////////////

        GridView1.SelectedIndex = Session("GridView1_SelRow")
        ChangeSelRow()
        GetPurchaseOrders()
        CheckButtonState()
    End Sub

    Private Sub GridView1_SelectedIndexChanging(sender As Object, e As GridViewSelectEventArgs) Handles GridView1.SelectedIndexChanging
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// изменение выделенной строки до того, как событие свершилось
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Session("GridView1_SelRowOld") = Session("GridView1_SelRow")
        Session("GridView1_SelRow") = e.NewSelectedIndex
    End Sub

    Protected Sub ChangeSelRow()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// изменение подсветки выбранной строки
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Try
            If Session("GridView1_SelRow") <> Session("GridView1_SelRowOld") Then
                GridView1.Rows(Session("GridView1_SelRow")).BackColor = Drawing.Color.FromArgb(68, 68, 255)
                If Trim(GridView1.Rows(Session("GridView1_SelRowOld")).Cells(7).Text) = "&nbsp;" Then
                    GridView1.Rows(Session("GridView1_SelRowOld")).BackColor = Drawing.Color.White
                Else
                    GridView1.Rows(Session("GridView1_SelRowOld")).BackColor = Drawing.Color.LightGreen
                End If
            Else
                GridView1.Rows(Session("GridView1_SelRow")).BackColor = Drawing.Color.FromArgb(68, 68, 255)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ButtonQuit_Click(sender As Object, e As EventArgs) Handles ButtonQuit.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Session("Login") = ""
        Response.Redirect("Login.aspx", True)
    End Sub

    Private Sub CheckButtonState()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и выставление состояния кнопок
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If GridView1.Rows.Count = 0 Then
            ButtonGet.Enabled = False
            InvoiceView.Enabled = False
        Else
            If Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(7).Text) = "&nbsp;" Then
                ButtonGet.Enabled = True
            Else
                ButtonGet.Enabled = False
            End If
            InvoiceView.Enabled = True
        End If
    End Sub

    Private Sub CheckButtonState2()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и выставление состояния кнопок
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If GridView2.Rows.Count = 0 Then
            ButtonSelectAll.Enabled = False
            ButtonDeSelectAll.Enabled = False
        Else
            ButtonSelectAll.Enabled = True
            ButtonDeSelectAll.Enabled = True
        End If
    End Sub

    Private Function CheckSelectedOrders() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка наличия и выбранности заказов на закупку для приемки товаров
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim cb As CheckBox

        For i As Integer = 0 To GridView2.Rows.Count - 1
            cb = GridView2.Rows(i).Cells(0).FindControl("POCheckBox")
            If cb.Checked = True Then
                CheckSelectedOrders = True
                Exit Function
            End If
        Next
        CheckSelectedOrders = False
    End Function

    Private Sub GetPurchaseOrders()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// вызов процедуры получения списка незакрытых заказов на закупку поставщика, приславшего выбранный документ
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If GridView1.Rows.Count = 0 Then
        Else
            LoadPurchaseOrders(Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(1).Text), Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(2).Text))
            MarkRequiredPurchaseOrders()
        End If
        CheckButtonState2()
    End Sub

    Private Sub LoadPurchaseOrders(MyINN As String, MyKPP As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура получения списка незакрытых заказов на закупку поставщика, приславшего выбранный документ
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim ds As New DataSet()

        Try
            MySQLStr = "dbo.spp_EDO_GetPurchaseOrders"
            Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                Try
                    Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.AddWithValue("@MyINN", MyINN)
                        cmd.Parameters.AddWithValue("@MyKPP", MyKPP)
                        cmd.Parameters.AddWithValue("@MyWH", WHList.SelectedValue)
                        Using da As New SqlDataAdapter()
                            da.SelectCommand = cmd
                            da.Fill(ds)
                            GridView2.DataSource = ds.Tables(0)
                            GridView2.DataBind()
                        End Using
                    End Using
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --9--> " & ex.Message)
                    End If
                Finally
                    MyConn.Close()
                End Try
            End Using
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --10--> " & ex.Message)
            End If
        End Try
    End Sub

    Private Sub MarkRequiredPurchaseOrders()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура маркировки заказов на закупку - помечаются те, в которых есть товары поставщика, присутствующие в инвойсе.
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyInvoice As Diadoc.Api.Proto.Documents.Document
        Dim MI As Byte()
        Dim MyGuid As String = ""
        Dim MySQLStr As String
        Dim ds As New DataSet()
        Dim cb As CheckBox

        MyGuid = Guid.NewGuid.ToString("N")

        Try
            MyInvoice = GetDocForView(Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(1).Text), Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(2).Text), _
                    Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(3).Text), Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(4).Text))
            For i As Integer = 0 To 50
                MI = MyApi.GetEntityContent(MyToken, Declarations.MyboxId, MyInvoice.MessageId, MyInvoice.EntityId)
                If MI.Length = 0 Then
                    System.Threading.Thread.Sleep(1 * 1000)
                Else
                    Exit For
                End If
            Next

            If MI.Length <> 0 Then  '--------------------документ есть
                If LoadInvoiceToTmpTables(MI, MyGuid) = True Then
                    MySQLStr = "SELECT PC010300.PC01052 AS ConsolidatedOrderNum "
                    MySQLStr = MySQLStr & "FROM PC010300 INNER JOIN "
                    MySQLStr = MySQLStr & "PC030300 ON PC010300.PC01001 = PC030300.PC03001 INNER JOIN "
                    MySQLStr = MySQLStr & "PL010300 ON PC010300.PC01003 = PL010300.PL01001 INNER JOIN "
                    MySQLStr = MySQLStr & "SC010300 ON PC030300.PC03005 = SC010300.SC01001 INNER JOIN "
                    MySQLStr = MySQLStr & "MyEDOInvoice_" & MyGuid & " ON SC010300.SC01060 = MyEDOInvoice_" & MyGuid & ".SupplierItemCode "
                    MySQLStr = MySQLStr & "WHERE (PL010300.PL01025 = N'" & Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(1).Text) & "') "
                    MySQLStr = MySQLStr & "AND (PL010300.PL01048 = N'" & Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(2).Text) & "') "
                    MySQLStr = MySQLStr & "AND (PC010300.PC01002 <> 0) "
                    MySQLStr = MySQLStr & "AND (PC030300.PC03010 <> 0)"
                    MySQLStr = MySQLStr & " AND (PC010300.PC01052 <> '') "
                    MySQLStr = MySQLStr & "AND (PC010300.PC01023 = N'" & WHList.SelectedValue & "') "
                    MySQLStr = MySQLStr & "GROUP BY PC010300.PC01052 "
                    MySQLStr = MySQLStr & "ORDER BY ConsolidatedOrderNum "
                    Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                        Try
                            Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                                cmd.CommandType = CommandType.Text
                                Using da As New SqlDataAdapter()
                                    da.SelectCommand = cmd
                                    da.Fill(ds)
                                    If ds.Tables(0).Rows.Count <> 0 Then
                                        For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
                                            For j As Integer = 0 To GridView2.Rows.Count - 1
                                                If Trim(ds.Tables(0).Rows(i).Item(0).ToString) = Trim(GridView2.Rows(j).Cells(1).Text) Then
                                                    cb = GridView2.Rows(j).Cells(0).FindControl("POCheckBox")
                                                    cb.Checked = True
                                                End If
                                            Next
                                        Next
                                    End If
                                End Using
                            End Using
                        Catch ex As Exception
                            If My.Settings.MyDebug = "YES" Then
                                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --11--> " & ex.Message)
                            End If
                        Finally
                            MyConn.Close()
                        End Try
                    End Using
                End If
                DeleteTmpTables(MyGuid)
            End If
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --12--> " & ex.Message)
            End If
        End Try
    End Sub

    Private Sub GridView2_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GridView2.RowCommand
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вызов процедуры просмотра заказа на закупку
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim ds As New DataSet()

        If e.CommandName = "PurchaseOrderView" Then
            POrderNum = e.CommandArgument.ToString()
            Try
                MySQLStr = "dbo.spp_EDO_GetPurchaseOrderDetails"
                Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                    Try
                        Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                            cmd.CommandType = CommandType.StoredProcedure
                            cmd.Parameters.AddWithValue("@MyPONum", Trim(POrderNum))
                            Using da As New SqlDataAdapter()
                                da.SelectCommand = cmd
                                da.Fill(ds)
                                GridView3.DataSource = ds.Tables(0)
                                GridView3.DataBind()
                            End Using
                        End Using
                    Catch ex As Exception
                        If My.Settings.MyDebug = "YES" Then
                            EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --13--> " & ex.Message)
                        End If
                    Finally
                        MyConn.Close()
                    End Try
                End Using
            Catch ex As Exception
                If My.Settings.MyDebug = "YES" Then
                    EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --14--> " & ex.Message)
                End If
            End Try
            POView.Disabled = False
            POView.Visible = True
            DivBG.Disabled = False
            DivBG.Visible = True
        End If
    End Sub

    Private Sub InvoiceView_Click(sender As Object, e As EventArgs) Handles InvoiceView.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// просмотр инвойса
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyIframe As StringBuilder

        PDFView.Disabled = False
        PDFView.Visible = True
        DivBG.Disabled = False
        DivBG.Visible = True

        Session("InvoiceToPDFDoc") = GetDocForView(Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(1).Text), Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(2).Text), _
            Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(3).Text), Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(4).Text))
        MyIframe = New StringBuilder
        MyIframe.Append("<iframe id=""iFrame1"" scrolling=""yes"" width=""100%"" height=""100%"" src=""GetInvoicePDF.aspx"" />")
        PDFLiteral.Text = MyIframe.ToString
    End Sub

    Private Sub ButtonPDFLiteralClose_Click(sender As Object, e As EventArgs) Handles ButtonPDFLiteralClose.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие фрейма просмотра инвойса
        '//
        '////////////////////////////////////////////////////////////////////////////////

        PDFView.Disabled = True
        PDFView.Visible = False
        DivBG.Disabled = True
        DivBG.Visible = False
    End Sub

    Private Function GetDocForView(MyINN As String, MyKPP As String, MyDocNum As String, MyDocdate As String) As Diadoc.Api.Proto.Documents.Document
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение объекта типа документ для выбранной строки
        '//
        '////////////////////////////////////////////////////////////////////////////////

        For i As Integer = 0 To Docdt.Rows.Count - 1
            If Docdt.Rows(i).Item(1) = MyINN And Docdt.Rows(i).Item(2) = MyKPP _
                And Docdt.Rows(i).Item(3) = MyDocNum And Docdt.Rows(i).Item(4) = MyDocdate Then
                GetDocForView = Docdt.Rows(i).Item(7)
                Exit Function
            End If
        Next
        GetDocForView = Nothing
    End Function

    Private Sub ButtonSelectAll_Click(sender As Object, e As EventArgs) Handles ButtonSelectAll.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// выбор всех заказов на закупку
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim cb As CheckBox

        For i As Integer = 0 To GridView2.Rows.Count - 1
            cb = GridView2.Rows(i).Cells(0).FindControl("POCheckBox")
            cb.Checked = True
        Next
    End Sub

    Private Sub ButtonDeSelectAll_Click(sender As Object, e As EventArgs) Handles ButtonDeSelectAll.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// снятие выбора всех заказов на закупку
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim cb As CheckBox

        For i As Integer = 0 To GridView2.Rows.Count - 1
            cb = GridView2.Rows(i).Cells(0).FindControl("POCheckBox")
            cb.Checked = False
        Next
    End Sub

    Private Sub ButtonPOViewClose_Click(sender As Object, e As EventArgs) Handles ButtonPOViewClose.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие фрейма просмотра заказа на закупку
        '//
        '////////////////////////////////////////////////////////////////////////////////

        POView.Disabled = True
        POView.Visible = False
        DivBG.Disabled = True
        DivBG.Visible = False
    End Sub

    Private Sub ButtonGet_Click(sender As Object, e As EventArgs) Handles ButtonGet.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Приемка счет фактуры
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyInvoice As Diadoc.Api.Proto.Documents.Document
        Dim MI As Byte()
        Dim MyGuid As String = ""

        MyGuid = Guid.NewGuid.ToString("N")

        If CheckSelectedOrders() = True Then
            Try
                MyInvoice = GetDocForView(Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(1).Text), Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(2).Text), _
                    Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(3).Text), Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(4).Text))
                For i As Integer = 0 To 50
                    MI = MyApi.GetEntityContent(MyToken, Declarations.MyboxId, MyInvoice.MessageId, MyInvoice.EntityId)
                    If MI.length = 0 Then
                        System.Threading.Thread.Sleep(1 * 1000)
                    Else
                        Exit For
                    End If
                Next

                If MI.length <> 0 Then  '--------------------документ есть - принимаем
                    If LoadInvoiceToTmpTables(MI, MyGuid) = True Then
                        MyErrString = AcceptInvoice(MyGuid)
                        If MyErrString = "" Then
                            '------------пометка о выполнении 
                            MyErrString = MarkInvoiceAsAccepted()
                        End If
                        If MyErrString = "" Then
                            '------------Вывод информации о результатах приемки
                            MyErrString = GetFinaldata(MyGuid)
                            TxtErr.Text = "Приемка инвойса завершена " + Chr(13) + MyErrString
                            TxtErr.Style.Add("text-align", "Left")
                            DivErr.Disabled = False
                            DivErr.Visible = True
                            DivBG.Disabled = False
                            DivBG.Visible = True
                        Else
                            MyErrString = "Приемка инвойса завершена с ошибками " + Chr(13) + MyErrString + Chr(13) + GetFinaldata(MyGuid)
                            TxtErr.Text = MyErrString
                            TxtErr.Style.Add("text-align", "Left")
                            DivErr.Disabled = False
                            DivErr.Visible = True
                            DivBG.Disabled = False
                            DivBG.Visible = True
                        End If
                    Else
                        TxtErr.Text = MyErrString
                        TxtErr.Style.Add("text-align", "center")
                        DivErr.Disabled = False
                        DivErr.Visible = True
                        DivBG.Disabled = False
                        DivBG.Visible = True
                    End If
                '---------удаление таблиц
                DeleteTmpTables(MyGuid)
                End If

            Catch ex As Exception
                If My.Settings.MyDebug = "YES" Then
                    EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --15--> " & ex.Message)
                End If
            End Try
        Else
            TxtErr.Text = "Необходимо выбрать хотя бы один заказ на закупку."
            TxtErr.Style.Add("text-align", "center")
            DivErr.Disabled = False
            DivErr.Visible = True
            DivBG.Disabled = False
            DivBG.Visible = True
        End If
    End Sub

    Private Function MarkInvoiceAsAccepted() As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Пометка СФ как принятой
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyErrStr As String
        Dim MyInvoice As Diadoc.Api.Proto.Documents.Document
        Dim MyPM As Diadoc.Api.Proto.Events.MessagePatchToPost
        Dim MyResAtt As Diadoc.Api.Proto.Events.ResolutionAttachment

        MyErrStr = ""
        '---------------------в клиенте
        For i As Integer = 0 To Docdt.Rows.Count - 1
            If Docdt.Rows(i).Item(1) = Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(1).Text) And Docdt.Rows(i).Item(2) = Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(2).Text) _
                And Docdt.Rows(i).Item(3) = Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(3).Text) And Docdt.Rows(i).Item(4) = Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(4).Text) Then
                Docdt.Rows(i).Item(8) = "Принято"
                Exit For
            End If
        Next

        '---------------------В Диадоке
        Try
            MyInvoice = GetDocForView(Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(1).Text), Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(2).Text), _
                    Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(3).Text), Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(4).Text))
            MyPM = New Diadoc.Api.Proto.Events.MessagePatchToPost
            MyResAtt = New Diadoc.Api.Proto.Events.ResolutionAttachment

            MyResAtt.InitialDocumentId = MyInvoice.EntityId
            MyResAtt.ResolutionType = Proto.Events.ResolutionType.Approve

            MyPM.AddResolution(MyResAtt)
            MyPM.BoxId = Declarations.MyboxId
            MyPM.MessageId = MyInvoice.MessageId

            MyApi.PostMessagePatch(MyToken, MyPM)
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --35--> " & ex.Message)
            End If
            MyErrStr = MyErrStr + ex.Message + Chr(13)
        End Try

        ShowInvoiceList()
        ChangeSelRow()
        GetPurchaseOrders()
        CheckButtonState()
        MarkInvoiceAsAccepted = MyErrStr
    End Function

    Private Sub ButtonErr_Click(sender As Object, e As EventArgs) Handles ButtonErr.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна с сообщением об ошибке
        '//
        '////////////////////////////////////////////////////////////////////////////////

        DivErr.Disabled = True
        DivErr.Visible = False
        DivBG.Disabled = True
        DivBG.Visible = False
    End Sub

    Private Function LoadInvoiceToTmpTables(ByRef MI As Byte(), MyGuid As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации об инвойсе во временные таблички
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim aa As New System.Globalization.NumberFormatInfo

        MyErrString = ""


        '--------------Создание временных таблиц------------------
        '--------временная таблица MyEDOInvoice
        Try
            MySQLStr = "CREATE TABLE MyEDOInvoice_" & MyGuid & "( "
            MySQLStr = MySQLStr & "[ID] int, "                                      '--ID строки
            MySQLStr = MySQLStr & "[Invoice] [nvarchar](35), "                      '--номер СФ
            MySQLStr = MySQLStr & "[InvoiceDate] [datetime], "                      '--дата СФ
            MySQLStr = MySQLStr & "[InvoiceCurrCode] int, "                         '--Код валюты СФ
            MySQLStr = MySQLStr & "[SalesmanCode] [nvarchar](3), "                  '--код продавца
            MySQLStr = MySQLStr & "[SalesmanName] [nvarchar](25), "                 '--имя продавца
            MySQLStr = MySQLStr & "[InvoiceCurrExchRate] float, "                   '--Курс валюты в инвойсе
            MySQLStr = MySQLStr & "[ConsPurchaseOrderNum] [nvarchar](10) NULL, "    '--Номер консолидированного заказа на закупку
            MySQLStr = MySQLStr & "[SupplierItemCode] [nvarchar](35), "             '--код товара поставщика
            MySQLStr = MySQLStr & "[QTY] float, "                                   '--количество
            MySQLStr = MySQLStr & "[SummWithoutVAT] float, "                        '--Сумма без НДС за строку
            MySQLStr = MySQLStr & "[Country] nvarchar(50), "                        '-- страна производителя
            MySQLStr = MySQLStr & "[GTD] nvarchar (255), "                          '-- ГТД
            MySQLStr = MySQLStr & "[RestQTY] float  "                               '--Остаток - непринятое количество
            MySQLStr = MySQLStr & ") "
            Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                Try
                    MyConn.Open()
                    Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                        cmd.CommandType = CommandType.Text
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --16--> " & ex.Message)
                    End If
                    MyErrString = ex.Message
                    LoadInvoiceToTmpTables = False
                    Exit Function
                Finally
                    MyConn.Close()
                End Try
            End Using
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --17--> " & ex.Message)
            End If
            MyErrString = ex.Message
            LoadInvoiceToTmpTables = False
            Exit Function
        End Try

        '--------временная таблица MyEDOOrder
        Try
            MySQLStr = "CREATE TABLE MyEDOOrder_" & MyGuid & "( "
            MySQLStr = MySQLStr & "[ConsPurchaseOrderNum] [nvarchar](10) "    '--Номер консолидированного заказа на закупку
            MySQLStr = MySQLStr & ") "
            Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                Try
                    MyConn.Open()
                    Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                        cmd.CommandType = CommandType.Text
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --18--> " & ex.Message)
                    End If
                    MyErrString = ex.Message
                    LoadInvoiceToTmpTables = False
                    Exit Function
                Finally
                    MyConn.Close()
                End Try
            End Using
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --19--> " & ex.Message)
            End If
            MyErrString = ex.Message
            LoadInvoiceToTmpTables = False
            Exit Function
        End Try

        '-----------------------------------Загрузка данных во временные таблички----------------
        '--------------Обобщенные заказы на закупку
        Dim cb As CheckBox

        For i As Integer = 0 To GridView2.Rows.Count - 1
            cb = GridView2.Rows(i).Cells(0).FindControl("POCheckBox")
            If cb.Checked = True Then
                Try
                    MySQLStr = "DELETE FROM MyEDOOrder_" & MyGuid & " "
                    MySQLStr = MySQLStr & "WHERE (ConsPurchaseOrderNum = N'" & Trim(GridView2.Rows(i).Cells(1).Text) & "') "
                    Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                        Try
                            MyConn.Open()
                            Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                                cmd.CommandType = CommandType.Text
                                cmd.ExecuteNonQuery()
                            End Using
                        Catch ex As Exception
                            If My.Settings.MyDebug = "YES" Then
                                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --20--> " & ex.Message)
                            End If
                            MyErrString = ex.Message
                            LoadInvoiceToTmpTables = False
                            Exit Function
                        Finally
                            MyConn.Close()
                        End Try
                    End Using
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --21--> " & ex.Message)
                    End If
                    MyErrString = ex.Message
                    LoadInvoiceToTmpTables = False
                    Exit Function
                End Try
                Try
                    MySQLStr = "INSERT INTO MyEDOOrder_" & MyGuid & " "
                    MySQLStr = MySQLStr & "Select N'" & Trim(GridView2.Rows(i).Cells(1).Text) & "' AS PO "
                    Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                        Try
                            MyConn.Open()
                            Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                                cmd.CommandType = CommandType.Text
                                cmd.ExecuteNonQuery()
                            End Using
                        Catch ex As Exception
                            If My.Settings.MyDebug = "YES" Then
                                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --22--> " & ex.Message)
                            End If
                            MyErrString = ex.Message
                            LoadInvoiceToTmpTables = False
                            Exit Function
                        Finally
                            MyConn.Close()
                        End Try
                    End Using
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --23--> " & ex.Message)
                    End If
                    MyErrString = ex.Message
                    LoadInvoiceToTmpTables = False
                    Exit Function
                End Try
            End If
        Next


        '--------------принимаемая СФ
        Dim MyDoc As XmlDocument
        Dim MyHeaderNode As XmlNode
        Dim MyFirstItemNode As XmlNode
        Dim MyItemNodeList As XmlNodeList
        Dim MyInvoice As String                         '--номер СФ
        Dim MyInvoiceDate As String                     '--дата СФ
        Dim MyInvoiceCurrCode As Integer                '--Код валюты СФ
        Dim MySalesmanCode As String                    '--код продавца
        Dim MySalesmanName As String                    '--имя продавца
        Dim MyInvoiceCurrExchRate As Double             '--Курс валюты в инвойсе
        Dim MySupplierItemCode As String                '--код товара поставщика
        Dim MyQTY As Double                             '--количество
        Dim MySummWithoutVAT As Double                  '--Сумма без НДС за строку
        Dim MyCountry As String                         '-- Код страны производителя
        Dim MyGTD As String                             '-- ГТД

        Dim Myxml As String = Encoding.GetEncoding(1251).GetString(MI)
        Try
            MyDoc = New XmlDocument
            MyDoc.LoadXml(Myxml)
            MyHeaderNode = MyDoc.DocumentElement.ChildNodes(1)
            MyFirstItemNode = MyHeaderNode.ChildNodes(1)
            MyInvoice = MyHeaderNode.ChildNodes(0).Attributes("НомерСчФ").Value
            MyInvoiceDate = Replace(MyHeaderNode.ChildNodes(0).Attributes("ДатаСчФ").Value, ".", "/")
            MyInvoiceCurrCode = "0" '--так как СФ всегда в рублях
            MySalesmanCode = Session("SalesmanCode")
            MySalesmanName = Session("SalesmanName")

            MyItemNodeList = MyFirstItemNode.ChildNodes
            For i As Integer = 0 To MyItemNodeList.Count - 2
                MySupplierItemCode = GetCodeNumberFromName(Trim(MyItemNodeList(i).Attributes("НаимТов").Value))

                If aa.CurrentInfo.NumberDecimalSeparator = "," Then
                    MyQTY = CDbl(Replace(MyItemNodeList(i).Attributes("КолТов").Value.ToString, ".", ","))
                    MySummWithoutVAT = CDbl(Replace(MyItemNodeList(i).Attributes("СтТовБезНДС").Value.ToString, ".", ","))
                Else
                    MyQTY = MyItemNodeList(i).Attributes("КолТов").Value
                    MySummWithoutVAT = MyItemNodeList(i).Attributes("СтТовБезНДС").Value
                End If

                MyCountry = MyItemNodeList(i).ChildNodes(2).Attributes("КодПроисх").Value
                MyGTD = MyItemNodeList(i).ChildNodes(2).Attributes("НомерТД").Value

                MySQLStr = "INSERT INTO MyEDOInvoice_" & MyGuid & " "
                MySQLStr = MySQLStr & "(ID, Invoice, InvoiceDate, InvoiceCurrCode, SalesmanCode, SalesmanName, InvoiceCurrExchRate, "
                MySQLStr = MySQLStr & "ConsPurchaseOrderNum, SupplierItemCode, QTY, SummWithoutVAT, Country, GTD, RestQTY) "
                MySQLStr = MySQLStr & "VALUES (" & CStr(i + 1) & ", "
                MySQLStr = MySQLStr & "N'" & MyInvoice & "', "
                MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & MyInvoiceDate & "', 103), "
                MySQLStr = MySQLStr & CStr(MyInvoiceCurrCode) & ", "
                MySQLStr = MySQLStr & "N'" & MySalesmanCode & "', "
                MySQLStr = MySQLStr & "N'" & MySalesmanName & "', "
                MySQLStr = MySQLStr & Replace(CStr(MyInvoiceCurrExchRate), ",", ".") & ", "
                MySQLStr = MySQLStr & "NULL, "
                MySQLStr = MySQLStr & "N'" & MySupplierItemCode & "', "
                MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ", "
                MySQLStr = MySQLStr & Replace(CStr(MySummWithoutVAT), ",", ".") & ", "
                MySQLStr = MySQLStr & "N'" & MyCountry & "', "
                MySQLStr = MySQLStr & "N'" & MyGTD & "', "
                MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ") "
                Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                    Try
                        MyConn.Open()
                        Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                            cmd.CommandType = CommandType.Text
                            cmd.ExecuteNonQuery()
                        End Using
                    Catch ex As Exception
                        If My.Settings.MyDebug = "YES" Then
                            EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --24--> " & ex.Message)
                        End If
                        MyErrString = ex.Message
                        LoadInvoiceToTmpTables = False
                        Exit Function
                    Finally
                        MyConn.Close()
                    End Try
                End Using
            Next
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --25--> " & ex.Message)
            End If
            MyErrString = ex.Message
            LoadInvoiceToTmpTables = False
            Exit Function
        End Try
        LoadInvoiceToTmpTables = True
    End Function

    Private Function AcceptInvoice(MyGuid As String) As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации из временных табличек в Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRezStr As String
        Dim MySQLStr As String
        Dim ds As New DataSet()
        Dim MyRestQty As Double

        MyRezStr = ""
        Try
            MySQLStr = "SELECT ID, Invoice, InvoiceDate, InvoiceCurrCode, SalesmanCode, "
            MySQLStr = MySQLStr & "SalesmanName, InvoiceCurrExchRate, View_10.ConsolidatedOrderNum, SupplierItemCode, QTY, "
            MySQLStr = MySQLStr & "SummWithoutVAT, Country, GTD, RestQTY "
            MySQLStr = MySQLStr & "FROM  MyEDOInvoice_" & MyGuid & " INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT TOP (100) PERCENT PC010300.PC01052 AS ConsolidatedOrderNum, SC010300.SC01060 "
            MySQLStr = MySQLStr & "FROM  PC010300 INNER JOIN "
            MySQLStr = MySQLStr & "PC030300 ON PC010300.PC01001 = PC030300.PC03001 INNER JOIN "
            MySQLStr = MySQLStr & "PL010300 ON PC010300.PC01003 = PL010300.PL01001 INNER JOIN "
            MySQLStr = MySQLStr & "SC010300 ON PC030300.PC03005 = SC010300.SC01001 "
            MySQLStr = MySQLStr & "WHERE (PL010300.PL01025 = N'" & Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(1).Text) & "') "
            MySQLStr = MySQLStr & "AND (PL010300.PL01048 = N'" & Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(2).Text) & "') "
            MySQLStr = MySQLStr & "AND (PC010300.PC01002 <> 0) "
            MySQLStr = MySQLStr & "AND (PC010300.PC01052 <> '') "
            MySQLStr = MySQLStr & "AND (PC030300.PC03010 <> 0) "
            MySQLStr = MySQLStr & "GROUP BY PC010300.PC01052, SC010300.SC01060) AS View_10 ON MyEDOInvoice_" & MyGuid & ".SupplierItemCode = View_10.SC01060 "
            MySQLStr = MySQLStr & "ORDER BY SupplierItemCode, QTY DESC, View_10.ConsolidatedOrderNum "
            Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                Try
                    Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                        MyConn.Open()
                        cmd.CommandType = CommandType.Text
                        Using da As New SqlDataAdapter()
                            da.SelectCommand = cmd
                            da.Fill(ds)
                            '---в цикле вызываем процедуру построчной приемки инвойса
                            For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
                                MySQLStr = "spp_EDO_PurchaseInvoice_AutoUploadLine"
                                Using cmd1 As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                                    cmd1.CommandType = CommandType.StoredProcedure
                                    cmd1.Parameters.AddWithValue("@ID", ds.Tables(0).Rows(i).Item(0))
                                    cmd1.Parameters.AddWithValue("@Invoice", ds.Tables(0).Rows(i).Item(1))
                                    cmd1.Parameters.AddWithValue("@InvoiceDateSTR", ds.Tables(0).Rows(i).Item(2))
                                    cmd1.Parameters.AddWithValue("@CurrCode", ds.Tables(0).Rows(i).Item(3))
                                    cmd1.Parameters.AddWithValue("@MySalesmanCode", ds.Tables(0).Rows(i).Item(4))
                                    cmd1.Parameters.AddWithValue("@MySalesmanName", ds.Tables(0).Rows(i).Item(5))
                                    cmd1.Parameters.AddWithValue("@PurchInvoiceExRate", ds.Tables(0).Rows(i).Item(6))
                                    cmd1.Parameters.AddWithValue("@ConsPOrder", ds.Tables(0).Rows(i).Item(7))
                                    cmd1.Parameters.AddWithValue("@ItemSuppCode", ds.Tables(0).Rows(i).Item(8))
                                    cmd1.Parameters.AddWithValue("@QTY", ds.Tables(0).Rows(i).Item(13))
                                    cmd1.Parameters.AddWithValue("@Price", ds.Tables(0).Rows(i).Item(10))
                                    cmd1.Parameters.AddWithValue("@Country", ds.Tables(0).Rows(i).Item(11))
                                    cmd1.Parameters.AddWithValue("@GTD", ds.Tables(0).Rows(i).Item(12))
                                    cmd1.Parameters.AddWithValue("@MyGuid", MyGuid)
                                    cmd1.Parameters.Add("@MyRezStr", SqlDbType.VarChar, 4000)
                                    cmd1.Parameters("@MyRezStr").Direction = ParameterDirection.Output
                                    cmd1.Parameters.Add("@MyRestQTY", SqlDbType.Decimal)
                                    cmd1.Parameters("@MyRestQTY").Direction = ParameterDirection.Output
                                    cmd1.ExecuteNonQuery()
                                    MyRezStr = MyRezStr + cmd1.Parameters("@MyRezStr").Value.ToString()
                                    MyRestQty = cmd1.Parameters("@MyRestQTY").Value

                                    '--------------корректировка остатков
                                    For j As Integer = 0 To ds.Tables(0).Rows.Count - 1
                                        If ds.Tables(0).Rows(i).Item(0) = ds.Tables(0).Rows(j).Item(0) Then
                                            ds.Tables(0).Rows(j).Item(13) = MyRestQty
                                        End If
                                    Next
                                    MySQLStr = "UPDATE MyEDOInvoice_" & MyGuid & " "
                                    MySQLStr = MySQLStr & "SET RestQTY = " & Replace(CStr(MyRestQty), ".", ",") & " "
                                    MySQLStr = MySQLStr & "WHERE (ID = " & CStr(ds.Tables(0).Rows(i).Item(0)) & ") "
                                    Using cmd2 As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                                        cmd2.CommandType = CommandType.Text
                                        cmd2.ExecuteNonQuery()
                                    End Using
                                End Using
                            Next
                        End Using
                    End Using
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --26--> " & ex.Message)
                    End If
                    MyRezStr = MyRezStr + " " + ex.Message
                    AcceptInvoice = MyRezStr
                    Exit Function
                Finally
                    MyConn.Close()
                End Try
            End Using
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --27--> " & ex.Message)
            End If
            MyRezStr = MyRezStr + " " + ex.Message
        End Try

        AcceptInvoice = MyRezStr
    End Function

    Private Function GetCodeNumberFromName(MyName As String) As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение кода товара из названия (предполагаем, что строка состоит из кода товара + пробел + название товара)
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyCodeLength As Integer

        MyCodeLength = InStr(MyName, " ")
        If MyCodeLength <= 0 Then
            GetCodeNumberFromName = ""
        Else
            GetCodeNumberFromName = Mid(MyName, 1, MyCodeLength)
        End If
    End Function

    Private Sub DeleteTmpTables(MyGuid As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление временных таблиц
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '--------временная таблица #_MyEDOInvoice
        Try
            MySQLStr = "DROP TABLE MyEDOInvoice_" & MyGuid & " "
            Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                Try
                    MyConn.Open()
                    Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                        cmd.CommandType = CommandType.Text
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --28--> " & ex.Message)
                    End If
                Finally
                    MyConn.Close()
                End Try
            End Using
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --29--> " & ex.Message)
            End If
        End Try

        '--------временная таблица #_MyEDOOrder
        Try
            MySQLStr = "DROP TABLE MyEDOOrder_" & MyGuid & " "
            Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                Try
                    MyConn.Open()
                    Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                        cmd.CommandType = CommandType.Text
                        cmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --30--> " & ex.Message)
                    End If
                Finally
                    MyConn.Close()
                End Try
            End Using
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --31--> " & ex.Message)
            End If
        End Try
    End Sub

    Private Function GetFinaldata(MyGuid As String) As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение итоговой информации по результатам приемки инвойса
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRezStr As String
        Dim MySQLStr As String
        Dim ds As New DataSet()
        Dim ds1 As New DataSet()

        MyRezStr = ""
        '------------Вывод информации о заказах на закупку, по которым была приемка
        MyRezStr = MyRezStr + "----------------Информация о заказах на закупку, по которым была приемка---------" + Chr(13)
        Try
            MySQLStr = "SELECT PC190300.PC19001 AS OrderNum, "
            MySQLStr = MySQLStr & "PC010300.PC01023 AS WhNum "
            MySQLStr = MySQLStr & "FROM PC190300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "PC010300 ON PC190300.PC19001 = PC010300.PC01001 "
            MySQLStr = MySQLStr & "WHERE (PC190300.PC19012 = N'" & Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(3).Text) & "') AND "
            MySQLStr = MySQLStr & "(PC190300.PC19010 = dateadd( day, datediff(day, 0, GETDATE()), 0)) "
            MySQLStr = MySQLStr & "GROUP BY PC190300.PC19001, "
            MySQLStr = MySQLStr & "PC010300.PC01023 "
            MySQLStr = MySQLStr & "ORDER BY OrderNum "
            Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                Try
                    Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                        MyConn.Open()
                        cmd.CommandType = CommandType.Text
                        Using da As New SqlDataAdapter()
                            da.SelectCommand = cmd
                            da.Fill(ds)
                            If ds.Tables(0).Rows.Count = 0 Then
                                MyRezStr = MyRezStr & "Импорт СФ N " & Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(3).Text) & " не произведен ни для одного заказа на закупку " & Chr(13)
                            Else
                                MyRezStr = MyRezStr & "Импорт СФ N " & Trim(GridView1.Rows(Session("GridView1_SelRow")).Cells(3).Text) & " произведен для следующих заказов на закупку: " & Chr(13)
                                MyRezStr = MyRezStr & "Заказ на закупку    Номер склада" & Chr(13)
                                For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
                                    MyRezStr = MyRezStr & Left(ds.Tables(0).Rows(i).Item(0) + "                      ", 22) & ds.Tables(0).Rows(i).Item(1) & Chr(13)
                                Next
                            End If
                        End Using
                    End Using
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --32--> " & ex.Message)
                    End If
                    MyRezStr = MyRezStr + " " + ex.Message + Chr(13)
                Finally
                    MyConn.Close()
                End Try
            End Using
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --33--> " & ex.Message)
            End If
            MyRezStr = MyRezStr + " " + ex.Message + Chr(13)
        End Try
        MyRezStr = MyRezStr + Chr(13)

        '------------Вывод информации о перепоставках
        MySQLStr = "SELECT SupplierItemCode, QTY, RestQTY "
        MySQLStr = MySQLStr & "FROM MyEDOInvoice_" & MyGuid & " WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (RestQTY <> 0) "
        Using MyConn1 As SqlConnection = New SqlConnection(Declarations.ConnString)
            Try
                Using cmd1 As SqlCommand = New SqlCommand(MySQLStr, MyConn1)
                    MyConn1.Open()
                    cmd1.CommandType = CommandType.Text
                    Using da As New SqlDataAdapter()
                        da.SelectCommand = cmd1
                        da.Fill(ds1)
                        If ds1.Tables(0).Rows.Count = 0 Then
                        Else
                            MyRezStr = MyRezStr + "----------------Информация о перепоставках---------------------------------" + Chr(13)
                            MyRezStr = MyRezStr & "Код товара поставщика               Количество в СФ     Непринятое количество" & Chr(13)
                            For i As Integer = 0 To ds1.Tables(0).Rows.Count - 1
                                MyRezStr = MyRezStr & Left(ds1.Tables(0).Rows(i).Item(0).ToString + "                                    ", 36) & Left(ds1.Tables(0).Rows(i).Item(1).ToString + "                    ", 20) & ds1.Tables(0).Rows(i).Item(2).ToString & Chr(13)
                            Next
                        End If
                    End Using
                End Using
            Catch ex As Exception
                If My.Settings.MyDebug = "YES" Then
                    EventLog.WriteEntry("AcceptanceFromEDO", "IncomingInvoicesList --34--> " & ex.Message)
                End If
                MyRezStr = MyRezStr + " " + ex.Message + Chr(13)
            End Try
        End Using
        GetFinaldata = MyRezStr
    End Function

    Private Sub WHList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles WHList.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор склада
        '//
        '////////////////////////////////////////////////////////////////////////////////

        GetPurchaseOrders()
    End Sub

    Private Sub GridView2_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles GridView2.RowDataBound
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка строк заказов на закупку
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Attributes.Add("OnMouseOver", "this.originalstyle=this.style.backgroundColor;this.style.backgroundColor='#EFEFFF'")
            e.Row.Attributes.Add("OnMouseOut", "this.style.backgroundColor=this.originalstyle;")
        End If
    End Sub
End Class