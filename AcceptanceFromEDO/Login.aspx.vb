Imports Diadoc.Api
Imports Diadoc.Api.Cryptography
Imports System.Data.SqlClient

Public Class Login
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка страницы авторизации
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Not IsPostBack Then
            '------------проверка создания лога для приложения
            If CheckLogsAvl() = True Then
                '------------чтение параметров из конфигурации
                If ReadParameters() = True Then
                    ButtonLogin.Enabled = True
                Else
                    ButtonLogin.Enabled = False
                End If
            Else
                ButtonLogin.Enabled = False
            End If
        End If
    End Sub

    Private Function ReadParameters() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// чтение параметров из конфигурационного файла
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Try
            Declarations.DiadocApiURL = My.Settings.DiadocApiURL
        Catch ex As Exception
            '--------ошибка - не прочитать DiadocApiURL
            LabelErr.Text = "Ошибка - не прочитать DiadocApiUR из конфигурационного файла."
            DivErr.Disabled = False
            DivErr.Visible = True
            DivBG.Disabled = False
            DivBG.Visible = True
            ReadParameters = False
            Exit Function
        End Try

        Try
            Declarations.DeveloperKey = My.Settings.DeveloperKey
        Catch ex As Exception
            '--------ошибка - не прочитать DeveloperKey
            LabelErr.Text = "Ошибка - не прочитать DeveloperKey из конфигурационного файла."
            DivErr.Disabled = False
            DivErr.Visible = True
            DivBG.Disabled = False
            DivBG.Visible = True
            ReadParameters = False
            Exit Function
        End Try

        Try
            Declarations.Debug = My.Settings.MyDebug
        Catch ex As Exception
            '--------ошибка - не прочитать Debug
            LabelErr.Text = "Ошибка - не прочитать настройку отладки из конфигурационного файла."
            DivErr.Disabled = False
            DivErr.Visible = True
            DivBG.Disabled = False
            DivBG.Visible = True
            ReadParameters = False
            Exit Function
        End Try

        Try
            Declarations.ConnString = My.Settings.conString
        Catch ex As Exception
            '--------ошибка - не прочитать строку соединения с БД
            LabelErr.Text = "Ошибка - не прочитать строку соединения с БД из конфигурационного файла."
            DivErr.Disabled = False
            DivErr.Visible = True
            DivBG.Disabled = False
            DivBG.Visible = True
            ReadParameters = False
            Exit Function
        End Try

        ReadParameters = True
    End Function

    Private Function CheckLogsAvl() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка наличия созданного лога для приложения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Try
            If EventLog.SourceExists("AcceptanceFromEDO") = False Then '--первый раз запустить от имени администратора для создания лога
                EventLog.CreateEventSource("AcceptanceFromEDO", "Application")
            End If
        Catch ex As Exception
            '--------ошибка - лог не создан
            LabelErr.Text = "Ошибка - необходимо создать EventLog ""AcceptanceFromEDO"" (от имени администратора)."
            DivErr.Disabled = False
            DivErr.Visible = True
            DivBG.Disabled = False
            DivBG.Visible = True
            CheckLogsAvl = False
            Exit Function
        End Try
        CheckLogsAvl = True
    End Function

    Private Sub ButtonErr_Click(sender As Object, e As EventArgs) Handles ButtonErr.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// закрытие Div с сообщением об ошибке
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LabelErr.Text = ""
        DivErr.Disabled = True
        DivErr.Visible = False
        DivBG.Disabled = True
        DivBG.Visible = False
    End Sub

    Private Sub ButtonLogin_Click(sender As Object, e As EventArgs) Handles ButtonLogin.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка пароля и переход на страницу со списком доступных инвойсов и заказов на закупку
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckData() = True Then
            If CheckPassword() = True Then
                If CheckScalaLogin() = True Then
                    Session("Login") = Trim(Login1.Text)
                    Session("Password") = Trim(Password1.Value)
                    Response.Redirect("IncomingInvoicesList.aspx", True)
                Else
                    '--------ошибка не найден логин в Scala
                    LabelErr.Text = "Ошибка - не найден незаблокированный логин в Scala или код продавца, соответствующий введенному логину."
                    DivErr.Disabled = False
                    DivErr.Visible = True
                    DivBG.Disabled = False
                    DivBG.Visible = True
            End If
            Else
                '--------ошибка - неверный логин / пароль
                LabelErr.Text = "Ошибка аутентификации - неверный логин / пароль."
                DivErr.Disabled = False
                DivErr.Visible = True
                DivBG.Disabled = False
                DivBG.Visible = True
            End If
        End If
    End Sub

    Private Function CheckData() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка заполненности полей
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(Login1.Text) = "" Then
            '--------ошибка - не введен логин
            LabelErr.Text = "ошибка - необходимо ввести логин."
            DivErr.Disabled = False
            DivErr.Visible = True
            DivBG.Disabled = False
            DivBG.Visible = True
            CheckData = False
            Exit Function
        End If

        If Trim(Password1.Value) = "" Then
            '--------ошибка - не введен пароль
            LabelErr.Text = "ошибка - необходимо ввести пароль."
            DivErr.Disabled = False
            DivErr.Visible = True
            DivBG.Disabled = False
            DivBG.Visible = True
            CheckData = False
            Exit Function
        End If
        CheckData = True
    End Function

    Private Function CheckPassword() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка введенного логина и пароля
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim Crypt As WinApiCrypt
        Dim MyApi As DiadocApi
        Dim MyToken As String

        Try
            Crypt = New WinApiCrypt
            MyApi = New DiadocApi(Declarations.DeveloperKey, Declarations.DiadocApiURL, Crypt)
            Session("MyApi") = MyApi
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", ex.Message)
            End If
            CheckPassword = False
            Exit Function
        End Try

        MyToken = GetMyToken(MyApi)
        If MyToken = "" Then
            CheckPassword = False
            Exit Function
        End If
        Session("MyToken") = MyToken

        CheckPassword = True
    End Function

    Private Function GetMyToken(ByRef MyApi As DiadocApi) As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение токена для работы с API
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyTokenLogin As String

        MyTokenLogin = ""
        Try
            MyTokenLogin = MyApi.Authenticate(Trim(Login1.Text), Trim(Password1.Value))
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", ex.Message)
            End If
        End Try

        GetMyToken = MyTokenLogin
    End Function

    Private Function CheckScalaLogin() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка наличия незаблокированного логина в Scala, соответствующего введенному логину и паролю
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                          '--рабочая строка
        Dim ds As New DataSet()

        Try
            MySQLStr = "SELECT ST010300.ST01001 AS SC,  ScalaSystemDB.dbo.ScaUsers.UserName AS SN "
            MySQLStr = MySQLStr & "FROM  ScalaSystemDB.dbo.ScaUsers WITH (NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 INNER JOIN "
            MySQLStr = MySQLStr & "tbl_EDO_Logins ON UPPER(ScalaSystemDB.dbo.ScaUsers.UserName) = UPPER(tbl_EDO_Logins.ScalaLogin) "
            MySQLStr = MySQLStr & "WHERE (tbl_EDO_Logins.EDOLogin = N'" & Trim(Login1.Text) & "') "
            MySQLStr = MySQLStr & "AND (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) "

            Using MyConn As SqlConnection = New SqlConnection(Declarations.ConnString)
                Try
                    MyConn.Open()
                    Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                        cmd.CommandType = CommandType.Text
                        Using da As New SqlDataAdapter()
                            da.SelectCommand = cmd
                            da.Fill(ds)
                            If ds.Tables(0).Rows.Count < 1 Then
                                '---записей нет
                                MyConn.Close()
                                If My.Settings.MyDebug = "YES" Then
                                    EventLog.WriteEntry("AcceptanceFromEDO", "нет незаблокированного логина в Scala и кода продавца, соответствующего введенному логину и паролю.")
                                End If
                                CheckScalaLogin = False
                                Exit Function
                            Else
                                Session("SalesmanCode") = ds.Tables(0).Rows(0).Item(0).ToString
                                Session("SalesmanName") = ds.Tables(0).Rows(0).Item(1).ToString
                            End If
                        End Using
                    End Using
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("AcceptanceFromEDO", ex.Message)
                    End If
                    CheckScalaLogin = False
                    Exit Function
                Finally
                    MyConn.Close()
                End Try
            End Using
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("AcceptanceFromEDO", ex.Message)
            End If
            CheckScalaLogin = False
            Exit Function
        End Try

        CheckScalaLogin = True
    End Function
End Class