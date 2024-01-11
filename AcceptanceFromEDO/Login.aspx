<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Login.aspx.vb" Inherits="AcceptanceFromEDO.Login" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" style="height: 100%;">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>ООО Электроскандия Рус</title>
</head>
<body style="height: 97%;">
    <form id="Login" runat="server" style="height: 97%;">
         <asp:Table runat="server" Width="100%" style="height: 97%;" >
            <asp:TableRow runat="server" Height="10%">
                <asp:TableCell  runat="server" >
                    <asp:Table Width="100%"  ID="Table1" runat="server">
                        <asp:TableRow runat="server" VerticalAlign="Middle">
                            <asp:TableCell  Width="10%" runat="server" VerticalAlign="Middle">
                                <asp:Image id="Image1" runat="server" ImageUrl="Images/Elektroskandia.jpg" Height="22px" Width="154px" /> 
                            </asp:TableCell>
                            <asp:TableCell Width="100%" runat="server" VerticalAlign="Middle" >
                                &nbsp &nbsp <asp:Label ID="Label2" runat="server" Text="Рабочее место автоматической приемки товаров из ЭДО" Font-Names="Arial" Font-Bold="True" Font-Size="Small" ForeColor="#003399"></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow runat="server" Height="90%">
                <asp:TableCell  runat="server" >
                    <asp:Table Width="100%" Height="100%" ID="Table2" runat="server" HorizontalAlign="Center" >
                        <asp:TableRow runat="server" Height="50%" VerticalAlign="Bottom">
                            <asp:TableCell  runat="server" Width="45%" HorizontalAlign="Right" >
                                <asp:Label ID="Label1" runat="server" Text="Логин" Font-Bold="True" Font-Names="Arial" Font-Size="Smaller" ForeColor="#003399"></asp:Label> &nbsp
                            </asp:TableCell>
                            <asp:TableCell  runat="server" Width="10%" >
                                <asp:TextBox ID="Login1" runat="server" Font-Bold="True" Font-Names="Arial" Font-Size="Smaller" Width="98%"></asp:TextBox>
                            </asp:TableCell>
                            <asp:TableCell  runat="server" Width="45%" >
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow runat="server" Height="50%" VerticalAlign="Top">
                            <asp:TableCell  runat="server" Width="45%" HorizontalAlign="Right" >
                                <asp:Label ID="Label3" runat="server" Text="Пароль" Font-Bold="True" Font-Names="Arial" Font-Size="Smaller" ForeColor="#003399" ></asp:Label> &nbsp
                            </asp:TableCell>
                            <asp:TableCell  runat="server" Width="10%" >
                                <input id="Password1" type="password" runat="server" size="30"/>
                            </asp:TableCell>
                            <asp:TableCell  runat="server" Width="45%" >
                                &nbsp<asp:Button ID="ButtonLogin" runat="server" Text="Вход" Font-Bold="True" Font-Names="Arial" Font-Size="Small" ForeColor="#003399" />
                            </asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <div id="DivErr" runat="server" disabled="true" visible="false" style="border: medium double #000080; height:50%; width:50%; left:25%; top:25%; background-color:white; position:absolute; z-index:2">
            <asp:Table Width="100%" Height="100%" ID="TableD1" runat="server"  BorderStyle="Double" BackColor="White">
                <asp:TableRow runat="server" >
                    <asp:TableCell Width="100%" Height="100%" runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:Label ID="LabelErr" runat="server" Text="Label" Font-Names="Arial" Font-Size="Medium" Font-Bold="True" ForeColor="#FF3300"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow runat="server" BackColor="#CCCCCC">
                    <asp:TableCell Width="100%" runat="server" HorizontalAlign="Right">
                        <asp:Button ID="ButtonErr" runat="server" Text="OK" Font-Names="Arial" Font-Size="8" />
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
         </div>
        <div id="DivBG" runat="server" disabled="true" visible="false" style="height:100%; width:100%; left:0; top:0; background-color:black; position:absolute; z-index:1; opacity:0.6; filter:alpha(opacity=60);" >
        </div>
    </form>
</body>
</html>

