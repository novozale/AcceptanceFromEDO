<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="IncomingInvoicesList.aspx.vb" Inherits="AcceptanceFromEDO.IncomingInvoicesList" EnableEventValidation="False" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>Выбор инвойсов для загрузки</title>
    <link rel="stylesheet" type="text/css" href="CSS/StyleSheet1.css" />
    <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js"></script>
    <script type="text/javascript">
        function ShowProgress() {
            setTimeout(function () {
                var blockBG = $(".blockBG");
                blockBG.show();
                var loading = $(".loading");
                loading.show();
            }, 1000);
        }
        $(document).ready(function () {
            $('#ButtonRun').click(function () {
                ShowProgress();
            });
        })
        $(document).ready(function () {
            $('#ButtonGet').click(function () {
                ShowProgress();
            });
        })
        $(document).ready(function () {
            $('#GridView1').click(function () {
                ShowProgress();
            });
        })
        $(document).ready(function () {
            $('#PagesList1').change(function () {
                ShowProgress();
            });
        })
        $(document).ready(function () {
            $('#QTYOnPageList1').change(function () {
                ShowProgress();
            });
        })
    </script>
    <script type = "text/javascript">
        window.onload = function () {
            var scrollY = parseInt('<%=Request.Form("scrollY")%>');
        if (!isNaN(scrollY)) {
            window.scrollTo(0, scrollY);
        }
    };
    window.onscroll = function () {
        var scrollY = document.body.scrollTop;
        if (scrollY == 0) {
            if (window.pageYOffset) {
                scrollY = window.pageYOffset;
            }
            else {
                scrollY = (document.body.parentElement) ? document.body.parentElement.scrollTop : 0;
            }
        }
        if (scrollY > 0) {
            var input = document.getElementById("scrollY");
            if (input == null) {
                input = document.createElement("input");
                input.setAttribute("type", "hidden");
                input.setAttribute("id", "scrollY");
                input.setAttribute("name", "scrollY");
                document.forms[0].appendChild(input);
            }
            input.value = scrollY;
        }
    };
</script>
 </head>
<body>
    <form id="form1" runat="server">
        <div id="DivHeader">
            <asp:Label ID="LabelTitle" runat="server" Text="Список входящих инвойсов" Font-Names="Arial" Font-Size="10" Font-Bold="True" ForeColor="#003399"></asp:Label>
        </div>
        <div id="DivMainControls">
            <div id="DivMainControlsLeft">
                &nbsp <asp:Label ID="Label1" runat="server" Text="Поставщик:" Font-Names="Arial" Font-Size="10" ></asp:Label>
                <asp:DropDownList ID="SupplierList" runat="server" AutoPostBack="true"></asp:DropDownList>
                <asp:Button ID="ButtonRun" runat="server" Text="Загрузить" />
                <asp:Button ID="ButtonGet" runat="server" Text="Приемка" />
            </div>
            <div id="DivMainControlsRight">
                <asp:Button ID="ButtonQuit" runat="server" Text="Выход" />
            </div>
        </div> 
        <div id="DivBody">
            <div id="DivTable1">
                <div id="DivHeaderTable1">
                    <asp:Label ID="Label2" runat="server" Text="Входящие инвойсы" Font-Names="Arial" Font-Size="10" Font-Bold="True"></asp:Label>
                </div>
                <div id="DivHeader2Table1">
                    <asp:Button ID="InvoiceView" runat="server" Text="Просмотреть" />
                 </div>
                <div id="DivBodyTable1">
                    <asp:GridView ID="GridView1" runat="server" BackColor="White" BorderColor="#E7E7FF" BorderStyle="Groove"
                        BorderWidth="1px" CellPadding="3" DataKeyNames="INN, KPP" Font-Names="Arial" Font-Size="8pt" ShowHeaderWhenEmpty="True" AutoGenerateColumns="False" SelectedIndex="0"
                        PageSize="100" Width="100%" style="word-wrap:break-word;">
                    <FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
                    <HeaderStyle BackColor="#000066" Font-Bold="True" ForeColor="#F7F7F7" />
                    <RowStyle BackColor="#F7F7F7" ForeColor="#4A3C8C" />
                    <SelectedRowStyle BackColor="#4444FF" Font-Bold="True" ForeColor="White" />
                    <Columns>
                        <asp:BoundField DataField="CompanyName" headertext="Компания" ItemStyle-Width="40%" />
                        <asp:BoundField DataField="INN" headertext="ИНН"/>
                        <asp:BoundField DataField="KPP" headertext="КПП"/>
                        <asp:BoundField DataField="DocNum" headertext=" N документа"/>
                        <asp:BoundField DataField="DocDate" headertext="Дата документа"/>
                        <asp:BoundField DataField="DocSumm" headertext="Сумма (с НДС)"/>
                        <asp:BoundField DataField="DocVAT" headertext="НДС"/>
                        <asp:BoundField DataField="DocState" headertext="Состояние"/>
                     </Columns>
                    </asp:GridView>
                </div>
                <div id="DivFooterTable1">
                    <asp:Label ID="Label7" runat="server" Text="Показать страницу номер" Font-Bold="True" Font-Names="Arial" Font-Size="8"></asp:Label>
                    <asp:DropDownList ID="PagesList1" runat="server" Font-Bold="True" Font-Names="Arial" Font-Size="8" AutoPostBack="true">
                    </asp:DropDownList>
                    <asp:Label ID="LabelQTYPages1" runat="server" Text="из" Font-Bold="True" Font-Names="Arial" Font-Size="8"></asp:Label>  &nbsp &nbsp &nbsp
                    <asp:Label ID="Label8" runat="server" Text="Показывать по " Font-Bold="True" Font-Names="Arial" Font-Size="8"></asp:Label>
                    <asp:DropDownList ID="QTYOnPageList1" runat="server" Font-Bold="True" Font-Names="Arial" Font-Size="8" AutoPostBack="true">
                        <asp:ListItem>10</asp:ListItem>
                        <asp:ListItem>12</asp:ListItem>
                        <asp:ListItem>15</asp:ListItem>
                    </asp:DropDownList>
                    <asp:Label ID="Label9" runat="server" Text=" записей на странице" Font-Bold="True" Font-Names="Arial" Font-Size="8"></asp:Label> &nbsp &nbsp
                </div>
            </div>
            <div id="DivTable2">
                <div id="DivHeaderTable2">
                    <asp:Label ID="Label3" runat="server" Text="Заказы на закупку для выбранного инвойса" Font-Names="Arial" Font-Size="10" Font-Bold="True"></asp:Label>
                </div>
                <div id="DivHeader2Table2">
                    <asp:Label ID="Label4" runat="server" Text="Склад номер" Font-Bold="True" Font-Names="Arial" Font-Size="8"></asp:Label>
                    <asp:DropDownList ID="WHList" runat="server" Font-Bold="True" Font-Names="Arial" Font-Size="8" AutoPostBack="true">
                        <asp:ListItem Value="01" Text="01 Санкт Петербург"></asp:ListItem>
                        <asp:ListItem Value="03" Text="03 Москва"></asp:ListItem>
                    </asp:DropDownList>&nbsp &nbsp
                    <asp:Button ID="ButtonSelectAll" runat="server" Text="Выбрать все заказы" />
                    <asp:Button ID="ButtonDeSelectAll" runat="server" Text="Снять выбор со всех заказов" />
                </div>
                <div id="DivBodyTable2">
                    <asp:GridView ID="GridView2" runat="server" Width="100%" BackColor="White" BorderColor="#E7E7FF" BorderStyle="Groove"
                        BorderWidth="1px" CellPadding="3" DataKeyNames="OrderNum" Font-Names="Arial" Font-Size="8pt" ShowHeaderWhenEmpty="True" AutoGenerateColumns="False"
                        PageSize="100">
                    <FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
                    <HeaderStyle BackColor="#000066" Font-Bold="True" ForeColor="#F7F7F7" />
                    <RowStyle BackColor="#F7F7F7" ForeColor="#4A3C8C" />
                    <SelectedRowStyle BackColor="#4444FF" Font-Bold="True" ForeColor="White" />
                    <Columns>
                        <asp:TemplateField HeaderText="Выбран" ItemStyle-HorizontalAlign="Center">
                            <ItemTemplate>
                                <asp:CheckBox ID="POCheckBox" runat="server"></asp:CheckBox>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="ConsolidatedOrderNum" headertext="Номер обобщенного заказа на закупку"/>
                        <asp:BoundField DataField="OrderNum" headertext="Номер заказа на закупку"/>
                        <asp:BoundField DataField="OrderType" headertext="Тип заказа"/>
                        <asp:BoundField DataField="OrderDate" headertext="Дата заказа"/>
                        <asp:BoundField DataField="WHNum" headertext="Номер склада"/>
                        <asp:BoundField DataField="OrderSum" headertext="Сумма заказа (руб)"/>
                        <asp:BoundField DataField="SalesOrderNum" headertext="Номер заказа на продажу"/>
                        <asp:TemplateField HeaderText="Действие" ItemStyle-HorizontalAlign="Center">
                            <ItemTemplate>
                                <asp:Button ID="ButtonOrderView" runat="server" Text="Просмотр заказа" CommandName="PurchaseOrderView" CommandArgument='<%#Bind("OrderNum") %>'></asp:Button>
                            </ItemTemplate>
                        </asp:TemplateField>
                     </Columns>
                    </asp:GridView>
                </div>
             </div>
        </div>
        <div id="blockBG" class="blockBG" >
        </div>
        <div id="loading" class="loading" style="align-content:center">
            <asp:Table Width="100%" Height="100%" ID="TableD1" runat="server">
                <asp:TableRow runat="server">
                    <asp:TableCell Width="100%" Height="100%" runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <img src="images/loading.gif" alt="" height="200" width="200" />
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </div>
        <div id="DivBG" runat="server" disabled="true" visible="false" style="position:fixed; top:0; right:0; bottom:0; left:0; background-color:black; position:fixed; overflow: auto; white-space: nowrap; z-index:900; opacity:0.4; filter:alpha(opacity=40);" >
        </div>
        <div id="PDFView" class="PDFView" runat="server" disabled="true" visible="false">
            <div id="DivPDFLiteralHeader" class="DivPDFLiteralHeader" runat="server">
                <asp:Button ID="ButtonPDFLiteralClose" runat="server" Text="Закрыть" />
            </div>
            <div id="DivPDFLiteralBody" class="DivPDFLiteralBody" runat="server">
                <asp:Literal ID="PDFLiteral" runat="server" />
            </div>
         </div>
        <div id="POView" class="POView" runat="server" disabled="true" visible="false">
            <div id="POViewHeader" class="POViewHeader" runat="server">
                <asp:Button ID="ButtonPOViewClose" runat="server" Text="Закрыть" />
            </div>
            <div id="POViewBody" class="POViewBody" runat="server">
                <asp:GridView ID="GridView3" runat="server" BackColor="White" BorderColor="#E7E7FF" BorderStyle="Groove"
                        BorderWidth="1px" CellPadding="3" DataKeyNames="StrNum" Font-Names="Arial" Font-Size="8pt" ShowHeaderWhenEmpty="True" AutoGenerateColumns="False" SelectedIndex="-1"
                        PageSize="100" Width="100%" style="word-wrap:break-word;">
                    <FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
                    <HeaderStyle BackColor="#000066" Font-Bold="True" ForeColor="#F7F7F7" />
                    <RowStyle BackColor="#F7F7F7" ForeColor="#4A3C8C" />
                    <SelectedRowStyle BackColor="#4444FF" Font-Bold="True" ForeColor="White" />
                    <Columns>
                        <asp:BoundField DataField="StrNum" headertext="N строки" />
                        <asp:BoundField DataField="ItemCode" headertext="Код"/>
                        <asp:BoundField DataField="ItemName" headertext="Название"/>
                        <asp:BoundField DataField="Price" headertext="Цена"/>
                        <asp:BoundField DataField="Curr" headertext="Валюта"/>
                        <asp:BoundField DataField="Discount" headertext="Скидка (%)"/>
                        <asp:BoundField DataField="Ordered" headertext="Заказано"/>
                        <asp:BoundField DataField="Received" headertext="Принято"/>
                        <asp:BoundField DataField="Invoiced" headertext="Инвойсировано"/>
                        <asp:BoundField DataField="UOM" headertext="Ед. измерения"/>
                     </Columns>
                    </asp:GridView>
            </div>
         </div>
        <div id="DivErr" runat="server" disabled="true" visible="false" style="border: medium double #000080; height:50%; width:50%; left:25%; top:25%; background-color:white; position:fixed; z-index:999">
            <asp:Table Width="100%" Height="100%" ID="Table7" runat="server"  BorderStyle="Double" BackColor="white">
                <asp:TableRow runat="server" >
                    <asp:TableCell Width="100%" Height="100%" runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:TextBox ID="TxtErr" runat="server" Height="98%" Width="98%" ReadOnly="True" TextMode="MultiLine" style="text-align:center;"></asp:TextBox>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow runat="server" BackColor="#CCCCCC">
                    <asp:TableCell Width="100%" runat="server" HorizontalAlign="Right">
                        <asp:Button ID="ButtonErr" runat="server" Text="OK" Font-Names="Arial" Font-Size="8" />
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </div>
    </form>
</body>
</html>
