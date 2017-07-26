<%@ Page Language="C#" AutoEventWireup="true" CodeFile="form1.aspx.cs" Inherits="form1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
   
</head>

<body>
    <form id="form2" runat="server">
    <div>
        <asp:Button ID="Button1" runat="server" Text="导入" OnClick="Button1_Click" />
    
        <asp:LinkButton ID="LinkButton1" runat="server" Visible="false" OnClick="LinkButton1_Click">下载</asp:LinkButton>
    </div>
    </form>
</body>
</html>
