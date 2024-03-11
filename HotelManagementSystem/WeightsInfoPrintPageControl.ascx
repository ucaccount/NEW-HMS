<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="SampleControl.ascx.cs" Inherits="YourNamespace.SampleControl" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Sample ASCX Control</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <h2>Sample ASCX Control Content</h2>
            
            <asp:Label ID="lblMessage" runat="server" Text="Hello from ASCX Control!" />
            
            <br />
            
            <asp:Button ID="btnClickMe" runat="server" Text="Click Me" OnClick="btnClickMe_Click" />
        </div>
    </form>
</body>
</html>