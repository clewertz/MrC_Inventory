<%@ Page Title="Query" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeFile="Query.aspx.cs" Inherits="_Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <p class="lead">
            <asp:Label ID="lblSearch" runat="server" Text="Search: "></asp:Label>
            <asp:TextBox ID="txtPCList" runat="server" Height="144px" TextMode="MultiLine"></asp:TextBox>
        </p>
        <p class="lead">
            <asp:Button ID="btnSearch" runat="server" OnClick="btnSearch_Click" Text="Button" />
        </p>
        <p class="lead">
            <asp:Label ID="lblReturn" runat="server" Text="Label"></asp:Label>
        </p>
       
    </div>

</asp:Content>
