<%@ Page Title="Create Project Baseline" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="CreateProjectBaseline.aspx.vb" Inherits="RemedyFM.CreateProjectBaseline" ClientIDMode="Static" %>
<%@ Register Assembly="Infragistics4.Web.v16.1, Version=16.1.20161.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" Namespace="Infragistics.Web.UI.ListControls" TagPrefix="ig" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <link href="Styles/Main.css" rel="stylesheet" type="text/css" />
    <link href="Styles/Loader.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" href="Styles/RFM.Default.css" type="text/css" runat="server" />
    <script type="text/javascript" src="Scripts/Defaults.js"> </script>
    <script type="text/javascript" src="Scripts/ProjectBaseline.js"> </script>
    <script type="text/javascript" id="igClientScript">

        function ShowWheel()
        {
            document.getElementById('btnCreateNew').style.display = "none";
            document.getElementById('div_loader').style.display = "block";
        }

        function HideMsgWindow()
        {
            document.getElementById('divMsg').style.display = "none";
        }

        
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="body" runat="server">
    <div style="width:720px; float:left; text-align:right; padding:4px;"></div>
    <table style="width: 100%; border-collapse: collapse;" cellpadding="0;" cellspacing="0">
        <tr>
            <td>
                The Project Baseline should be created only when the SOW or a Project Change Request is approved. When created, the new Project Baseline will replace the previous one.
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                Please select Project Name: 
            </td>
        </tr>
        <asp:UpdatePanel ID="UpdatePanel2" runat="server">
            <ContentTemplate>
        <tr><td>
            <ig:WebDropDown ID="wddRelease" runat="server" DropDownContainerWidth="330px" AutoPostBackFlags-SelectionChanged="On" AutoPostBackFlags-ValueChanged="Off" TabIndex="1" Width="255px">
            </ig:WebDropDown>
            <asp:TextBox ID="tbxSelectedRelease" runat="server" ClientIDMode="Static"  Visible="True" style="display:none" />
            <asp:TextBox ID="tbxBaselineDate" runat="server" ClientIDMode="Static"  Visible="True" style="display:none" />
            <asp:TextBox ID="tbxErrorMsg" runat="server" ClientIDMode="Static"  Visible="True" style="display:none" />
        </td></tr>
        <tr>
            <td>&nbsp</td>
        </tr>
            <tr>
                <td>
                    Current Project Baseline Date: <asp:TextBox ID="tbxDate" runat="server" ClientIDMode="Static" ReadOnly="True" />
                </td>
            </tr>
            <tr>
                <td>&nbsp</td>
            </tr>
            <tr>
                <td>
                    <asp:Button ID="btnConfirm" runat="server" Text="Create Baseline" Width="120px" TabIndex="2" OnClientClick="CreateProjectBaselineConfirm(event)" Enabled="False" />
                </td>
            </tr>
            </ContentTemplate>
        </asp:UpdatePanel>
    </table>    
</asp:Content>
