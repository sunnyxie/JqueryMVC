<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ReleaseSearchDialog.ascx.vb" Inherits="RemedyFM.ReleaseSearchDialog" %>

<%@ Register Assembly="Infragistics4.Web.v16.1, Version=16.1.20161.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" Namespace="Infragistics.Web.UI.LayoutControls" TagPrefix="ig" %>
<%@ Register Assembly="Infragistics4.Web.v16.1, Version=16.1.20161.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" Namespace="Infragistics.Web.UI.GridControls" TagPrefix="ig" %>
<%@ Register Assembly="Infragistics4.Web.v16.1, Version=16.1.20161.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" Namespace="Infragistics.Web.UI.EditorControls" TagPrefix="ig" %>
<%@ Register Assembly="Infragistics4.Web.v16.1, Version=16.1.20161.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" Namespace="Infragistics.Web.UI" TagPrefix="ig" %>
<%@ Register Assembly="Infragistics4.Web.v16.1, Version=16.1.20161.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" Namespace="Infragistics.Web.UI.ListControls" TagPrefix="ig" %>

<ig:WebDialogWindow runat="server" ID="wdwRelease" Height="450px" Width="542px"
    Modal="true" Moveable="true" Top="55px" AutoPostBackFlags-WindowStateChange="Off"
    Left="110px" InitialLocation="Centered" WindowState="Hidden">
    
    <Header CloseBox-Visible="false" MaximizeBox-Visible="false" MinimizeBox-Visible="false"
        CaptionText="Select the Release Number" Font-Size="Medium" BackColor="#888888">
        <MaximizeBox Visible="false"></MaximizeBox>
        <CloseBox Visible="False" />
    </Header>
    
    <ContentPane>
        <Template>
            <asp:UpdatePanel ID="upReleasePanel" runat="server" UpdateMode="Conditional">
                <ContentTemplate>
                    <br />
                    <table>
                        <tr>
                            <td colspan="3"><span>Please enter <strong>Release number and/or Release name(Wildcards '%')</strong></span></td>
                            <td></td>
                        </tr>
                        <tr>
                            <td><span>Release Number:</span></td>
                            <td>
                                <asp:TextBox ID="txtReleaseNumber" runat="server"></asp:TextBox>
                            </td>
                            <td></td>
                            <td style="text-align: left;">
                               <ig:WebDropDown ID="wddStatus" MultipleSelectionType="Checkbox" EnableMultipleSelection="true" runat="server"
                                 EnableClosingDropDownOnSelect="false" Width="98px" DisplayMode="DropDownList" EnableAutoCompleteFirstMatch="false" EnableCustomValues="false" EnableCustomValueSelection="false" 
                                  ReadOnly="true" DropDownContainerWidth="128" CssClass="PopupDropdown" StyleSetName="default">
                            </ig:WebDropDown>
                            </td>
                        </tr>
                        <tr>
                            <td><span>Release Name:</span></td>
                            <td>
                                <asp:TextBox ID="txtReleaseName" runat="server"></asp:TextBox>
                            </td>
                            <td></td>
                            <td style="text-align: left;">
                               <asp:Button ID="btnSearch" runat="server" Text="Search" Width="98px" align="right"/>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3"></td>
                            <td></td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <ig:WebDataGrid ID="wdgRelease" runat="server" Height="290px" Width="494px"
                                    autopostback="false" AutoGenerateColumns="False">
                                    <Columns>
                                        <ig:BoundDataField DataFieldName="Key" Key="Key" Width="105px" Header-Text="Release Number">
                                            <Header Text="Release Number"></Header>
                                        </ig:BoundDataField>
                                        <ig:BoundDataField DataFieldName="Value" Key="Value" Header-Text="Release Name">
                                            <Header Text="Release Name"></Header>
                                        </ig:BoundDataField>
                                        <ig:BoundDataField DataFieldName="Status" Key="Status" Width="88px" Header-Text="Status">
                                            <Header Text="Status"></Header>
                                        </ig:BoundDataField>

                                    </Columns>
                                    <Behaviors>
                                        <ig:Activation>
                                        </ig:Activation>
                                        <ig:ColumnResizing>
                                        </ig:ColumnResizing>
                                        <ig:Selection Enabled="true" CellClickAction="Row" RowSelectType="Single">
                                            <AutoPostBackFlags CellSelectionChanged="false" RowSelectionChanged="true" ColumnSelectionChanged="false" />
                                        </ig:Selection>
                                        <ig:Paging Enabled="true" PagerMode="NumericFirstLast" PageSize="20" QuickPages="4" CurrentPageLinkCssClass="CurrentPager" PageLinkCssClass="OtherPager">
                                        </ig:Paging>
                                    </Behaviors>
                                </ig:WebDataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3">
                                <asp:Label ID="lblError" runat="server" Text="" ForeColor="#CC3300" Font-Bold="True"></asp:Label>
                            </td>
                            <td align="right">
                                <asp:Button ID="btnSelect" runat="server" Text="Select" />
                                <asp:Button ID="btnClose" runat="server" Text="Cancel" />
                            </td>
                        </tr>
                    </table>
                </ContentTemplate>
            </asp:UpdatePanel>
        </Template>
    </ContentPane>
</ig:WebDialogWindow>


