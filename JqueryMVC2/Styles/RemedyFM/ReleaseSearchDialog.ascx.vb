Imports CanadianNatural.ApplicationServices.BusinessObjects.MereMortals
Imports RFM.BusinessObject

Public Class ReleaseSearchDialog
    Inherits System.Web.UI.UserControl

    Public Event ReleaseSelected()

    Private mstrReleaseNumber As String = ""
    Private mstrReleaseName As String = ""
    Private miPageIndicator As Integer = Values.RELEASE_POP_IN_OTHER
    Protected regexException As Regex = New Regex("ORA-\d{5}\s*:(.+)\s*ORA-\d{5}")

#Region " Public Properties and Routines. "
    Public Property ReleaseNumber() As String
        Get
            Return mstrReleaseNumber
        End Get
        Set(value As String)
            mstrReleaseNumber = value
        End Set
    End Property

    Public Property ReleaseName As String
        Get
            Return mstrReleaseName
        End Get
        Set(value As String)
            mstrReleaseName = value
        End Set
    End Property

    Public Property PageIndicator As Integer
        Get
            Return miPageIndicator
        End Get
        Set(value As Integer)
            miPageIndicator = value
        End Set
    End Property

    Public Sub MakeVisible(pblnVisible As Boolean)
        If pblnVisible Then
            wdwRelease.WindowState = Infragistics.Web.UI.LayoutControls.DialogWindowState.Normal
            ResetGrid()
            ApplySearch()
        Else
            wdwRelease.WindowState = Infragistics.Web.UI.LayoutControls.DialogWindowState.Hidden
        End If
    End Sub
#End Region

#Region " Form Events. "
    Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Page.IsPostBack Then
            mstrReleaseNumber = CType(ViewState("mstrReleaseNumber"), String)
            mstrReleaseName = CType(ViewState("mstrReleaseName"), String)
            miPageIndicator = CType(ViewState("miPageIndicator"), Integer)
        Else
            wdgRelease.Behaviors.Paging.PageIndex = 0
            wdgRelease.Behaviors.Selection.SelectedRows.Clear()
            wddStatus.CurrentValue = String.Empty
            LoadComboBox(RFM.BusinessObject.Constants.Reference.ReleaseStatus)
            SetDefaultValues(wddStatus)
            'ApplySearch()
        End If

    End Sub

    Protected Sub SetDefaultValues(wdd As Infragistics.Web.UI.ListControls.WebDropDown)
        For i = 0 To wdd.Items.Count - 1
            Dim val As String = wdd.Items(i).Value.ToString

            If val = Values.ProjectStatusInitiate OrElse
              val = Values.ProjectStatusPlanning OrElse
              val = Values.ProjectStatusBuild Then
                wdd.Items(i).Selected = True
            End If
        Next
    End Sub

    Private Sub Page_PreRender(sender As Object, e As EventArgs) Handles Me.PreRender
        ' Store the module level variables because web pages don't retain their values.

        ' Store the simple data type variables.
        ViewState.Add("mstrReleaseNumber", mstrReleaseNumber)
        ViewState.Add("mstrReleaseName", mstrReleaseName)
        ViewState.Add("miPageIndicator", miPageIndicator)
    End Sub
#End Region

#Region " Control Events. "
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        lblError.Text = ""
        MakeVisible(False)
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        lblError.Text = ""

        wdgRelease.Behaviors.Paging.PageIndex = 0
        wdgRelease.Behaviors.Selection.SelectedRows.Clear()

        ApplySearch()
    End Sub

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        lblError.Text = ""

        If wdgRelease.Behaviors.Selection.SelectedRows.Count = 0 Then
            lblError.Text = "Please select a valid release first"
        Else
            ''''' Since Infragistics does not give up its data, Using a look up call to get the release name.
            ''''Dim VendorRefList As New ReferenceList(Of ReferenceListEntity)(Values.ApplicationKey, Constants.Reference.VendorNumberFilter, mstrReleaseNumber)
            ''''VendorRefList.GetCNRLEntities(Nothing)
            ''''If VendorRefList.GetRowCount = 0 Then
            ''''    ' Leave the vendor name as a number if it cannot find the data.
            ''''Else
            ''''    mstrReleaseName = VendorRefList.GetCurrentEntity().Value.ToString
            ''''End If

            ' Raise an event the calling form will get and hide the pop up.
            RaiseEvent ReleaseSelected()
            MakeVisible(False)
        End If
    End Sub

    ' User selected another page.
    Private Sub wdgRelease_PageIndexChanged(sender As Object, e As Infragistics.Web.UI.GridControls.PagingEventArgs) Handles wdgRelease.PageIndexChanged
        ApplySearch()
    End Sub

    Private Sub wdgRelease_RowSelectionChanged(sender As Object, e As Infragistics.Web.UI.GridControls.SelectedRowEventArgs) Handles wdgRelease.RowSelectionChanged
        If e.CurrentSelectedRows.Count > 0 Then
            mstrReleaseNumber = e.CurrentSelectedRows.GetIDPair(0).Item("key")(0).ToString
            ''''mstrReleaseName = e.CurrentSelectedRows.GetIDPair(0).Item("key")(0).ToString
        End If
    End Sub
#End Region

    Private Sub ApplySearch()
        Dim filterName As String = txtReleaseName.Text.Trim
        Dim filterNum As String = txtReleaseNumber.Text.Trim
        Dim intUserAccessLevel As Integer = Values.Client_Role_Users

        intUserAccessLevel = Global_asax.USER_ACCESS_LEVEL
        Try

            If Not CheckSearchTextValid(filterNum) OrElse Not CheckSearchTextValid(filterName) Then
                lblError.Text = "Your search text is too long, please limit it to 64 characters."
            Else
                Dim pstrStatus As String = GetSelectedStatuses()

                Dim busSearchList As New SearchList(Of SearchListEntity)
                Dim dstData As System.Data.DataSet = busSearchList.LoadReleaseListWithSecurity(filterNum, filterName, pstrStatus, intUserAccessLevel, miPageIndicator)

                wdgRelease.DataSource = dstData.Tables(0)
                wdgRelease.DataKeyFields = "KEY"
                wdgRelease.DataBind()
            End If
        Catch ex As Exception
            Dim match As Match = regexException.Match(ex.Message)
            If match.Success Then
                lblError.Text = "Error:" & vbCrLf & match.Groups(1).Value
            Else
                lblError.Text = "Error:" & vbCrLf & ex.Message
            End If

        End Try
    End Sub

    Private Function GetSelectedStatuses() As String
        If wddStatus.SelectedItems.Count = 0 Then
            Return String.Empty
        Else
            Dim pstrResult As String = String.Empty
            For Each item In wddStatus.SelectedItems
                pstrResult += item.Value + ","
            Next

            pstrResult = pstrResult.TrimEnd(","c)
            Return pstrResult
        End If
    End Function

    Private Function CheckSearchTextValid(filterNum As String) As Boolean
        If filterNum.Length > 64 Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function IsEmptyOrPercentSign(filter As String) As Boolean
        If filter = String.Empty OrElse filter = "%" OrElse filter = "%%" Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub ResetGrid()
        wdgRelease.Behaviors.Selection.SelectedRows.Clear()
    End Sub

    Private Sub LoadComboBox(pstrReference As String, Optional isNumber As Boolean = False)
        Dim cbo As Infragistics.Web.UI.ListControls.WebDropDown = wddStatus
        Dim objRefList As New AppReflist(pstrReference)
        cbo.Items.Clear()

        cbo.DataSource = objRefList.GetCNRLEntities(Nothing)
        cbo.TextField = "Value"
        cbo.ValueField = "Key"
        cbo.DataBind()
    End Sub
    'Private Function LoadResultList(ByVal refListName As String, filter As String) As ReferenceList(Of ReferenceListEntity)
    '    Dim RefList1 As New ReferenceList(Of ReferenceListEntity)(Values.ApplicationKey, refListName, filter)

    '    Return RefList1
    'End Function
End Class