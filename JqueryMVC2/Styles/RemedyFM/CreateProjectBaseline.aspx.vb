Imports Infragistics.Web.UI.ListControls

Public Class CreateProjectBaseline
    Inherits System.Web.UI.Page

    Protected mMasterObject As RFM.BusinessObject.ProjectBaselineObj
    Private UserAccessLevel As Integer = Values.Client_Role_Users
    Protected regexException As Regex = New Regex("ORA-\d{5}\s*:(.+)\s*ORA-\d{5}")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        UserAccessLevel = Global_asax.USER_ACCESS_LEVEL
        If IsPostBack Then

        Else
            LoadProjectData()
        End If
    End Sub

    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        Try
            mMasterObject = New RFM.BusinessObject.ProjectBaselineObj

            Dim loSelectedItem As DropDownItem = wddRelease.SelectedItem
            tbxSelectedRelease.Text = ""
            tbxBaselineDate.Text = ""

            If (loSelectedItem Is Nothing) Then
                Return
            ElseIf loSelectedItem.Value = "0" Then
                Return
            End If

            Try
                ' Inform the Remedy side
                RemedySideProcess(wddRelease.SelectedItem.Value) ' My.Settings.userRemedy

                mMasterObject.GetRemedyUpdateStatus(wddRelease.SelectedItem.Value)
                Dim entBaseLine As RFM.BusinessObject.ProjectBaselineEntity = mMasterObject.GetCurrentEntity
                If mMasterObject.GetRowCount <> 1 OrElse entBaseLine.ProjectName = "NONE" _
                    OrElse entBaseLine.ProjectName = String.Empty Then
                    tbxErrorMsg.Text = "Remedy side error: " & vbCrLf _
                     & "We can’t update the baseline dates in Remedy now. Please try again later."
                    Page.ClientScript.RegisterStartupScript(Type.GetType("System.String"), "addScript", "ShoweErrorMsg();", True)
                    Return
                End If
            Catch ex As Exception
                tbxErrorMsg.Text = "Remedy side error: " & ex.Message & vbCrLf _
                     & "We can’t update the baseline dates in Remedy now. Please try again later."
                Page.ClientScript.RegisterStartupScript(Type.GetType("System.String"), "addScript", "ShoweErrorMsg();", True)
                Return
            End Try

            tbxSelectedRelease.Text = loSelectedItem.Text
            mMasterObject.CreateDataByProjectID(loSelectedItem.Value)

            For Each loEntity As RFM.BusinessObject.ProjectBaselineEntity In mMasterObject.GetCNRLEntities(Nothing, False)
                tbxBaselineDate.Text = loEntity.ProjectName
            Next

            If (tbxBaselineDate.Text <> "") Then
                Page.ClientScript.RegisterStartupScript(Type.GetType("System.String"), "addScript2", "showmessage()", True)
            End If


        Catch ex As Exception
            Dim msg As String
            Dim match As Match = regexException.Match(ex.Message)
            If match.Success Then
                msg = "Project baseline creation error." & vbCrLf & match.Groups(1).Value
            Else
                msg = "Project baseline creation error." & vbCrLf & ex.Message
            End If

            tbxErrorMsg.Text = msg
            Page.ClientScript.RegisterStartupScript(Type.GetType("System.String"), "addScript", "ShoweErrorMsg()", True)
        End Try
    End Sub

    Private Sub LoadProjectData()
        Try
            mMasterObject = New RFM.BusinessObject.ProjectBaselineObj
            mMasterObject.SetFilterValue("_AccessLevel", UserAccessLevel)
            wddRelease.DataSourceID = Nothing
            wddRelease.Items.Clear()
            ' wddRelease.AppendDataBoundItems = True

            wddRelease.DataSource = mMasterObject.GetCNRLEntities(Nothing)
            wddRelease.ValueField = "ProjectID"
            wddRelease.TextField = "ProjectName"
            wddRelease.DataBind()
            wddRelease.Enabled = True
            'wddRelease.Items.Add(New DropDownItem(String.Empty, String.Empty))
        Catch ex As Exception
            tbxErrorMsg.Text = ex.Message
            Page.ClientScript.RegisterStartupScript(Type.GetType("System.String"), "addScript", "ShoweErrorMsg()", True)
        End Try

    End Sub

    Private Sub wddRelease_SelectionChanged(sender As Object, e As DropDownSelectionChangedEventArgs) Handles wddRelease.SelectionChanged
        Try
            mMasterObject = New RFM.BusinessObject.ProjectBaselineObj

            Dim loSelectedItem As DropDownItem = wddRelease.SelectedItem
            tbxSelectedRelease.Text = ""
            tbxDate.Text = ""

            If (loSelectedItem.Value = "0") Then
                btnConfirm.Enabled = False
                Return
            End If

            tbxSelectedRelease.Text = loSelectedItem.Text
            mMasterObject.GetDataByProjectID(loSelectedItem.Value)

            For Each loEntity As RFM.BusinessObject.ProjectBaselineEntity In mMasterObject.GetCNRLEntities(Nothing, False)
                tbxDate.Text = loEntity.ProjectName
            Next

            btnConfirm.Enabled = True
        Catch ex As Exception
            tbxErrorMsg.Text = ex.Message
            Page.ClientScript.RegisterStartupScript(Type.GetType("System.String"), "addScript", "ShoweErrorMsg()", True)
        End Try
    End Sub

    Public Function RemedySideProcess(ReleaseID As String) As Boolean
        Dim Request As RemedyReference.Release_Modify_ServiceRequest
        Dim Response As RemedyReference.Release_Modify_ServiceResponse
        Try

            Using remedyClient As RemedyReference.RMS_Release_UpdateBaselineFlagPortPortTypeClient =
                 New RemedyReference.RMS_Release_UpdateBaselineFlagPortPortTypeClient()

                Dim AuthenInfo As New RemedyReference.AuthenticationInfo()
                AuthenInfo.userName = My.Settings.userRemedy
                AuthenInfo.password = My.Settings.passwordRemedy

                ' Or ...
                Request = New RemedyReference.Release_Modify_ServiceRequest(AuthenInfo,
                                                    RemedyReference.BaselineType.New, ReleaseID)

                Response = CType(remedyClient, RemedyReference.RMS_Release_UpdateBaselineFlagPortPortType).Release_Modify_Service(Request)
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message)
            'tbxErrorMsg.Text = "Remedy side error: " & ex.Message
            'Page.ClientScript.RegisterStartupScript(Type.GetType("System.String"), "addScript", "ShoweErrorMsg()", True)
        End Try

        Return True
    End Function
End Class