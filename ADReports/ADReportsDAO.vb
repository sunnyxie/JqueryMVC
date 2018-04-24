Imports Oracle.DataAccess.Types
Imports Oracle.DataAccess.Client

Public Class ADReportsDAO
    Inherits clsDBAccess

    Friend Function InsertBulkRecordsToMembers(numRecords As Integer, arMemberType() As String,
                                     arAccountName() As String, arDisplayName() As String,
                                     arAccountType() As String, arSID() As String,
                                     arFromGroup() As String, sMemDomain As String,
                                     RunDateTime As String, arUserIDs() As String) As Boolean

        Dim cmd As OracleCommand = CreateDbCommand("AD_REPORTS.DA_ADREPORTSLOADING.insert_bulk_rows_members")
        'AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_DOMAIN", arMemDomain)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_MEMBERSHIPTYPE", arMemberType)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_ACCOUNTNAME", arAccountName)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_USERIDS", arUserIDs)

        AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_ACCOUNTDISPLAYNAME", arDisplayName)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_ACCOUNTTYPE", arAccountType)

        AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_ACCOUNTSID", arSID)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_FROMGROUP", arFromGroup)

        AddedInputVarChar(cmd, "in_MEMBERS_DOMAIN", sMemDomain)
        AddedInputVarChar(cmd, "in_sig_date", RunDateTime)
        AddedOutputNumber(cmd, "ret_status")
        AddedOutputVarChar(cmd, "ret_msg")

        cmd.ExecuteNonQuery()
        'cnn.Close()
        If (Not IsDBNull(cmd.Parameters("ret_status").Value)) Then

            If (CInt(cmd.Parameters("ret_status").Value.ToString) = 0) Then
                Return True
            Else
                Throw New Exception(cmd.Parameters("ret_msg").Value.ToString)
            End If
        Else
            Return False
        End If

        Return False
    End Function

    Friend Function InsertBulkRecordsToMemTbl(numRecords As Integer, sFromGroup As String,
                                 sMemDomain As String,   arMemberType() As String,
                                 arAccountName() As String, arDisplayName() As String,
                                 arAccountType() As String, arSID() As String,
                                 RunDateTime As String, arUserIDs() As String) As Boolean

        Dim cmd As OracleCommand = CreateDbCommand("AD_REPORTS.DA_ADREPORTSLOADING.insert_bulk_rows_members")
        'AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_DOMAIN", arMemDomain)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_MEMBERSHIPTYPE", arMemberType)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_ACCOUNTNAME", arAccountName)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_USERIDS", arUserIDs)

        AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_ACCOUNTDISPLAYNAME", arDisplayName)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_ACCOUNTTYPE", arAccountType)

        AddNewOracleParameterVarchar(cmd, numRecords, "in_MEMBERS_ACCOUNTSID", arSID)

        AddedInputVarChar(cmd, "in_MEMBERS_FROMGROUP", sFromGroup)
        AddedInputVarChar(cmd, "in_MEMBERS_DOMAIN", sMemDomain)
        AddedInputVarChar(cmd, "in_sig_date", RunDateTime)
        AddedOutputNumber(cmd, "ret_status")
        AddedOutputVarChar(cmd, "ret_msg")

        cmd.ExecuteNonQuery()
        'cnn.Close()
        If (Not IsDBNull(cmd.Parameters("ret_status").Value)) Then

            If (CInt(cmd.Parameters("ret_status").Value.ToString) = 0) Then
                Return True
            Else
                Throw New Exception(cmd.Parameters("ret_msg").Value.ToString)
            End If
        Else
            Return False
        End If

        Return False
    End Function

    Sub AddNewOracleParameterVarchar(cmd As OracleCommand, numRecords As Integer, paramName As String,
                              arValue() As String)
        Dim p1 As New OracleParameter(paramName, OracleDbType.Varchar2, ParameterDirection.Input)
        p1.CollectionType = OracleCollectionType.PLSQLAssociativeArray
        p1.Size = numRecords
        p1.Value = arValue ' From x In {1, 2, 3, 4, 5} Select x
        cmd.Parameters.Add(p1)
    End Sub

    Sub AddNewOracleParameterOther(cmd As OracleCommand, numRecords As Integer, paramName As String,
                          arValue() As Object, type As OracleDbType)
        Dim p1 As New OracleParameter(paramName, type, ParameterDirection.Input)
        p1.CollectionType = OracleCollectionType.PLSQLAssociativeArray
        p1.Size = numRecords
        p1.Value = arValue ' From x In {1, 2, 3, 4, 5} Select x
        cmd.Parameters.Add(p1)
    End Sub

    Friend Function InsertBulkRecordsToADReports(numRecords As Integer, arObjectPath() As String,
                                    arObjectType() As String, arAllowDeny() As String,
                                    arDisplayName() As String, arAccountName() As String,
                                    arAccountType() As String, arFromGroup() As String,
                                    arApplyTo() As String, arApplyDirectChOnly() As String,
                                    arPermissions() As String, RunDateTime As String,
                                    arUserIDs() As String) As Boolean

        Dim cmd As OracleCommand = CreateDbCommand("AD_REPORTS.DA_ADREPORTSLOADING.insert_bulk_new_ad_reports")
        'cmd.CommandText = "associative_array.array_insert"


        AddNewOracleParameterVarchar(cmd, numRecords, "in_object_path", arObjectPath)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_object_type", arObjectType)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_allow_deny", arAllowDeny)

        AddNewOracleParameterVarchar(cmd, numRecords, "in_display_name", arDisplayName)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_account_name", arAccountName)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_user_ids", arUserIDs)

        AddNewOracleParameterVarchar(cmd, numRecords, "in_account_type", arAccountType)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_from_group", arFromGroup)

        AddNewOracleParameterVarchar(cmd, numRecords, "in_apply_to", arApplyTo)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_apply_direct_only", arApplyDirectChOnly)
        AddNewOracleParameterVarchar(cmd, numRecords, "in_permissions", arPermissions)

        AddedInputVarChar(cmd, "in_sig_date", RunDateTime)
        AddedOutputNumber(cmd, "ret_status")
        AddedOutputVarChar(cmd, "ret_msg")

        cmd.ExecuteNonQuery()
        If (Not IsDBNull(cmd.Parameters("ret_status").Value)) Then

            If (CInt(cmd.Parameters("ret_status").Value.ToString) = 0) Then
                Return True
            Else
                Throw New Exception(cmd.Parameters("ret_msg").Value.ToString)
            End If
        Else
            Return False
        End If

        Return False
    End Function

    ' Migrate data from external Tables. 
    Friend Function MigrateDataToEternalTables() As Boolean
        Dim cmd As OracleCommand = CreateDbCommand("AD_REPORTS.DA_ADREPORTSLOADING.MigrateDataToEternalTables")

        AddedOutputNumber(cmd, "ret_status")
        AddedOutputVarChar(cmd, "ret_msg")

        cmd.ExecuteNonQuery()
        If (Not IsDBNull(cmd.Parameters("ret_status").Value)) Then

            If (CInt(cmd.Parameters("ret_status").Value.ToString) = 0) Then
                Return True
            Else
                Throw New Exception(cmd.Parameters("ret_msg").Value.ToString)
            End If
        Else
            Return False
        End If

        Return True
    End Function

End Class
