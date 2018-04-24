Imports System.Data
Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Public Class clsExcel
    Dim DomainNameStarter As String = "CNRL"
    Dim DomainNameStartEM As String = "EMEA"
    Dim AccountIDStarter As String = "CNRL"

    Dim ReportsColIndex As New Dictionary(Of String, Integer)
    Dim MembersColIndex As New Dictionary(Of String, Integer)
    Dim IMRequiredColumns As ICollection(Of String)

    Function GetRequiredColumnMembers() As ICollection(Of String)
        Return {
         Values.ColMembersMembershipType,
        Values.ColMembersAccountName,
        Values.ColMembersDisplayName,
        Values.ColMembersAccountType,
        Values.ColMembersAccountSID,
        Values.ColMembersFromGroup
            }
    End Function

    Function GetRequiredColumnReport() As ICollection(Of String)
        Return {
         Values.ColReportObjectPath.ToUpper(),
         Values.ColReportObjectType.ToUpper(),
         Values.ColReportAllowDeny.ToUpper(),
         Values.ColReportDisplayName.ToUpper(),
         Values.ColReportAccountName.ToUpper(),
         Values.ColReportAccountType.ToUpper(),
         Values.ColReportFromGroup.ToUpper(),
         Values.ColReportApplyTo.ToUpper(),
         Values.ColReportApplyDirectOnly.ToUpper(),
         Values.ColReportPermissions.ToUpper
            }
    End Function

    Friend Shared Function ColumnAsChar(ByVal col As Integer) As String
        Dim alpha() As Char = "0ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray
        Dim x As Integer, y As Integer
        ColumnAsChar = ""
        Try
            If col > 26 Then
                x = Math.DivRem(col, 26, y)
                ColumnAsChar = alpha(x) & alpha(y)
            Else
                ColumnAsChar = alpha(col)
            End If
        Catch ex As Exception
            Trace.WriteLine("Error in clsExcel.ColumnAsChar : " & ex.Source & " : " & ex.ToString, "ADReportsLoading")
        End Try
    End Function

    ' Get all the columns headers into IMColumnIndexes. 
    Function GetExcelRowsOrderForMember(ByRef olecon As System.Data.OleDb.OleDbConnection, ByVal sheetName As String) As Boolean
        Dim bResult As Boolean = False
        Using olecomm As New OleDbCommand
            ' Get the first row, title Row. 
            Dim sCommandTe As String = String.Format("Select {0} From [{1}] ", "TOP 6 * ", sheetName)

            olecomm.CommandText = sCommandTe
            olecomm.Connection = olecon
            Dim bGetDomain As Boolean = False

            Dim reader As OleDbDataReader = olecomm.ExecuteReader()
            Dim colname As String = String.Empty
            While reader.Read

                If Not bGetDomain Then
                    For i As Integer = 0 To reader.FieldCount - 1
                        If reader.IsDBNull(i) Then
                            Continue For
                        End If

                        colname = reader.GetString(i).Trim.ToUpper
                        If i = 0 AndAlso colname.StartsWith(DomainNameStarter) Then
                            bGetDomain = True
                            Exit For
                        End If
                    Next

                    ' Get the Column headers. 
                Else
                    For i As Integer = 0 To reader.FieldCount - 1
                        colname = reader.GetString(i).Trim.ToUpper
                        If IMRequiredColumns.Contains(colname) Then
                            MembersColIndex(colname) = i ' + 1
                        End If
                    Next
                    bGetDomain = False

                    If MembersColIndex.Count = IMRequiredColumns.Count Then
                        bResult = True
                    Else
                        MembersColIndex.Clear()
                    End If
                End If

                If bResult Then
                    Exit While
                End If
            End While

            reader.Close()
            Return bResult
        End Using
    End Function

    ' Get all the columns headers into ReportsColIndex, 
    ' Return the Columns header Row Number.
    Function GetExcelRowsOrderForReport(ByRef olecon As System.Data.OleDb.OleDbConnection,
                                        ByVal sheetName As String) As Integer
        Dim bResult As Boolean = False
        Using olecomm As New OleDbCommand
            ' Get the first row, title Row. 
            Dim sCommandTe As String = String.Format("Select {0} From [{1}] ", "TOP 6 * ", sheetName)

            olecomm.CommandText = sCommandTe
            olecomm.Connection = olecon

            Dim reader As OleDbDataReader = olecomm.ExecuteReader()
            Dim colname As String = String.Empty
            Dim nRowNum As Integer = 0

            While reader.Read
                nRowNum += 1
                ' Empty rows.
                If reader.IsDBNull(0) AndAlso reader.IsDBNull(1) AndAlso
                    reader.IsDBNull(2) AndAlso reader.IsDBNull(3) Then
                    Continue While
                End If

                For i As Integer = 0 To reader.FieldCount - 1
                    If reader.IsDBNull(i) Then
                        Trace.WriteLine("GetExcelRowsOrderForReport, Data Field " & i.ToString & " is DB NULL.")
                        Continue For
                    End If

                    colname = reader.GetString(i).Trim.ToUpper
                    If IMRequiredColumns.Contains(colname) Then
                        ReportsColIndex(colname) = i ' + 1
                    End If
                Next

                If ReportsColIndex.Count = IMRequiredColumns.Count Then
                    bResult = True
                Else
                    ReportsColIndex.Clear()
                End If

                ' Ready to exit.
                If bResult Then
                    Exit While
                End If
            End While

            reader.Close()

            If Not bResult Then
                nRowNum = 0
            End If
            Return nRowNum
        End Using
    End Function

    Function GetSheetName(ID As Integer) As String
        If ID = 1 Then
            Return "Report$"
        Else
            Return "Members$"
        End If
    End Function

    Public Function GetAllDataToDB(sFileName As String, oDBClass As ADReportsDAO, _
                                   InsertDatetime As String, ByRef sErrMsg As String) As Integer

        Dim bHasError As Boolean = False
        Dim sConnString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
        sFileName & ";Extended Properties=""Excel 12.0;IMEX=1;ImportMixedTypes=Text;TypeGuessRows=0;HDR=NO;"""

        Using olecon As New System.Data.OleDb.OleDbConnection(sConnString)

            Try
                olecon.Open()
                Try

                    ADReportsDAO.BeginTransaction()
                    ' oExcel.WriteAllTheDataToDB(fileName, DBConn, DBInsertDatetime.ToString("yyyy-MM-dd HH:mm:ss"))
                    IMRequiredColumns = GetRequiredColumnReport()
                    bHasError = Not GetDataInReportSheet(olecon, oDBClass, InsertDatetime)
                    ADReportsDAO.DBCommit()
                Catch ex As Exception
                    ADReportsDAO.DBRollback()
                    bHasError = True

                    sErrMsg = ex.Message.ToString
                    Trace.WriteLine("Error: clsExcel.GetDataInReportSheet : " & ex.Source & " : " & ex.Message.ToString, "ADReportsLoading")
                    Return ErrorCode.EN_ERROR_CODE.ERROR_IN_WRITE_REPORT_TO_DB
                Finally
                    ADReportsDAO.EndTransaction()
                    Trace.Flush()
                End Try

                Trace.WriteLine(Date.Now.ToString() & ": " & "Report Sheet of the Excel file complete with " &
                    IIf(bHasError, "errors.", "success.").ToString)

                If bHasError Then
                    sErrMsg = "Error in clsExcel.GetDataInReportSheet. "
                    Trace.WriteLine("Error: clsExcel.GetDataInReportSheet. ", "ADReportsLoading")
                    Return ErrorCode.EN_ERROR_CODE.ERROR_IN_WRITE_REPORT_DATA
                End If


            Catch ex As Exception
                Trace.WriteLine("Error: Failed to open OleDbConnection.: " & ex.Message.ToString, "ADReportsLoading")
                sErrMsg = ex.Message.ToString
                Return ErrorCode.EN_ERROR_CODE.ERROR_IN_OPEN_DB_CONNECTION
            Finally
                olecon.Close()
                Trace.Flush()
            End Try
        End Using

        Trace.WriteLine(Date.Now.ToString() & ": Report Task completed successfully.")
        Trace.Flush()

        Return ErrorCode.EN_ERROR_CODE.ERROR_CODE_SUCCESS
    End Function

    Public Function GetDataInMembersSheet(olecon As System.Data.OleDb.OleDbConnection, oDBClass As ADReportsDAO, InsertDatetime As String) As Boolean
        Dim Sheet1 As String = GetSheetName(2)

        Using olecomm As New OleDbCommand
            If Not GetExcelRowsOrderForMember(olecon, Sheet1) Then
                Return False
            End If

            ' Get the first row, title Row. 
            Dim sCommandTe As String = String.Format("Select {0} From [{1}] ", "* ", Sheet1)

            olecomm.CommandText = sCommandTe
            olecomm.Connection = olecon
            Dim sDomain As String = ""
            Dim nRowCount As Integer = 0
            Dim bGetDomain As Boolean = False
            Dim bTitleBarRow As Boolean = False

            Dim reader As OleDbDataReader = olecomm.ExecuteReader()
            Dim colname As String = String.Empty
            Dim MembershipType As IList(Of String) = New List(Of String)
            Dim AccountName As IList(Of String) = New List(Of String)
            Dim DisplayName As IList(Of String) = New List(Of String)
            Dim AccountType As IList(Of String) = New List(Of String)
            Dim AccountSID As IList(Of String) = New List(Of String)
            Dim FromGroup As IList(Of String) = New List(Of String)
            Dim UserIDs As IList(Of String) = New List(Of String)

            While reader.Read
                If bTitleBarRow Then  ' Go to the next line directly. 
                    bTitleBarRow = False
                    Continue While
                End If

                bGetDomain = False

                'For i As Integer = 0 To reader.FieldCount - 1
                If reader.IsDBNull(0) AndAlso reader.IsDBNull(1) AndAlso
                    reader.IsDBNull(2) AndAlso reader.IsDBNull(3) Then
                    Continue While
                End If

                Try
                    Dim i As Integer = 0
                    colname = reader.GetString(i).Trim.ToUpper
                    If i = 0 AndAlso (colname.StartsWith(DomainNameStarter) OrElse
                                      colname.StartsWith(DomainNameStartEM)) Then
                        bGetDomain = True
                        bTitleBarRow = True

                        ' Bulk Insert to db
                        If MembershipType.Count > 0 Then
                            InsertCollectionsToDB(oDBClass, sDomain, MembershipType, AccountName, DisplayName, AccountType,
                                               AccountSID, FromGroup, InsertDatetime, UserIDs)
                        End If

                        ClearTheCollections(MembershipType, AccountName, DisplayName, AccountType,
                                            AccountSID, FromGroup, UserIDs)
                        sDomain = colname
                        nRowCount = 0
                        Continue While
                    End If

                    'Next
                    ' extract the actually data.
                    Dim sAccName As String = If(Not reader.IsDBNull(MembersColIndex(Values.ColMembersAccountName)),
                                                reader.GetString(MembersColIndex(Values.ColMembersAccountName)).Trim,
                                                "")
                    MembershipType.Add(If(Not reader.IsDBNull(MembersColIndex(Values.ColMembersMembershipType)),
                                          reader.GetString(MembersColIndex(Values.ColMembersMembershipType)), ""))
                    AccountName.Add(sAccName)
                    DisplayName.Add(If(Not reader.IsDBNull(MembersColIndex(Values.ColMembersDisplayName)),
                                       reader.GetString(MembersColIndex(Values.ColMembersDisplayName)), ""))
                    AccountType.Add(If(Not reader.IsDBNull(MembersColIndex(Values.ColMembersAccountType)),
                                       reader.GetString(MembersColIndex(Values.ColMembersAccountType)), ""))
                    AccountSID.Add(If(Not reader.IsDBNull(MembersColIndex(Values.ColMembersAccountSID)),
                                      reader.GetString(MembersColIndex(Values.ColMembersAccountSID)), ""))
                    FromGroup.Add(If(Not reader.IsDBNull(MembersColIndex(Values.ColMembersFromGroup)),
                        reader.GetString(MembersColIndex(Values.ColMembersFromGroup)), ""))

                    If Not sAccName.Trim = "" Then
                        If sAccName.IndexOf("\"c) <> -1  Then
                            Dim user_id As String = sAccName.Substring(sAccName.IndexOf("\"c) + 1)
                            'If user_id.ToUpper.EndsWith("_ADM") Then
                            '    user_id = user_id.Substring(0, user_id.Length - 4)
                            'End If
                            UserIDs.Add(user_id)
                        Else
                            UserIDs.Add("")
                        End If
                    Else
                        UserIDs.Add("")
                    End If
                    nRowCount = nRowCount + 1


                    If nRowCount >= Values.ReadRowMaximum Then
                        If MembershipType.Count = 0 Then
                            nRowCount = 0
                            Continue While
                        End If

                        ' Write those ReadRowMaximum records to DB. 
                        InsertCollectionsToDB(oDBClass, sDomain, MembershipType, AccountName, DisplayName, AccountType,
                                               AccountSID, FromGroup, InsertDatetime, UserIDs)
                        nRowCount = 0
                        ClearTheCollections(MembershipType, AccountName, DisplayName, AccountType,
                                    AccountSID, FromGroup, UserIDs)
                    End If

                    'UserIDs.Add(reader.GetString(MembersColIndex(Values.ColMembersMembershipType)))
                Catch ex As Exception  ' Have blank cells on this row.
                    If reader.FieldCount > 0 Then
                        Trace.WriteLine(Date.Now.ToString() & ": " & ex.Message)
                        Trace.WriteLine("Application has ignored this below row:")
                        Dim strRow As String = String.Empty
                        For j As Integer = 0 To reader.FieldCount - 1
                            If Not reader.IsDBNull(j) Then
                                strRow = strRow & reader.GetString(j) & ","
                            Else
                                strRow = strRow & " ,"
                            End If
                        Next
                        Trace.WriteLine(strRow.TrimEnd(","c))
                    End If
                End Try
            End While

            ' Check the Last set of Rows, And write to DB. 
            If MembershipType.Count > 0 Then
                InsertCollectionsToDB(oDBClass, sDomain, MembershipType, AccountName, DisplayName, AccountType,
                                   AccountSID, FromGroup, InsertDatetime, UserIDs)
            End If
            nRowCount = 0
            reader.Close()

        End Using

        Return True
    End Function

    Public Function GetDataInReportSheet(olecon As System.Data.OleDb.OleDbConnection, oDBClass As ADReportsDAO, InsertDatetime As String) As Boolean
        Dim Sheet1 As String = GetSheetName(1)

        Using olecomm As New OleDbCommand
            Dim nTitleRowNum As Integer = GetExcelRowsOrderForReport(olecon, Sheet1)
            If 0 = nTitleRowNum Then
                Return False
            End If

            ' Get the first row, title Row. 
            Dim sCommandTe As String = String.Format("Select {0} From [{1}]  ", "* ", Sheet1)

            olecomm.CommandText = sCommandTe
            olecomm.Connection = olecon
            Dim nRowCount As Integer = 0

            Dim reader As OleDbDataReader = olecomm.ExecuteReader()
            Dim colname As String = String.Empty
            Dim lObjectPath As New List(Of String)
            Dim lObjectType As New List(Of String)
            Dim lAllowDeny As New List(Of String)
            Dim lDisplayName As New List(Of String)
            Dim lAccountName As New List(Of String)
            Dim lAccountType As New List(Of String)
            Dim lFromGroup As New List(Of String)
            Dim lApplyTo As New List(Of String)
            Dim lPermissions As New List(Of String)
            Dim lApplyDirectOnly As New List(Of String)
            Dim UserIDs As New List(Of String)
            Dim nBeginRow As Integer = 0

            While reader.Read

                ' First few lines are Before the Title Bar line.
                If nBeginRow + 1 <= nTitleRowNum Then
                    nBeginRow += 1
                    Continue While
                End If

                'For i As Integer = 0 To reader.FieldCount - 1
                If reader.IsDBNull(0) AndAlso reader.IsDBNull(1) AndAlso
                    reader.IsDBNull(2) AndAlso reader.IsDBNull(3) Then
                    Continue While
                End If

                nRowCount += 1
                Try
                    'Next
                    ' extract the actually data.
                    Dim sAccName As String = reader.GetString(ReportsColIndex(Values.ColReportAccountName)).Trim
                    lObjectPath.Add(reader.GetString(ReportsColIndex(Values.ColReportObjectPath)))
                    lObjectType.Add(reader.GetString(ReportsColIndex(Values.ColReportObjectType)))
                    lAllowDeny.Add(reader.GetString(ReportsColIndex(Values.ColReportAllowDeny)))
                    lDisplayName.Add(reader.GetString(ReportsColIndex(Values.ColReportDisplayName)))
                    lAccountName.Add(sAccName)
                    lAccountType.Add(reader.GetString(ReportsColIndex(Values.ColReportAccountType)))
                    lFromGroup.Add(If(reader.IsDBNull(ReportsColIndex(Values.ColReportFromGroup)), "",
                                      reader.GetString(ReportsColIndex(Values.ColReportFromGroup))))
                    lApplyTo.Add(reader.GetString(ReportsColIndex(Values.ColReportApplyTo)))
                    lApplyDirectOnly.Add(reader.GetString(ReportsColIndex(Values.ColReportApplyDirectOnly)))
                    lPermissions.Add(reader.GetString(ReportsColIndex(Values.ColReportPermissions)))

                    If Not sAccName = "" Then
                        If sAccName.IndexOf("\"c) <> -1  Then
                            Dim user_id As String = sAccName.Substring(sAccName.IndexOf("\"c) + 1)

                            'If user_id.ToUpper.EndsWith("_ADM") Then
                            '    user_id = user_id.Substring(0, user_id.Length - 4)
                            'End If
                            UserIDs.Add(user_id)
                        Else
                            UserIDs.Add("")
                        End If
                    Else
                        UserIDs.Add("")
                    End If

                    If nRowCount >= Values.ReadRowMaximum Then

                        ' Bulk Insert to db
                        If lObjectPath.Count > 0 Then
                            oDBClass.InsertBulkRecordsToADReports(lObjectPath.Count,
                                                     lObjectPath.ToArray,
                                                     lObjectType.ToArray,
                                                     lAllowDeny.ToArray,
                                                     lDisplayName.ToArray,
                                                     lAccountName.ToArray,
                                                     lAccountType.ToArray,
                                                     lFromGroup.ToArray,
                                                     lApplyTo.ToArray,
                                                     lApplyDirectOnly.ToArray,
                                                     lPermissions.ToArray,
                                                     InsertDatetime,
                                                     UserIDs.ToArray())
                        End If

                        ClearTheReportCollections(lObjectPath,
                                                     lObjectType,
                                                     lAllowDeny,
                                                     lDisplayName,
                                                     lAccountName,
                                                     lAccountType,
                                                     lFromGroup,
                                                     lApplyTo,
                                                     lApplyDirectOnly,
                                                     lPermissions,
                                                     UserIDs)
                        nRowCount = 0
                        Continue While
                    End If

                Catch ex As Exception  ' Have blank cells on this row.
                    If reader.FieldCount > 0 Then
                        Trace.WriteLine(Date.Now.ToString() & ": " & ex.Message)
                        Trace.WriteLine("Application has ignored this below row:")
                        Dim strRow As String = String.Empty
                        For j As Integer = 0 To reader.FieldCount - 1
                            If Not reader.IsDBNull(j) Then
                                strRow = strRow & reader.GetString(j) & ","
                            Else
                                strRow = strRow & " ,"
                            End If
                        Next
                        Trace.WriteLine(strRow.TrimEnd(","c))
                    End If
                End Try
            End While

            reader.Close()
            ' Check the Last set of Rows, write to DB. 
            ' Bulk Insert to db
            If lObjectPath.Count > 0 Then
                oDBClass.InsertBulkRecordsToADReports(lObjectPath.Count,
                                         lObjectPath.ToArray,
                                         lObjectType.ToArray,
                                         lAllowDeny.ToArray,
                                         lDisplayName.ToArray,
                                         lAccountName.ToArray,
                                         lAccountType.ToArray,
                                         lFromGroup.ToArray,
                                         lApplyTo.ToArray,
                                         lApplyDirectOnly.ToArray,
                                         lPermissions.ToArray,
                                         InsertDatetime,
                                         UserIDs.ToArray())
            End If
            nRowCount = 0

        End Using

        Return True
    End Function

    ' Check if the excel file contains all the required columns headers. 
    Function IsTileRowsValid() As Boolean
        If MembersColIndex.Count < Values.ColMembersMax Then
            Return False
        End If

        Dim allHeaders As String() = {Values.ColMembersMembershipType,
                              Values.ColMembersAccountName,
                               Values.ColMembersDisplayName,
                               Values.ColMembersAccountType,
                               Values.ColMembersAccountSID,
                               Values.ColMembersFromGroup
                              }
        For Each sHeader In allHeaders
            If Not MembersColIndex.Keys.Contains(sHeader) Then
                Return False
            End If
        Next

        Return True
    End Function

    ' Clear all the data IList. 
    Sub ClearTheReportCollections(lObjectPath As IList(Of String),
                                                 lObjectType As IList(Of String),
                                                 lAllowDeny As IList(Of String),
                                                 lDisplayName As IList(Of String),
                                                 lAccountName As IList(Of String),
                                                 lAccountType As IList(Of String),
                                                 lFromGroup As IList(Of String),
                                                 lApplyTo As IList(Of String),
                                                 lApplyDirectOnly As IList(Of String),
                                                 lPermissions As IList(Of String),
                                                 UserIDs As IList(Of String))
        lObjectPath.Clear()
        lObjectType.Clear()
        lAllowDeny.Clear()
        lDisplayName.Clear()
        lAccountName.Clear()
        lAccountType.Clear()
        lFromGroup.Clear()
        lApplyTo.Clear()
        lApplyDirectOnly.Clear()
        lPermissions.Clear()
        UserIDs.Clear()
    End Sub

    Sub ClearTheCollections(MembershipType As IList(Of String),
         AccountName As IList(Of String),
         DisplayName As IList(Of String),
         AccountType As IList(Of String),
         AccountSID As IList(Of String),
         FromGroup As IList(Of String),
         userIDs As IList(Of String))
        MembershipType.Clear()
        AccountName.Clear()
        DisplayName.Clear()
        AccountType.Clear()
        AccountSID.Clear()
        FromGroup.Clear()
        userIDs.Clear()
    End Sub

    Sub InsertCollectionsToDB(oDBClass As ADReportsDAO, sDomain As String, MembershipType As IList(Of String),
        AccountName As IList(Of String),
        DisplayName As IList(Of String),
        AccountType As IList(Of String),
        AccountSID As IList(Of String),
        FromGroup As IList(Of String),
        InsertDatetime As String,
        UserIDs As IList(Of String))
        oDBClass.InsertBulkRecordsToMembers(MembershipType.Count, MembershipType.ToArray,
                                        AccountName.ToArray, DisplayName.ToArray, AccountType.ToArray,
                                        AccountSID.ToArray, FromGroup.ToArray, sDomain, InsertDatetime,
                                        UserIDs.ToArray)
    End Sub



    ' if the head five records are null, then ignore the row.
    Private Function CheckValidRows(a As Object, b As Object, c As Object,
                                    d As Object, e As Object) As Boolean
        If a Is Nothing AndAlso b Is Nothing AndAlso c Is Nothing _
            AndAlso d Is Nothing AndAlso e Is Nothing Then
            Return False
        End If

        Return True
    End Function
End Class
