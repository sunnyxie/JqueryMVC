Option Strict Off
Imports System.Text.RegularExpressions
Public Class TextFileProcessing
    Friend Shared mDBClass As ADReportsDAO
    Friend Shared mUserListHash As Hashtable

    Friend Shared Function CsvFilesProcessing(ByRef configDO As CConfigDO, DBConn As ADReportsDAO, _
                                       DBInsertDatetime As DateTime, ByRef sErrMsg As String) As Integer
        Dim oExcel As clsExcel = New clsExcel
        Dim bHasError As Boolean = False

        If Not My.Computer.FileSystem.FileExists(configDO.ADInfoUserListFile) Then

            Trace.WriteLine("Error: " & "The .csv File:" & System.IO.Path.GetFileName(configDO.ADInfoUserListFile) _
                            & " you provided does not exist.", "ADReportsLoading")
            Return ErrorCode.EN_ERROR_CODE.ERROR_FILE_NOT_EXISTS
        End If

        If Not My.Computer.FileSystem.FileExists(configDO.ADInfoUserMembersFile) Then

            Trace.WriteLine("Error: " & "The .csv File:" & System.IO.Path.GetFileName(configDO.ADInfoUserMembersFile) _
                            & " you provided does not exist.", "ADReportsLoading")
            Return ErrorCode.EN_ERROR_CODE.ERROR_FILE_NOT_EXISTS
        End If
        'Return oExcel.GetAllDataToDB(fileName, DBConn, DBInsertDatetime.ToString("yyyy-MM-dd HH:mm:ss"), sErrMsg)

        mUserListHash = ReadUserListsCSV(configDO.ADInfoUserListFile)
        If mUserListHash Is Nothing Then
            Return ErrorCode.EN_ERROR_CODE.ERROR_IN_GET_USER_IDS
        End If

        mDBClass = DBConn
        Try
            ADReportsDAO.BeginTransaction()
            ReadMembersCSV(DBInsertDatetime, configDO.ADInfoUserMembersFile)
            ADReportsDAO.DBCommit()
        Catch ex As Exception
            ADReportsDAO.DBRollback()
            bHasError = True

            sErrMsg = ex.Message.ToString
            Trace.WriteLine("Error:TextFileProcessing.ReadMembersCSV: " & ex.Source & " : " & ex.Message.ToString, "ADReportsLoading")
            Return ErrorCode.EN_ERROR_CODE.ERROR_IN_WRITE_CVSDATA_TO_DB
        Finally
            ADReportsDAO.EndTransaction()
            Trace.Flush()
        End Try

        Trace.WriteLine(Date.Now.ToString() & ": CSV Files processing is completed successfully.")

        Return ErrorCode.EN_ERROR_CODE.ERROR_CODE_SUCCESS
    End Function

    ' Read the .csv file, and return all data in a Hashtable
    Shared Function ReadUserListsCSV(ByVal filename As String) As Hashtable
        Dim items As Array

        Try
            'Read the .csv file
            items = (From line In IO.File.ReadAllLines(filename) _
            Select Array.ConvertAll(Regex.Split(line, ",(?=(?:[^""]*""[^""]*"")*[^""]*$)"), Function(v) _
            v.ToString.TrimStart(""" ".ToCharArray).TrimEnd(""" ".ToCharArray))).ToArray

        Catch
            MessageBox.Show(String.Format("We could not open the file {0}, please make sure the file is closed!", filename))
            Trace.WriteLine(String.Format("Error:We could not open the file {0}, please make sure the file is closed!{1}", filename, vbCrLf))
            Return Nothing
        End Try

        'Dim dtData As New DataTable
        'For x As Integer = 0 To items(0).GetUpperBound(0)
        '    dtData.Columns.Add()
        'Next

        Dim resHash As New Hashtable
        For Each a In items

            If a.length < 2 Then
                Trace.WriteLine(String.Format("Warning:Length is unusual, something unnormal happened at file: {0} below line.", filename))
                For Each word As String In a
                    Trace.Write(word & ", ")
                Next
                'Write a new line character. 
                Trace.Write(String.Format("{0}", vbCrLf))
                Trace.Flush()

                Continue For
            ElseIf a(0).ToString.ToUpper.Trim = "NAME" Then
                Continue For
            ElseIf resHash.ContainsKey(a(0)) Then
                Trace.WriteLine(String.Format("The user DisplayName: '{0}' showed twice or more in the file '{1}'.", a(0), IO.Path.GetFileName(filename)))
                Continue For
            End If

            Dim AtIndex As Integer = a(1).ToString.IndexOf("@"c)
            If AtIndex = 0 OrElse AtIndex = -1 Then
                resHash.Add(a(0), "")
            Else
                resHash.Add(a(0), a(1).ToString.Substring(0, AtIndex))
            End If
        Next

        Return resHash
    End Function

    Shared Function ReadMembersCSV(DBInsertDatetime As DateTime, ByVal filename As String) As Boolean
        Dim items As Array

        Try
            'Read the .csv file
            items = (From line In IO.File.ReadAllLines(filename) _
            Select Array.ConvertAll(Regex.Split(line, ",(?=(?:[^""]*""[^""]*"")*[^""]*$)"), Function(v) _
            v.ToString.TrimStart(""" ".ToCharArray).TrimEnd(""" ".ToCharArray))).ToArray

        Catch
            MessageBox.Show(String.Format("We could not open the file {0}, please make sure the file is closed!", filename))
            Trace.WriteLine(String.Format("We could not open the file {0}, please make sure the file is closed!", filename))
            Return Nothing
        End Try

        'Dim dtData As New DataTable
        'For x As Integer = 0 To items(0).GetUpperBound(0)
        '    dtData.Columns.Add()
        'Next

        ' Read file Line by line.
        For Each a In items

            If a.length < 2 Then
                Trace.Write(String.Format("Length is unusual, something unnormal happened at file: {0} below line.{1}", filename, vbCrLf))
                For Each word As String In a
                    Trace.Write(word & ", ")
                Next
                'Write a new line character. 
                Trace.Write(String.Format("{0}", vbCrLf))
                Trace.Flush()

                Continue For
            ElseIf a(0).ToString.ToUpper.Trim = "NAME" Then
                Continue For
            End If

            Dim members As String() = Array.ConvertAll(a(1).ToString.Split(New Char() {";"c}), Function(v) _
            v.ToString.Trim())
            Dim arLen As Integer = members.Length
            Dim arUserIds(arLen - 1) As String
            Dim arMemberType(arLen - 1) As String
            Dim arAccountName(arLen - 1) As String
            Dim arAccountType(arLen - 1) As String
            Dim arSID(arLen - 1) As String

            Dim tmpID As String
            For i As Integer = 0 To arLen - 1
                tmpID = mUserListHash(members(i))
                If tmpID IsNot Nothing Then
                    arUserIds(i) = tmpID.ToString
                Else
                    arUserIds(i) = String.Empty
                End If

                arMemberType(i) = String.Empty
                arAccountName(i) = String.Empty
                arAccountType(i) = String.Empty
                arSID(i) = String.Empty
            Next
            ' Insert one Group at a time
            mDBClass.InsertBulkRecordsToMemTbl(arLen, a(0).ToString,
                                     "", arMemberType,
                                     arAccountName, members,
                                     arAccountType, arSID,
                                     DBInsertDatetime.ToString("yyyy-MM-dd HH:mm:ss"),
                                     arUserIds)
        Next

        Return True
    End Function
End Class
