Option Strict Off
Imports CanadianNatural.ApplicationServices.Security.Interface
Imports CanadianNatural.ApplicationServices.Security.AppSecure
Imports Oracle.DataAccess.Types
Imports Oracle.DataAccess.Client

Public Class clsDBAccess
    Private mCurrentUser As ICNRLUser = Nothing
    Private Shared DataRepositoryCon As OracleConnection
    Private Shared moDbTransaction As OracleTransaction = Nothing

    Private AS_URL As String
    Private AS_Environment As String
    Private AS_AppID As String
    Private AS_ConnectionName As String

    ' Constructor
    Sub New()
        AS_URL = My.Settings.SecurityConnectionInformation
        AS_Environment = My.Settings.SecurityEnvironmentID
        AS_AppID = My.Settings.SecurityApplicationID
        AS_ConnectionName = Values.AppSecureConnString
    End Sub

    Public Property AppSecureURL() As String
        Get
            Return AS_URL
        End Get
        Set(ByVal value As String)
            AS_URL = value
        End Set
    End Property

    Public Property AppSecureEnvironment() As String
        Get
            Return AS_Environment
        End Get
        Set(ByVal value As String)
            AS_Environment = value
        End Set
    End Property

    Public Property AppSecureAppID() As String
        Get
            Return AS_AppID
        End Get
        Set(ByVal value As String)
            AS_AppID = value
        End Set
    End Property

    Public Property AppSecureConnectionName() As String
        Get
            Return AS_ConnectionName
        End Get
        Set(ByVal value As String)
            AS_ConnectionName = value
        End Set
    End Property

    '''<summary> 
    '''Method to connect to a database using AppSecure.
    '''Requires AppSecureServerURL, AppSecureEnvironment and AppSecureApplicationID settings in app.config
    '''</summary> 
    '''<returns>True if successful else False</returns> 
    Public Function ConnectToDataRepository() As Boolean
        Dim authenticationProvider As ICNRLAuthenticationProvider = New SecurityProvider
        ' Login using the app.config settings
        Dim sErrorMessage As String = ""

        ConnectToDataRepository = False
        Try
            mCurrentUser = authenticationProvider.Login(ICNRLAuthenticationProvider.UserLoginType.Automatic, AppSecureURL, AppSecureEnvironment, AppSecureAppID, "", "", sErrorMessage)
            If mCurrentUser IsNot Nothing Then
                'DataRepositoryCon = New OracleConnection(mCurrentUser.Applicaton.GetDBConnectionString(AppSecureConnectionName)) ' PRIMARY
                'DataRepositoryCon = New OracleConnection(AppSecureConnectionName) ' PRIMARY
                DataRepositoryCon = New OracleConnection() ' PRIMARY
                mCurrentUser.Applicaton.OpenDBConnection(AppSecureConnectionName, DataRepositoryCon)

                ConnectToDataRepository = True
            Else
                Throw New Exception("Unable to create an instance of ICNRLUser")
            End If
        Catch ex As Exception
            Throw New Exception("clsDBAccess.ConnectToDataRepository. Error:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to disconnect from a database.
    '''</summary> 
    '''<returns>True if successful else False</returns> 
    Public Function DisconnectFromDataRepository() As Boolean
        If DataRepositoryCon Is Nothing Then
            Return True
        End If

        Dim sErrorMessage As String = ""

        DisconnectFromDataRepository = False
        Try
            If DataRepositoryCon.State = ConnectionState.Open Then
                DataRepositoryCon.Close()
                DisconnectFromDataRepository = True
            End If
        Catch ex As Exception
            Throw New Exception("clsDBAccess.DisconnectFromDataRepository. Error:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to check user's privileges using AppSecure
    '''</summary> 
    '''<param name="Operation">AppSecure Operation to check</param> 
    '''<returns>True if user has requested privilege else False</returns> 
    Public Function CanPerformOperation(ByVal Operation As String) As Boolean
        CanPerformOperation = False
        Try
            If mCurrentUser IsNot Nothing Then
                CanPerformOperation = mCurrentUser.CanPerformOperation(Operation)
            End If
        Catch ex As Exception
            Throw New Exception("clsDBAccess.CanPerformOperation. Error:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to check user's privileges using AppSecure
    '''</summary> 
    '''<param name="OperationName">AppSecure Operation Name to check</param> 
    '''<returns>True if user has requested privilege else False</returns> 
    Public Function CanPerformOperationByName(ByVal OperationName As String) As Boolean
        CanPerformOperationByName = False
        Try
            If mCurrentUser IsNot Nothing Then
                CanPerformOperationByName = mCurrentUser.CanPerformOperationByName(OperationName)
            End If
        Catch ex As Exception
            Throw New Exception("clsDBAccess.CanPerformOperationByName. Error:" & ex.Message)
        End Try
    End Function

    Protected Overrides Sub Finalize()
        Try
            If DataRepositoryCon IsNot Nothing Then
                If DataRepositoryCon.State = ConnectionState.Open Then
                    DataRepositoryCon.Close()
                    DataRepositoryCon = Nothing
                    If mCurrentUser IsNot Nothing Then
                        mCurrentUser = Nothing
                    End If
                End If
            End If
        Catch ex As Exception
            ' do nothing
        Finally
            MyBase.Finalize()
        End Try
    End Sub

    Public Overrides Function ToString() As String
        If DataRepositoryCon Is Nothing Then
            Return MyBase.ToString()
        Else
            Return DataRepositoryCon.ConnectionString
        End If
    End Function

#Region " DB Transaction "

    Public Shared Sub BeginTransaction()
        moDbTransaction = DataRepositoryCon.BeginTransaction
    End Sub

    Public Shared Sub EndTransaction()
        moDbTransaction = Nothing
    End Sub

    Public Shared Sub DBCommit()
        moDbTransaction.Commit()
    End Sub

    Public Shared Sub DBRollback()
        moDbTransaction.Rollback()
    End Sub

#End Region
    Protected Function CreateCommandBase(ByVal storedProcName As String) As Oracle.DataAccess.Client.OracleCommand
        Dim command As Oracle.DataAccess.Client.OracleCommand = New Oracle.DataAccess.Client.OracleCommand(storedProcName)
        command.CommandType = CommandType.StoredProcedure
        Return command
    End Function

    ' If stored procedure did not return datasets ?
    Protected Function CreateDbCommand(ByVal storedProcName As String) As OracleCommand
        Dim loDBCommand As OracleCommand = CreateCommandBase(storedProcName)
        loDBCommand.Connection = DataRepositoryCon

        If (moDbTransaction IsNot Nothing) Then
            loDBCommand.Transaction = moDbTransaction
        End If

        Return loDBCommand
    End Function

    Private Sub AddedOutputParamater(ByVal voDBCommand As OracleCommand, _
                                     ByVal vsParameterName As String, _
                                     ByVal voOracleType As Oracle.DataAccess.Client.OracleDbType, _
                                     ByVal vsLength As Integer)
        voDBCommand.Parameters.AddRange(New Oracle.DataAccess.Client.OracleParameter() { _
                New Oracle.DataAccess.Client.OracleParameter(vsParameterName, voOracleType, vsLength, _
                                                             System.Data.ParameterDirection.Output, True, CType(0, Byte), _
                                                             CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing)})

    End Sub

    Protected Sub AddedOutputRefCursor(ByVal voDBCommand As OracleCommand, _
                                       ByVal vsParameterName As String)
        AddedOutputParamater(voDBCommand, vsParameterName, OracleDbType.RefCursor, 0)
    End Sub

    Protected Sub AddedOutputVarChar(ByVal voDBCommand As OracleCommand, _
                                     ByVal vsParameterName As String)
        AddedOutputParamater(voDBCommand, vsParameterName, OracleDbType.Varchar2, 1024)
    End Sub

    Protected Sub AddedOutputVarChar(ByVal voDBCommand As OracleCommand, _
                                     ByVal vsParameterName As String, _
                                     ByVal vsLength As Integer)
        AddedOutputParamater(voDBCommand, vsParameterName, OracleDbType.Varchar2, vsLength)
    End Sub


    Protected Sub AddedOutputNumber(ByVal voDBCommand As OracleCommand, _
                                     ByVal vsParameterName As String)
        AddedOutputParamater(voDBCommand, vsParameterName, OracleDbType.Double, 100)
    End Sub

    Private Sub AddedInputParamater(ByVal voDBCommand As OracleCommand, _
                                    ByVal vsParameterName As String, _
                                    ByVal voValue As Object, _
                                    ByVal voOracleType As Oracle.DataAccess.Client.OracleDbType)

        voDBCommand.Parameters.AddRange(New Oracle.DataAccess.Client.OracleParameter() { _
         New Oracle.DataAccess.Client.OracleParameter(vsParameterName, voOracleType)})
        voDBCommand.Parameters(vsParameterName).Value = voValue

    End Sub

    Protected Sub AddedInputVarChar(ByVal voDBCommand As OracleCommand, _
                                    ByVal vsParameterName As String, ByVal vsValue As String)

        If (vsValue Is Nothing) Then
            vsValue = ""
        End If

        AddedInputParamater(voDBCommand, vsParameterName, vsValue, OracleDbType.Varchar2)
    End Sub

    Protected Sub AddedInputNumber(ByVal voDBCommand As OracleCommand, _
                                   ByVal vsParameterName As String, ByVal voValue As Object)
        AddedInputParamater(voDBCommand, vsParameterName, voValue, OracleDbType.Double)
    End Sub

    Protected Sub AddedInputDate(ByVal voDBCommand As OracleCommand, _
                                 ByVal vsParameterName As String, ByVal vsValue As String)
        AddedInputParamater(voDBCommand, vsParameterName, vsValue, OracleDbType.Date)
    End Sub

#Region "Get and update data"
    Protected Function GetDataViaDataTable(ByVal voDBCommand As OracleCommand) As DataTable
        Dim loDataTable As DataTable = Nothing

        ' Add Parameters
        AddedOutputNumber(voDBCommand, "ret_status")
        AddedOutputVarChar(voDBCommand, "ret_msg")
        AddedOutputRefCursor(voDBCommand, "inret_list")

        voDBCommand.ExecuteNonQuery()

        If (voDBCommand.Parameters("ret_status").Value = 0) Then
            loDataTable = New DataTable
            loDataTable.Load(voDBCommand.Parameters("inret_list").Value)
            Return loDataTable
        Else
            Throw New Exception(voDBCommand.Parameters("ret_msg").Value)
        End If

    End Function

    Protected Function UpdateDataViaDBCommand(ByVal voDBCommand As OracleCommand) As Boolean

        ' Add Parameters
        AddedOutputNumber(voDBCommand, "ret_status")
        AddedOutputVarChar(voDBCommand, "ret_msg")

        voDBCommand.ExecuteNonQuery()

        If (Not IsDBNull(voDBCommand.Parameters("ret_status").Value)) Then

            If (voDBCommand.Parameters("ret_status").Value = 0) Then
                Return True
            Else
                Throw New Exception(voDBCommand.Parameters("ret_msg").Value & ".")
            End If
        Else
            Return True
        End If

    End Function

    Protected Function DeleteDataViaDBCommand(ByVal voDBCommand As OracleCommand) As Boolean

        ' Add Parameters
        AddedOutputNumber(voDBCommand, "ret_status")
        AddedOutputVarChar(voDBCommand, "ret_msg")

        voDBCommand.ExecuteNonQuery()

        If (Not IsDBNull(voDBCommand.Parameters("ret_status").Value)) Then

            If (voDBCommand.Parameters("ret_status").Value = 0) Then
                Return True
            Else
                Return False 'Throw New Exception(voDBCommand.Parameters("ret_msg").Value)
            End If
        Else
            Return False
        End If
        Return False
    End Function

    Protected Function GetDataViaPackage(ByVal vsProcedureName As String) As DataTable
        Dim loDBCommand As OracleCommand = CreateDbCommand(vsProcedureName)
        Return GetDataViaDataTable(loDBCommand)
    End Function

    Protected Sub UpdateDataViaPackage(ByVal vsProcedureName As String)
        Dim loDBCommand As OracleCommand = CreateDbCommand(vsProcedureName)
        UpdateDataViaDBCommand(loDBCommand)
    End Sub
#End Region
End Class
