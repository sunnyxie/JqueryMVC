Imports OakLeaf.MM.Main.Business

Public Class AppDataAccessBase(Of EntityType As {CNRLBusinessEntity, New})
    Inherits CNRLDataAccessOracle(Of EntityType)

#Region "Constructor/Destructor"

    Public Sub New(ByVal ownerObject As mmBusinessObject)
        MyBase.New(Constants.Values.ApplicationKey, ownerObject)
    End Sub


#End Region

#Region "Properties"

#End Region

#Region "Public Methods"
    ' Created this function as the framework has changed how CreateConnection worked in previous versions.
    Public Function CreateNewConnection() As System.Data.IDbConnection
        Dim newConnection As System.Data.IDbConnection = New System.Data.OracleClient.OracleConnection
        CNRLObjectFactory.SecurityFactory(DatabaseKey).CurrentUser.Applicaton.OpenDBConnection(DatabaseSetKey, newConnection)
        Return newConnection
    End Function

    Public Function ManuallyExecuteRecord(ByVal pobjCommand As System.Data.OracleClient.OracleCommand, Optional ByRef pobjTxn As System.Data.IDbTransaction = Nothing) As Boolean
        Dim blnResult As Boolean = True
        Dim blnOpen As Boolean = False
        Dim objConnection As System.Data.IDbConnection = Nothing

        ' Determine which connection object to use.
        If pobjTxn Is Nothing Then
            objConnection = CreateNewConnection()
            blnOpen = True
            pobjCommand.Connection = objConnection
        Else
            pobjCommand.Connection = pobjTxn.Connection
            pobjCommand.Transaction = pobjTxn
        End If

        ' Execute the command.
        Try
            pobjCommand.ExecuteNonQuery()
        Catch ex As Exception
            blnResult = False
        Finally
            If blnOpen Then objConnection.Close()
        End Try

        Return blnResult
    End Function

    Public Function ManuallyGetValue(ByVal pobjCommand As System.Data.OracleClient.OracleCommand) As String
        Dim strValue As String = ""
        Dim objConnection As System.Data.IDbConnection

        ' Open a db connection to use.
        objConnection = CreateNewConnection()
        pobjCommand.Connection = objConnection

        ' Execute the command.
        Try
            strValue = CType(pobjCommand.ExecuteScalar(), String)
        Catch ex As Exception
            strValue = ""
        Finally
            objConnection.Close()
        End Try

        Return strValue
    End Function

    ' This routine would be for admin functions.  Currently not used.
    Public Function ManuallyExecuteSQL(ByVal pstrSQL As String) As Boolean
        Dim blnResult As Boolean = True
        Dim objConnection As System.Data.IDbConnection
        Dim objCommand As System.Data.OracleClient.OracleCommand = Me.CreateCommandBase("")

        ' Open a db connection to use.
        objConnection = CreateNewConnection()
        objCommand.Connection = objConnection

        ' Setup the command object.
        objCommand.CommandType = CommandType.Text
        objCommand.CommandText = pstrSQL

        ' Execute the command.
        Try
            blnResult = ManuallyExecuteRecord(objCommand)
        Catch ex As Exception
            blnResult = False
        Finally
            objConnection.Close()
        End Try

        Return blnResult
    End Function

    Public Function GetNextSequenceNo(ByVal pstrTableName As String) As System.Data.IDbCommand
        Dim oCommand As System.Data.OracleClient.OracleCommand = Me.CreateCommandBase("")
        Dim strSQL As String = ""

        strSQL = "SELECT REMEDYFM." + pstrTableName + "_SEQ.NEXTVAL"
        strSQL += "  FROM DUAL"

        oCommand.CommandType = CommandType.Text
        oCommand.CommandText = strSQL

        Return oCommand
    End Function
#End Region

#Region "Private/Protected Methods"
    Public Overrides Function CreateConnection() As IDbConnection
        Dim conn As OracleClient.OracleConnection = MyBase.CreateConnection()
        Dim cmd As OracleClient.OracleCommand = conn.CreateCommand
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "DBMS_SESSION.SET_IDENTIFIER"
        cmd.Parameters.Add(New OracleClient.OracleParameter("client_id", Constants.Values.g_real_user_id))

        cmd.ExecuteNonQuery()
        Return conn
    End Function

#End Region

End Class
