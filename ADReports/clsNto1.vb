Public Class clsNto1
    '''<summary> 
    '''Method to verify existance of a node in the TAMSHIERARCHY table.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="NodeName">Name of node to verify </param> 
    '''<param name="NodeID">ID of node if it exists </param> 
    '''<param name="ParentID">ID of parent node if it exists </param> 
    '''<returns>True if node exists else False</returns> 
    Public Function HierarchyNodeExists(ByRef Nto1Con As clsDBAccess, ByVal NodeName As String, _
                                        ByRef NodeID As Long, ByRef ParentID As Long, _
                                        ByVal Nto1LevelName As String) As Boolean
        Dim aReader As Data.OracleClient.OracleDataReader = Nothing

        HierarchyNodeExists = False
        Try
            Dim sSQL As String = "SELECT TH.ID, TH.PARENTNODE, LEVELNAME " & _
                                 "FROM TAMSHIERARCHY TH " & _
                                 "WHERE TH.NODENAME = '" & NodeName & "'"
            aReader = Nto1Con.ExecuteOracleSQLDataReader(sSQL)
            If aReader IsNot Nothing Then
                If aReader.HasRows Then
                    aReader.Read()
                    If aReader.Item("ID").ToString() <> "" Then
                        HierarchyNodeExists = True
                        NodeID = CLng(aReader.Item("ID"))
                        If aReader.Item("PARENTNODE").ToString() <> "" Then
                            ParentID = CLng(aReader.Item("PARENTNODE"))
                        End If
                        If aReader.Item("LEVELNAME").ToString() <> "" Then
                            Nto1LevelName = aReader.Item("LEVELNAME").ToString()
                        End If
                    End If
                End If
            End If
            aReader = Nothing
        Catch ex As Exception
            Throw New Exception("Error in Nto1.HierarchyNodeExists:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to get the node name for a record in the TAMSHIERARCHY table.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="NodeID">ID of node </param> 
    '''<returns>Node name if it exists else empty string</returns> 
    Public Function GetHierarchyNodeName(ByRef Nto1Con As clsDBAccess, ByVal NodeID As Long) As String
        Dim aReader As Data.OracleClient.OracleDataReader = Nothing

        GetHierarchyNodeName = ""
        Try
            Dim sSQL As String = "SELECT TH.NODENAME " & _
                                 "FROM TAMSHIERARCHY TH " & _
                                 "WHERE TH.ID = " & NodeID
            aReader = Nto1Con.ExecuteOracleSQLDataReader(sSQL)
            If aReader IsNot Nothing Then
                If aReader.HasRows Then
                    aReader.Read()
                    If aReader.Item("NODENAME").ToString() <> "" Then
                        GetHierarchyNodeName = aReader.Item("NODENAME").ToString()
                    End If
                End If
            End If
            aReader = Nothing
        Catch ex As Exception
            Throw New Exception("Error in Nto1.GetHierarchyNodeName:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to verify existance of a circuit in the TAMSINVENTORY table.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="Circuit">Circuit to verify </param> 
    '''<param name="CircuitID">ID of circuit if it exists </param> 
    '''<param name="OwnerID">ID of circuit owner if it exists </param> 
    '''<returns>True if circuit exists else False</returns> 
    Public Function CircuitExists(ByRef Nto1Con As clsDBAccess, ByVal Circuit As String, _
                                  ByRef CircuitID As Long, ByRef OwnerID As Long) As Boolean
        Dim aReader As Data.OracleClient.OracleDataReader = Nothing

        CircuitExists = False
        Try
            Dim sSQL As String = "SELECT TI.CIRCUITID, TI.OWNER " & _
                                 "FROM TAMSINVENTORY TI " & _
                                 "WHERE TI.CIRCUIT = '" & Circuit & "'"
            aReader = Nto1Con.ExecuteOracleSQLDataReader(sSQL)
            If aReader IsNot Nothing Then
                If aReader.HasRows Then
                    aReader.Read()
                    If aReader.Item("CIRCUITID").ToString() <> "" Then
                        CircuitExists = True
                        CircuitID = CLng(aReader.Item("CIRCUITID"))
                        If aReader.Item("OWNER").ToString() <> "" Then
                            OwnerID = CLng(aReader.Item("OWNER"))
                        End If
                    End If
                End If
            End If
            aReader = Nothing
        Catch ex As Exception
            Throw New Exception("Error in Nto1.CircuitExists:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to verify existance of a cost allocation record in the TAMSCOSTALLOC table for a specific Circuit ID.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="CircuitID">ID of circuit </param> 
    '''<param name="CostAllocID">Array containing ID of cost allocation records if any exist </param> 
    '''<param name="CostCenter">Array containing cost centers for records if any exist </param> 
    '''<param name="Percent">Array containing percent allocation for records if any exist </param> 
    '''<returns>True if one or more cost allocation records exists else False</returns> 
    Public Function CostAllocExists(ByRef Nto1Con As clsDBAccess, ByVal CircuitID As Long, _
                                    ByRef CostAllocID() As Long, ByRef CostCenter() As String, _
                                    ByRef Percent() As Double) As Boolean
        Dim aReader As Data.OracleClient.OracleDataReader = Nothing
        Dim i As Long = 0

        CostAllocExists = False
        Try
            Dim sSQL As String = "SELECT TC.UNIQUEID, TC.COSTCENTER, TC.PERCENT " & _
                                 "FROM TAMSCOSTALLOC TC " & _
                                 "WHERE TC.ID = " & CircuitID
            aReader = Nto1Con.ExecuteOracleSQLDataReader(sSQL)
            If aReader IsNot Nothing Then
                If aReader.HasRows Then
                    CostAllocExists = True
                    Do While aReader.Read()
                        ReDim Preserve CostAllocID(i)
                        ReDim Preserve CostCenter(i) ' this is the only one that can be NULL
                        ReDim Preserve Percent(i)
                        CostAllocID(i) = CLng(aReader.Item("UNIQUEID"))
                        CostCenter(i) = aReader.Item("COSTCENTER").ToString()
                        Percent(i) = aReader.Item("PERCENT").ToString()
                        i = i + 1
                    Loop
                End If
            End If
            aReader = Nothing
        Catch ex As Exception
            Throw New Exception("Error in Nto1.CostAllocExists:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to update the parent ID of a node in the TAMSHIERARCHY table.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="NodeID">ID of node to update </param> 
    '''<param name="ParentID">ID of parent node </param> 
    '''<returns>True if parent ID of node updated else False</returns> 
    Public Function UpdateHierarchyParentID(ByRef Nto1Con As clsDBAccess, ByVal NodeID As Long, _
                                            ByVal ParentID As Long) As Boolean
        Dim retVal As Integer = 0

        UpdateHierarchyParentID = False
        Try
            Dim sSQL As String = "UPDATE TAMSHIERARCHY " & _
                                 "SET PARENTNODE = " & ParentID & " " & _
                                 "WHERE ID = " & NodeID
            retVal = Nto1Con.ExecuteOracleSQLNonQuery(sSQL)
            If retVal > 0 Then
                UpdateHierarchyParentID = True
            End If
        Catch ex As Exception
            Throw New Exception("Error in Nto1.UpdateHierarchyParentID:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to update the LevelName of a node in the TAMSHIERARCHY table.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="NodeID">ID of node to update </param> 
    '''<param name="LevelName">Update value for LevelName </param> 
    '''<returns>True if parent ID of node updated else False</returns> 
    Public Function UpdateHierarchyLevelName(ByRef Nto1Con As clsDBAccess, ByVal NodeID As Long, _
                                    ByVal LevelName As String) As Boolean
        Dim retVal As Integer = 0

        UpdateHierarchyLevelName = False
        Try
            Dim sSQL As String = "UPDATE TAMSHIERARCHY " & _
                                 "SET LEVELNAME = '" & LevelName & "' " & _
                                 "WHERE ID = " & NodeID
            retVal = Nto1Con.ExecuteOracleSQLNonQuery(sSQL)
            If retVal > 0 Then
                UpdateHierarchyLevelName = True
            End If
        Catch ex As Exception
            Throw New Exception("Error in Nto1.UpdateHierarchyLevelName:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to insert a node in the TAMSHIERARCHY table.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="NodeName">Node name </param> 
    '''<param name="LevelName">Level Name </param> 
    '''<param name="ParentID">Parent Node ID </param> 
    '''<returns>True if node inserted else False</returns> 
    Public Function InsertHierarchyNode(ByRef Nto1Con As clsDBAccess, ByRef NodeID As Long, _
                                        ByVal NodeName As String, ByVal LevelName As String, _
                                        ByVal ParentID As Long) As Boolean
        Dim retVal As Integer = 0

        InsertHierarchyNode = False
        Try
            Dim sSQL As String = "INSERT INTO TAMSHIERARCHY (ID, NODENAME, PARENTNODE, LEVELNAME) " & _
                                 "VALUES ((select max(ID) + 1 from tamshierarchy), '" & NodeName & "', " & _
                                 ParentID & ", '" & LevelName & "')"
            retVal = Nto1Con.ExecuteOracleSQLNonQuery(sSQL)
            If retVal > 0 Then
                ' Get NodeID
                Dim aReader As Data.OracleClient.OracleDataReader = Nothing
                Dim SQL As String = "SELECT ID " & _
                                    "FROM TAMSHIERARCHY " & _
                                    "WHERE NODENAME = '" & NodeName & "'"
                aReader = Nto1Con.ExecuteOracleSQLDataReader(SQL)
                If aReader IsNot Nothing Then
                    If aReader.HasRows Then
                        aReader.Read()
                        NodeID = CLng(aReader.Item("ID"))
                    End If
                End If
                InsertHierarchyNode = True
            End If

        Catch ex As Exception
            Throw New Exception("Error in Nto1.InsertHierarchyNode:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to update the node name of a record in the TAMSHIERARCHY table.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="HierarchyID">ID of record to update </param> 
    '''<param name="OwnerName">Name of owner </param> 
    '''<returns>True if owner name updated else False</returns> 
    Public Function UpdateHierarchyOwnerName(ByRef Nto1Con As clsDBAccess, ByVal HierarchyID As Long, _
                                       ByVal OwnerName As String) As Boolean
        Dim retVal As Integer = 0

        UpdateHierarchyOwnerName = False
        Try
            Dim sSQL As String = "UPDATE TAMSHIERARCHY " & _
                                 "SET NODENAME = '" & OwnerName & "' " & _
                                 "WHERE ID = " & HierarchyID
            retVal = Nto1Con.ExecuteOracleSQLNonQuery(sSQL)
            If retVal > 0 Then
                UpdateHierarchyOwnerName = True
            End If
        Catch ex As Exception
            Throw New Exception("Error in Nto1.UpdateHierarchyOwnerName:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to update the owner ID of a circuit in the TAMSINVENTORY table.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="CircuitID">ID of record to update </param> 
    '''<param name="OwnerID">ID of owner (from TAMSHIERARCHY table)</param> 
    '''<returns>True if owner name updated else False</returns> 
    Public Function UpdateCircuitOwnerID(ByRef Nto1Con As clsDBAccess, ByVal CircuitID As Long, _
                                       ByVal OwnerID As Long) As Boolean
        Dim retVal As Integer = 0

        UpdateCircuitOwnerID = False
        Try
            Dim sSQL As String = "UPDATE TAMSINVENTORY " & _
                                 "SET OWNER = " & OwnerID & " " & _
                                 "WHERE CIRCUITID = " & CircuitID
            retVal = Nto1Con.ExecuteOracleSQLNonQuery(sSQL)
            If retVal > 0 Then
                UpdateCircuitOwnerID = True
            End If
        Catch ex As Exception
            Throw New Exception("Error in Nto1.UpdateCircuitOwnerID:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to update a TAMSINVENTORY record.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="CircuitID">ID of record to update </param> 
    '''<param name="OwnerID">ID of owner (from TAMSHIERARCHY table)</param> 
    '''<returns>True if owner name updated else False</returns> 
    Public Function UpdateTAMSInventory(ByRef Nto1Con As clsDBAccess, ByVal CircuitID As Long, _
                                       ByVal OwnerID As Long, ByVal ClassName As String, _
                                       ByVal InvStatus As String, ByVal Vendor As String) As Boolean
        Dim retVal As Integer = 0

        UpdateTAMSInventory = False
        Try
            Dim sSQL As String = "UPDATE TAMSINVENTORY " & _
                                 "SET CLASS = '" & ClassName & "', " & _
                                 "STATUS = '" & InvStatus & "', " & _
                                 "VENDOR = '" & Vendor & "', " & _
                                 "OWNER = " & OwnerID & " " & _
                                 "WHERE CIRCUITID = " & CircuitID
            retVal = Nto1Con.ExecuteOracleSQLNonQuery(sSQL)
            If retVal > 0 Then
                UpdateTAMSInventory = True
            End If
        Catch ex As Exception
            Throw New Exception("Error in Nto1.UpdateTAMSInventory:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to insert a record (circuit) in the TAMSINVENTORY table.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="Circuit">Circuit (phone number) </param> 
    '''<param name="CircuitType">Class name (type - e.g. Cell) </param> 
    '''<param name="OwnerID">ID of owner (from TAMShierarchy table)</param> 
    '''<param name="Vendor">Vendor Name </param> 
    '''<param name="Status">Status of circuit</param> 
    '''<returns>True if circuit inserted else False</returns> 
    Public Function InsertCircuit(ByRef Nto1Con As clsDBAccess, ByVal CircuitID As Long, ByVal Circuit As String, _
                                  ByVal CircuitType As String, ByVal OwnerID As Long, _
                                  ByVal Vendor As String, ByVal Status As String) As Boolean
        Dim retVal As Integer = 0

        InsertCircuit = False
        Try
            Dim sSQL As String = "INSERT INTO TAMSINVENTORY (CIRCUITID, CIRCUIT, CLASS, OWNER, VENDOR, STATUS) " & _
                                 "VALUES ((select max(CIRCUITID) + 1 from tamsinventory), '" & Circuit & "', '" & _
                                 CircuitType & "', " & OwnerID & ", '" & Vendor & "', '" & Status & "')"
            retVal = Nto1Con.ExecuteOracleSQLNonQuery(sSQL)
            If retVal > 0 Then
                ' Get CircuitID
                Dim aReader As Data.OracleClient.OracleDataReader = Nothing
                Dim SQL As String = "SELECT CIRCUITID " & _
                                    "FROM TAMSINVENTORY " & _
                                    "WHERE CIRCUIT = '" & Circuit & "'"
                aReader = Nto1Con.ExecuteOracleSQLDataReader(SQL)
                If aReader IsNot Nothing Then
                    If aReader.HasRows Then
                        aReader.Read()
                        CircuitID = CLng(aReader.Item("CIRCUITID"))
                    End If
                End If
                InsertCircuit = True
            End If
        Catch ex As Exception
            Throw New Exception("Error in Nto1.InsertCircuit:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to update the percent of a cost allocation record in the TAMSCOSTALLOC table.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="CircuitID">Circuit ID of record to update </param> 
    '''<param name="CostCenter">Cost Centre of record to update </param> 
    '''<param name="Percent">Update percent </param> 
    '''<returns>True if percent updated else False</returns> 
    Public Function UpdateCostCentrePercent(ByRef Nto1Con As clsDBAccess, ByVal CircuitID As Long, _
                                            ByVal CostCenter As String, ByVal Percent As Double) As Boolean
        Dim retVal As Integer = 0

        UpdateCostCentrePercent = False
        Try
            Dim sSQL As String = "UPDATE TAMSCOSTALLOC " & _
                                 "SET PERCENT = " & Percent & " " & _
                                 "WHERE ID = " & CircuitID & " " & _
                                 "AND COSTCENTER = '" & CostCenter & "'"
            retVal = Nto1Con.ExecuteOracleSQLNonQuery(sSQL)
            If retVal > 0 Then
                UpdateCostCentrePercent = True
            End If
        Catch ex As Exception
            Throw New Exception("Error in Nto1.UpdateCostCentrePercent:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to insert a record in the TAMSCOSTALLOC table for a circuit.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="CircuitID">ID of circuit (from TAMSinventory table) </param> 
    '''<param name="CostCenter">Cost Center (can be null) </param> 
    '''<param name="Percent">Percent allocation to the cost center</param> 
    '''<returns>True if record inserted else False</returns> 
    Public Function InsertCostAllocRecord(ByRef Nto1Con As clsDBAccess, ByVal CircuitID As Long, _
                                  ByVal CostCenter As String, ByVal Percent As Double) As Boolean
        Dim retVal As Integer = 0
        Dim sSQL As String = ""

        InsertCostAllocRecord = False
        Try
            If CostCenter <> "" Then
                sSQL = "INSERT INTO TAMSCOSTALLOC (ID, SOURCE, COSTCENTER, PERCENT, GLPLAN) " & _
                       "VALUES (" & CircuitID & ", 'I', '" & CostCenter & "', " & Percent & ", 'Default')"
            Else
                sSQL = "INSERT INTO TAMSCOSTALLOC (ID, SOURCE, COSTCENTER, PERCENT, GLPLAN) " & _
                       "VALUES (" & CircuitID & ", 'I', NULL, " & Percent & ", 'Default')"
            End If
            retVal = Nto1Con.ExecuteOracleSQLNonQuery(sSQL)
            If retVal > 0 Then
                InsertCostAllocRecord = True
            End If
        Catch ex As Exception
            Throw New Exception("Error in Nto1.InsertCostAllocRecord:" & ex.Message)
        End Try
    End Function

    '''<summary> 
    '''Method to insert a record in the TAMSCOSTALLOC table for a circuit.
    '''</summary> 
    '''<param name="Nto1Con">Connection to Nto1 database in clsDBAccess </param> 
    '''<param name="CostAllocID">ID of Allocation Record (from TAMSCostAlloc table) </param> 
    '''<param name="CostCenter">Cost Center (can be null) </param> 
    '''<param name="Percent">Percent allocation to the cost center</param> 
    '''<returns>True if record updated else False</returns> 
    Public Function UpdateCostAllocRecord(ByRef Nto1Con As clsDBAccess, ByVal CostAllocID As Long, _
                                  ByVal CostCenter As String, ByVal Percent As Double) As Boolean
        Dim retVal As Integer = 0
        Dim sSQL As String = ""

        UpdateCostAllocRecord = False
        Try
            If CostCenter <> "" Then
                sSQL = "UPDATE TAMSCOSTALLOC " & _
                       "SET COSTCENTER = '" & CostCenter & "', " & _
                       "PERCENT = " & Percent & " " & _
                       "WHERE UNIQUEID = " & CostAllocID
            Else
                sSQL = "UPDATE TAMSCOSTALLOC " & _
                       "SET COSTCENTER = NULL, " & _
                       "PERCENT = " & Percent & " " & _
                       "WHERE UNIQUEID = " & CostAllocID
            End If
            retVal = Nto1Con.ExecuteOracleSQLNonQuery(sSQL)
            If retVal > 0 Then
                UpdateCostAllocRecord = True
            End If
        Catch ex As Exception
            Throw New Exception("Error in Nto1.UpdateCostAllocRecord:" & ex.Message)
        End Try
    End Function

End Class
