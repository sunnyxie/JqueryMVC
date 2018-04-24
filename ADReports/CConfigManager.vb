' Retrieve & Save data from xml file. 
Public Class CConfigManager
    Private m_sConfigFileName As String =
    System.IO.Path.GetFileNameWithoutExtension(
        System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "Conf.xml"
    'System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) & "Conf.xml"
    Private m_oConfig As New CConfigDO

    Public ReadOnly Property ConfigFileName As String
        Get
            Return m_sConfigFileName
        End Get
    End Property

    Public Property Config As CConfigDO
        Get
            Return m_oConfig
        End Get
        Set(value As CConfigDO)
            m_oConfig = value
        End Set
    End Property

  
    'Load configfile
    Public Sub LoadConfig()
        If (System.IO.File.Exists(m_sConfigFileName)) Then
            Dim srReader As System.IO.StreamReader = Nothing
            Try
                srReader = System.IO.File.OpenText(m_sConfigFileName)
                Dim tType As Type = m_oConfig.GetType()
                Dim xsSerializer As System.Xml.Serialization.XmlSerializer = New System.Xml.Serialization.XmlSerializer(tType)
                Dim oData As Object = xsSerializer.Deserialize(srReader)
                m_oConfig = CType(oData, CConfigDO)
            Catch ex As Exception
                Console.WriteLine("Errors with the file:" & m_sConfigFileName & ", we will create a new one.")
                Trace.WriteLine("Errors with the file:" & m_sConfigFileName & ", we will create a new one.")
            Finally
                If srReader IsNot Nothing Then
                    srReader.Close()
                End If
                If m_oConfig Is Nothing Then
                    m_oConfig = New CConfigDO
                End If
            End Try

        End If
    End Sub


    ' Save configfile
    Public Sub SaveConfig()
        Dim swWriter As System.IO.StreamWriter = System.IO.File.CreateText(m_sConfigFileName)
        Dim tType As Type = m_oConfig.GetType()
        If (tType.IsSerializable) Then
            Dim xsSerializer As System.Xml.Serialization.XmlSerializer = New System.Xml.Serialization.XmlSerializer(tType)
            xsSerializer.Serialize(swWriter, m_oConfig)
            swWriter.Close()
        End If
    End Sub

End Class
