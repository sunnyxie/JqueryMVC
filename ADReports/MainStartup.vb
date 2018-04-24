Public Module Program

    ' uncheck the 'Enable Application framework' setting in order to run this. 
    Public Sub Main()
        Dim bRunInConsole As Boolean = False
        Dim DBConn As ADReportsDAO = Nothing
        Dim tInsertDatetime As DateTime
        Static mConfigManager As New CConfigManager

        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        ' Customized error log file
        Trace.Listeners.Add(New TextWriterTraceListener("AppMessages.log"))

        Try
            DBConn = New ADReportsDAO
            DBConn.ConnectToDataRepository()
        Catch ex As Exception
            Trace.WriteLine(Date.Now.ToString() & ": " & ex.Message)
            Trace.Close()
            Return
        End Try

        ' Load file path configurations. 
        mConfigManager.LoadConfig()
        ' mConfigManager.SaveConfig()  'for testing.

        Dim nGUIIndex As Integer = Program.GuiMarkExists(Environment.GetCommandLineArgs())
        'If Not SetExcelFileName(nGUIIndex, sExcelFile) Then
        '    Trace.WriteLine(Date.Now.ToString() & ": " & "File provided is not exist or is not a valid excel file.")
        '    Return
        'End If
        SetInsertDatetime(nGUIIndex, tInsertDatetime)

        bRunInConsole = RunInConsoleMode()
        If bRunInConsole Then
            Trace.WriteLine(Date.Now.ToString() & ": " & "Run in Console Mode. File: " & mConfigManager.Config.ADReportExcelFile)
            Dim sErrMsg As String = String.Empty

            MainForm.FilesMainProcessing(mConfigManager.Config.ADReportExcelFile, DBConn, tInsertDatetime, sErrMsg)

            TextFileProcessing.mDBClass = DBConn
            TextFileProcessing.CsvFilesProcessing(mConfigManager.Config, DBConn, tInsertDatetime, sErrMsg)
            'TextFileProcessing.ReadUserListsCSV(mConfigManager.Config.ADInfoUserListFile)

            ' At the end line.
            Trace.Flush()
        Else
            Trace.WriteLine(Date.Now.ToString() & ": " & "Run in GUI Mode.")
            Dim mform As New MainForm
            mform.SetDBInstance(DBConn)
            mform.SetExcelFileAndTime(mConfigManager.Config.ADReportExcelFile, tInsertDatetime)
            Trace.Flush()

            Application.Run(mform)
        End If

        'Trace.Flush()
        Trace.Close()
    End Sub

    Function RunInConsoleMode() As Boolean
        Dim arguments As String() = Environment.GetCommandLineArgs()

        If arguments.Count > 1 Then
            Dim nGUIMark As Integer = GuiMarkExists(arguments)
            If nGUIMark = 0 Then  'NO /G MARK, THEN RUN IN CONSOLE

                Return True
            Else

                Return False
            End If
        ElseIf arguments.Count = 1 Then
            Return True
        End If

        Return True
    End Function

    ' Check in the parameters list if exists "/g" OR "GUI" mark.
    Function GuiMarkExists(arguments As String()) As Integer
        If arguments.Count = 1 Then
            Return 0
        End If

        For i As Integer = 1 To arguments.Count - 1
            If arguments(i).ToUpper = "/G" OrElse arguments(i).ToUpper = "GUI" Then
                Return i
            End If
        Next

        Return 0
    End Function

    ' Get the filename from command line parameter.
    Function SetExcelFileName(nGUIIndex As Integer, ByRef sExcelFile As String) As Boolean
        Dim nFileNameIndex As Integer = 1
        Dim arguments As String() = Environment.GetCommandLineArgs()
        If nGUIIndex > 0 AndAlso nGUIIndex <= nFileNameIndex Then
            nFileNameIndex += 1
        End If

        If arguments.Count > nFileNameIndex Then
            If My.Computer.FileSystem.FileExists(arguments(nFileNameIndex)) Then
                sExcelFile = arguments(nFileNameIndex).ToString.Trim
                If IO.Path.GetExtension(sExcelFile).ToLower = ".xlsx" _
                        OrElse IO.Path.GetExtension(sExcelFile).ToLower = ".xls" Then
                    Return True
                End If
            End If

            ' File not exists or not valid excel file. 
            Return False
        End If

        sExcelFile = String.Empty
        Return True
    End Function

    ' Get the datetime from command line parameter, or set the system Datetime
    Sub SetInsertDatetime(nGUIIndex As Integer, ByRef DBInsertDatetime As DateTime)
        Dim nDateIndex As Integer = 1
        If nGUIIndex > 0 AndAlso nGUIIndex <= nDateIndex Then
            nDateIndex += 1
        End If

        Dim arguments As String() = Environment.GetCommandLineArgs()
        If arguments.Count > nDateIndex Then
            If IsDate(arguments(nDateIndex)) Then

                DBInsertDatetime = CDate(arguments(nDateIndex))
                Return
            End If
        End If

        DBInsertDatetime = DateTime.Now
    End Sub
End Module
