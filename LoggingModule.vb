Imports System.IO

Module LoggingModule
    Private LogFileInfo As FileInfo = Nothing
    Private BakFileInfo As FileInfo = Nothing

    Friend Sub Write(ByVal LogData As String, Optional ByVal FileName As String = "", Optional ByVal AppendFile As Boolean = True, _
                     Optional ByVal IncludeTimeStamp As Boolean = True, Optional ByVal LogToLog As Boolean = True, _
                     Optional ByVal LogToConsole As Boolean = True, Optional ByVal MyColor As ConsoleColor = ConsoleColor.Green)
        Dim TimeStamp As String = DateTime.Now.ToString("[MM-dd-yyyy HH:mm:ss]")
        Const MAX_SIZE As Long = 5242880
        Dim LogFileLock As New Object

        'Write text line to console
        If LogToConsole AndAlso Not IsSilent Then
            Console.ForegroundColor = MyColor
            If IncludeTimeStamp AndAlso LogData <> "" Then
                Console.WriteLine(TimeStamp & vbTab & LogData)
            Else
                Console.WriteLine(LogData)
            End If
        End If

        If LogToLog Then
            Dim ThisLog As String = String.Empty
            If (FileName = Nothing) OrElse (FileName = String.Empty) OrElse (FileName = "") Then
                If LogFile <> "" Then
                    ThisLog = LogFile
                Else
                    ThisLog = "C:\Office\PumpSrv\Log\FuelLoader.log"
                End If
            Else
                ThisLog = FileName
            End If '(filename = Nothing) OrElse (filename = String.Empty) OrElse (filename = "") Then

            Dim Backup As String = ThisLog & ".bak"

            If LogFileInfo Is Nothing Then
                LogFileInfo = New IO.FileInfo(ThisLog)
            ElseIf (LogFileInfo.FullName <> ThisLog) Then
                SyncLock LogFileLock
                    LogFileInfo = New IO.FileInfo(ThisLog)
                End SyncLock 'LogFileInfo
            End If 'Logging.LogFileInfo Is Nothing Then

            If BakFileInfo Is Nothing Then
                BakFileInfo = New IO.FileInfo(Backup)
            ElseIf (BakFileInfo.FullName <> Backup) Then
                SyncLock LogFileLock
                    BakFileInfo = New IO.FileInfo(Backup)
                End SyncLock 'LogFileInfo
            End If 'Logging.LogFileInfo Is Nothing Then

            SyncLock LogFileLock
                If LogFileInfo.Directory.Exists = False Then
                    LogFileInfo.Directory.Create()
                End If 'LogFileInfo.Directory.Exists = False Then

                If ((LogFileInfo.Directory.Exists = True) AndAlso (AppendFile = False)) Then
                    With LogFileInfo
                        .Delete()
                        .Refresh()
                    End With 'LogFileInfo
                End If '((LogFileInfo.Directory.Exists = True) AndAlso (appendfile = False)) Then

                If ((LogFileInfo.Exists = True) AndAlso (LogFileInfo.Length > MAX_SIZE)) Then
                    If BakFileInfo.Exists Then
                        BakFileInfo.Delete()
                    End If
                    Try
                        LogFileInfo.MoveTo(Backup)
                    Catch ex As Exception
                        'Do something if unable to move log file
                    End Try
                End If

                If LogFileInfo.Exists = False Then
                    With LogFileInfo.Open(IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite, IO.FileShare.Read)
                        .Close()
                        .Dispose()
                    End With 'LogFileInfo.Open(IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite, IO.FileShare.Read)
                End If 'LogFileInfo.Exists = False Then

                LogFileInfo.Refresh()

                Dim LogText As String = CType(IIf(LogFileInfo.Length > 0, vbNewLine, String.Empty), String) & CType(IIf(IncludeTimeStamp, TimeStamp, String.Empty), String) & vbTab & _
                                LogData

                Dim LogBuf As Byte() = System.Text.Encoding.ASCII.GetBytes(LogText)
                With LogFileInfo.Open(IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite, IO.FileShare.Read)
                    .Position = LogFileInfo.Length

                    .Write(LogBuf, 0, LogBuf.Length)

                    .Close()
                    .Dispose()
                End With 'Logging.LogFileInfo.Open(IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite, IO.FileShare.Read)
            End SyncLock 'Logging.LogFileInfo
        End If
    End Sub 'Write(ByVal text As String)
End Module 'Logging
