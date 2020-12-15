Imports System.IO
Imports Microsoft.Win32
Imports System.Text
Imports System.Configuration
Imports System.ServiceProcess
Imports System.Security.Cryptography
Imports System.Data.SqlClient


Module FuelLoader
    Private IsDebug, IsPumpSrvRunning, IsQDEXRunning, DoNotWait, WaitForFunction As Boolean
    Public LogFile As String = ""
    Public OsMode As String = ""
    Public UseQdx As Boolean = True
    Public EmptyMobile As String = "False"
    Private BinaryFlags As String = ""
    Private DecimalFlag As Long = 0
    Private tables As SortedList = New SortedList()
    Private fuelprocess As SortedList = New SortedList()
    Private brokerprocess As SortedList = New SortedList()
    Private reconfigTable As New SortedList(Of String, String)
    Private EmptyList As String = ""
    Private reconfigureParam As String = ""
    Private fmsParam As String = ""
    Private EmptyType As String = "Z"
    Public IsSilent, IsForced, IsCalledByFCC As Boolean
    Private SleepTimeMS As Integer = 54000
    Private Const StartMessage As String = "Please wait for fuel to {0} ..."
    Private Const EndMessage As String = "FuelLoader {0} process has completed."
    Private MyTitle, MyVersion, RegFileList As String
    Private FuelProcessListValidated As Boolean = False
    Private BrokerProcessListValidated As Boolean = False
    Private RegFileListValidated As Boolean = False
    Private AppLdrValidated As Boolean = False
    Private Salt As String = "P!l0+fLy!Nj01"
    Public myConn As SqlConnection = New SqlConnection("Initial Catalog=Fuel;" &
                "Data Source=localhost\FUELDB;Integrated Security=True;")
    Public myCmd As SqlCommand
    Public myReader As SqlDataReader
    Public myProcessId As Integer = Process.GetCurrentProcess().Id

    Enum DoToService
        StopIt = 0
        StartIt = 1
    End Enum

    Enum CreateBatFiles
        QDEX = 0
        FCC = 1
    End Enum


    Sub Main()
        Dim CurrentCommand As String
        Dim Args As New ArrayList(Environment.GetCommandLineArgs())
        Dim MachineName As String ' SiteId
        Dim NewArgs(1) As String

        MyVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString
        MachineName = Environment.MachineName.ToUpper()
        SetPaths()
        If Environment.Is64BitOperatingSystem Then
            OsMode = "64 bit"
            UseQdx = False
        Else
            OsMode = "32 bit"
        End If
        'If log directory does not exist, create it.
        If Not Directory.Exists("C:\Office\PumpSrv\Log") Then
            Try
                Directory.CreateDirectory("C:\Office\PumpSrv\Log")
            Catch ex As Exception

            End Try

        End If

        LoggingModule.Write(String.Format("Loading Pilot FuelLoader ver. {0} on {1}.", MyVersion, MachineName))
        LoggingModule.Write(String.Format("Windows OS mode: {0}", OsMode))

        If IO.File.Exists("C:\Office\PumpSrv\FuelLoader.dbg") Then IsDebug = True
        'IsDebug = True

        MyTitle = "Retalix FuelLoader ver. " & MyVersion
        If IsDebug Then MyTitle = "DEBUG MODE - " & MyTitle
        Console.Title = MyTitle
        InitializeData()
        Args.RemoveAt(0)

        'Scan through command line arguments to determine if silent flag is present.
        'If silent flag is present, no logging to the console will occur.
        If Args.Count > 0 Then
            For Each arg In Args
                CurrentCommand = arg.ToString.ToLower
                LoggingModule.Write(String.Format("Processing command line switch: {0}", CurrentCommand))
                If (((CurrentCommand.Chars(0) = "-"c) OrElse (CurrentCommand.Chars(0) = "/"c)) AndAlso (CurrentCommand.Length > 1)) Then
                    CurrentCommand = CurrentCommand.Substring(1)
                End If
                Select Case CurrentCommand
                    Case "silent"
                        IsSilent = True
                    Case "debug"
                        IsDebug = True
                    Case Else

                End Select
            Next
        End If

        'Read back first command line argument only

        If IsDebug AndAlso Args.Count < 2 Then
            Dim DebugParams As String = ""
            Dim ArgReplace() As String
            LoggingModule.Write("Enter parameters: ", MyColor:=ConsoleColor.Yellow)
            DebugParams = Console.ReadLine
            DebugParams = "NULL " & DebugParams 'First parameter of command line argument is ignored. Keep lined up.
            ArgReplace = DebugParams.Split(" "c)
            If ArgReplace.Count > NewArgs.Count Then
                Array.Resize(NewArgs, ArgReplace.Count)
            End If
            Array.Copy(ArgReplace, NewArgs, ArgReplace.Count)
        Else
            If Args.Count > NewArgs.Count Then
                Array.Resize(NewArgs, Args.Count)
            End If
            Args.CopyTo(NewArgs)
        End If

        Dim runList As SortedList = BuildRunList(NewArgs)
        ProcessRunList(runlist)

        LoggingModule.Write("FuelLoader is terminating.")
        LoggingModule.Write("", IncludeTimeStamp:=False)
        Console.ForegroundColor = ConsoleColor.White 'Return console to default white color font.

    End Sub

    Private Sub DisplayHelp(ByVal LogToLogBool As Boolean)
        LoggingModule.Write("Below are a list of supported parameters for FuelLoader (commands are not case sensitive).", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/HelpLog = Same as help but logs help details to log", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Start = Launch Fuel", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Stop = Shutdown Fuel", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Load = Load QDEX Engine", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Close = Close QDEX Engine", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Kill = Terminates all fuel processes", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Regserver = Initializes FuelLoader settings. Forces required registry values.", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Regall = Initializes fuel only component settings such as PumpSrv, Cl2PumpSrv, FuelMobileSrv, Generator and FCC.", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/RegallFuel = Initializes fuel component settings such as PumpSrv, Cl2PumpSrv, FuelMobileSrv, Generator, FCC and also including FuelLoader.", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Appldr = Loads fuel via AppLdr.", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Empty = Empty ALL QDX files", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Empty=# = Empty specific table by table number", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Empty=#,# = Empty specific table numbers in comma delimited list", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Empty=#-# = Empty specific table numbers in hyphenated range", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Empty=#,#-# = Empty specific table numbers in comma delimited list including range", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Empty={table name} = Empty specific table by table name", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Delete = Deletes QDEX files", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/NoWait = No pause after processes such as Kill and Empty", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Silent = No output to console", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Reconfigure = Reconfigure Fuel", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Reconfigure={table name} = Reconfigure specific table. Table must be one word (i.e. convertlayer, carwash, generalparams).", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Reconfigure={integer(s)} = Reconfigure specific values given by NCR (i.e. 1048576,4,0,0)", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Reconfigure={table name},{integer} = Reconfigure specific table but override additional paramters. (i.e. generalparams,4).", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/FMS={Start/Stop} = Start / Stop FuelMobileSrv.", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/StartBroker = Starts fuel broker (mobile broker).", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/StopBroker = Stops fuel broker (mobile broker).", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/AppLdr = Starts fuel with AppLdr.", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Debug = Debug mode", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/Z = Empty flag for QDEX empty with ZERO.", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("/E = Empty flag for QDEX empty with NULL.", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("*** NOTE 1: {table} name must be provided in lower case, no spaces. ***", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("*** NOTE 2: There must be NO spaces in the Empty parameter!!! ***", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("Examples of multi-parameter use types:", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("C:\Office\PumpSrv\Fueloader.exe /kill /delete /appldr (Stops fuel, deletes QDEX, and launches fuel)", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("C:\Office\PumpSrv\Fueloader.exe /kill /delete /appldr /force /nowait (Stops fuel, deletes QDEX, and launches fuel with no confirmation and no pause)", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("C:\Office\PumpSrv\Fueloader.exe /Empty=2 (Empties OLA table 2)", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("C:\Office\PumpSrv\Fueloader.exe /Empty=2,4 (Empties OLA tables 2 & 4)", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("C:\Office\PumpSrv\Fueloader.exe /Empty=3-5 (Empties OLA tables 3 through 5)", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("C:\Office\PumpSrv\Fueloader.exe /Empty=1,4-6 (Empties OLA tables 1 and 4 through 6)", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("C:\Office\PumpSrv\Fueloader.exe /Empty=prepay (Empties Prepay table)", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("C:\Office\PumpSrv\Fueloader.exe /Empty=1,receipte (Empties OLA table 1 and ReceiptE table)", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("C:\Office\PumpSrv\Fueloader.exe /Reconfigure=generalparams (Reconfigures GeneramParams)", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("C:\Office\PumpSrv\Fueloader.exe /Reconfigure=generalparams,4 (Reconfigures GeneramParams, overriding the second param as 4)", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("Supported reconfigure table options:", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("   all,carwash,database,grades,messages,modes,pricepoles,pumps,pureproducts,receipt,", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("   servicefees,tanks,taxes,terminals,convertlayer,generalparms", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
        LoggingModule.Write("", IncludeTimeStamp:=False, LogToLog:=LogToLogBool)
    End Sub

    Private Sub InitializeData()
        'Retalix FuelLoader
        'Process to Launch / Shutdown Retalix Fuel
        'Define enabled QDEX tabled. Used for DB parameter processing
        tables.Add(1, "prepay")
        tables.Add(2, "ola")
        tables.Add(3, "pumptstat")
        tables.Add(4, "pumptotals")
        tables.Add(5, "tankread")
        tables.Add(6, "delivery")
        tables.Add(7, "alarm")
        tables.Add(8, "carwash")
        tables.Add(9, "receipt")
        tables.Add(10, "rfs")
        tables.Add(13, "extrainfo")
        tables.Add(14, "items")
        tables.Add(15, "alarmsrv")
        tables.Add(22, "receipte")
        tables.Add(23, "recextin")
        tables.Add(27, "receiptz")
        tables.Add(29, "p2pdata")

        reconfigTable.Add("carwash", "1,65535")
        reconfigTable.Add("database", "2,65535")
        reconfigTable.Add("grades", "4,-1")
        reconfigTable.Add("messages", "16,15")
        reconfigTable.Add("modes", "32,65535")
        reconfigTable.Add("pricepoles", "128,255")
        reconfigTable.Add("pumps", "256,-1")
        reconfigTable.Add("pureproducts", "512,65535")
        reconfigTable.Add("receipt", "1024,0")
        reconfigTable.Add("servicefees", "8192,65535")
        reconfigTable.Add("tanks", "32768,-1")
        reconfigTable.Add("taxes", "65536,65535")
        reconfigTable.Add("terminals", "131072,-1")
        reconfigTable.Add("convertlayer", "524288,0")
        reconfigTable.Add("generalparams", "1048576,0")

        Dim HashIt As SaltedHash = New SaltedHash()

        'Define fuel processes for kill process
        Dim fuelprocesses As String = ReadSetting("ProcessList", "FuelSrvBroker.exe,*PilotFlyingJ.Store.Broker.WinService,appldr.exe,tcpcomsrv2.exe,cl2pum~1.exe," &
                                                  "cl2pum~2.exe,cl2pum~3.exe,cl2pum~4.exe,cl2pum~5.exe,cl2pumpsrv.exe,fcc.exe," &
                                                  "fccsys~1.exe,fccsystem.exe,fuelpricechangesrv.exe,genera~1.exe,genera~2.exe," &
                                                  "generator.exe,vp.exe,srsrv.exe,tcpcomcln.exe,mosaiccomsrv.exe,transactiontransmitterserver.exe," &
                                                  "storemonitor.exe,pumpsrv.exe,q-load32.exe,q-dex32.exe,FuelMobileSrv.exe,FMSMonitor.exe")
        Dim fuelprocessesHash As String = ReadSetting("ProcessListHash", "")

        'Validate Fuelprocesses hash and if invalid, set fuel process list to null
        If HashIt.VerifyMd5Hash(MD5.Create(), fuelprocesses, fuelprocessesHash, Salt, 32) Then
            FuelProcessListValidated = True

        End If
        LoggingModule.Write(String.Format("Process List Validated =  {0}", FuelProcessListValidated.ToString))

        Dim processes As String() = fuelprocesses.Split(New Char() {","c})

        ' Loop through result strings with For Each.
        Dim pCnt As Integer = 0
        Dim process As String
        For Each process In processes
            process = Trim(process)
            fuelprocess.Add(pCnt, process)
            pCnt += 1
        Next

        ' Have a separate broker process list with the fully defined path as it is needed for the StartBroker functionality with the legacy broker
        Dim brokerprocesses As String = ReadSetting("BrokerList", "C:\PilotApps\MobileTransactions\FuelSrvBrokerTray\FuelSrvBroker.exe,*PilotFlyingJ.Store.Broker.WinService")
        Dim brokerprocessesHash As String = ReadSetting("BrokerListHash", "")

        'Validate Brokerprocesses hash and if invalid, set broker process list to null
        If HashIt.VerifyMd5Hash(MD5.Create(), brokerprocesses, brokerprocessesHash, Salt, 32) = False Then
            brokerprocesses = ""
        Else
            BrokerProcessListValidated = True
        End If
        LoggingModule.Write(String.Format("Broker Process List =  {0}", brokerprocesses))
        LoggingModule.Write(String.Format("Broker Process List Validated =  {0}", BrokerProcessListValidated.ToString))

        Dim brokers As String() = brokerprocesses.Split(New Char() {","c})

        pCnt = 0
        Dim broker As String
        For Each broker In brokers
            broker = Trim(broker)
            brokerprocess.Add(pCnt, broker)
            pCnt += 1
        Next

        RegFileList = ReadSetting("RegisterAppsList", "C:\Office\PumpSrv\PumpSrv.Exe,C:\Office\PumpSrv\Cl2PumpSrv.Exe,C:\Office\PumpSrv\FuelMobileSrv.Exe,C:\Office\FCC\FCC.Exe,C:\Office\PumpSrv\Generator.exe")
        Dim RegFileListHash As String = ReadSetting("RegisterAppsListHash", "")

        'Validate RegFileList hash and if invalid, set list to null
        If HashIt.VerifyMd5Hash(MD5.Create(), RegFileList, RegFileListHash, Salt, 32) = False Then
            RegFileList = ""
        Else
            RegFileListValidated = True
        End If
        LoggingModule.Write(String.Format("Register Apps List =  {0}", RegFileList))
        LoggingModule.Write(String.Format("Register Apps List Validated =  {0}", RegFileListValidated.ToString))

        'Validate AppLdr.exe before using it to launch applications
        Dim AppLdrHash As String = ReadSetting("AppLdrHash", "")
        If HashIt.VerifyMd5FileHash(MD5.Create(), "C:\Office\exe\AppLdr.exe", AppLdrHash, Salt, 32) Then
            AppLdrValidated = True
        End If
        LoggingModule.Write(String.Format("AppLdr Validated =  {0}", AppLdrValidated.ToString))

        HashIt = Nothing

        'Build binary flag based on current database defined in the list
        For j = 1 To 32
            If tables.ContainsKey(j) Then
                BinaryFlags = "1" & BinaryFlags
            Else
                BinaryFlags = "0" & BinaryFlags
            End If
        Next
        LoggingModule.Write(String.Format("Binary Flags =  {0}", BinaryFlags))

        'Default sleep time as defined by NCR is 54000, but can be configured differently. Set default if not valid values. 
        Try
            SleepTimeMS = CInt(ReadSetting("SleepTime(ms)", "54000"))
        Catch ex As Exception
            SleepTimeMS = 54000
        End Try

        If SleepTimeMS < 30000 Or SleepTimeMS > 60000 Then
            SleepTimeMS = 54000
        End If

        DecimalFlag = ConvertBinaryFlagListToLong(BinaryFlags) 'Define default decimal flag based on the current database types defined
    End Sub

    Private Sub SetPaths()
        'Define log file

        LogFile = "C:\Office\PumpSrv\Log\FuelLoader.log"
    End Sub

    Private Function BuildRunList(ByVal NewArgs() As String) As SortedList
        Dim CommandValue, CurrentCommand As String
        Dim CommandFlags As String()
        Dim runCommandList As SortedList = New SortedList()
        Dim commandCounter As Integer = 0
        Dim IsRunListCommand As Boolean
        For Each arg In NewArgs
            If arg Is Nothing Then
                Exit For
            End If
            CurrentCommand = CStr(arg.ToLower())
            If (((CurrentCommand.Chars(0) = "-"c) OrElse (CurrentCommand.Chars(0) = "/"c)) AndAlso (CurrentCommand.Length > 1)) Then
                CurrentCommand = CurrentCommand.Substring(1)
            End If
            CommandFlags = CurrentCommand.Split("="c)
            CurrentCommand = CommandFlags(0)
            If CommandFlags.Count = 2 Then
                CommandValue = CommandFlags(1)
            Else
                CommandValue = ""
            End If

            LoggingModule.Write(String.Format("Command line parameter: {0}", CurrentCommand))
            IsRunListCommand = True
            Select Case CurrentCommand
                Case "start"
                    runCommandList.Add(commandCounter, "Start")

                Case "stop"
                    runCommandList.Add(commandCounter, "Stop")

                Case "stopbroker"
                    runCommandList.Add(commandCounter, "StopBroker")

                Case "startbroker"
                    runCommandList.Add(commandCounter, "StartBroker")

                Case "load"
                    runCommandList.Add(commandCounter, "Load")

                Case "close"
                    runCommandList.Add(commandCounter, "Close")

                Case "kill"
                    runCommandList.Add(commandCounter, "Kill")

                Case "empty"
                    runCommandList.Add(commandCounter, "Empty")
                    EmptyList = CommandValue

                Case "reconfigure"
                    runCommandList.Add(commandCounter, "Reconfigure")
                    If CommandValue = "" Then
                        reconfigureParam = "all"
                    Else
                        reconfigureParam = CommandValue.ToLower
                    End If

                Case "fms"
                    If CommandValue <> "" Then
                        runCommandList.Add(commandCounter, "FMS")
                        fmsParam = CommandValue

                    End If

                Case "pause"
                    runCommandList.Add(commandCounter, "Pause")

                Case "help"
                    runCommandList.Add(commandCounter, "Help")

                Case "helplog"
                    runCommandList.Add(commandCounter, "HelpLog")

                Case "?"
                    runCommandList.Add(commandCounter, "Help")

                Case "regserver"
                    runCommandList.Add(commandCounter, "Register")

                Case "regall"
                    runCommandList.Add(commandCounter, "Regall")

                Case "regallfuel"
                    runCommandList.Add(commandCounter, "RegallFuel")

                Case "delete"
                    runCommandList.Add(commandCounter, "Delete")

                Case "setup"
                    runCommandList.Add(commandCounter, "Setup")

                Case "appldr"
                    runCommandList.Add(commandCounter, "Appldr")

                Case "nowait"
                    DoNotWait = True
                    IsRunListCommand = False

                Case "silent"
                    IsSilent = True
                    DoNotWait = True
                    IsForced = True
                    IsRunListCommand = False

                Case "force"
                    IsForced = True
                    IsRunListCommand = False

                Case "fcc"
                    IsCalledByFCC = True
                    IsRunListCommand = False

                Case "debug"
                    IsDebug = True
                    IsRunListCommand = False

                Case "e"
                    EmptyType = "E" 'Empty QDEX table flag
                    IsRunListCommand = False

                Case "z"
                    EmptyType = "Z" 'Zero QDEX table flag
                    IsRunListCommand = False

                Case Else
                    LoggingModule.Write("Parameter is not supported.", MyColor:=ConsoleColor.Red)
            End Select
            If IsRunListCommand Then commandCounter += 1
        Next
        Return runCommandList
    End Function


    Private Function ShellRun(ByVal command As String, ByVal parameters As String, ByVal wait As Boolean,
                              Optional ByVal WindowStyle As ProcessWindowStyle = ProcessWindowStyle.Hidden, Optional ByVal RunInShell As Boolean = False,
                              Optional ByVal Redirect As Boolean = False, Optional LoadUserProfile As Boolean = False, Optional IsValidated As Boolean = False) As Integer

        LoggingModule.Write("Entering ShellRun routine (ProcessStartInfo)")
        LoggingModule.Write(String.Format("Application: {0}", command))
        LoggingModule.Write(String.Format("Parameters: {0}", parameters))
        LoggingModule.Write(String.Format("Wait: {0}", wait.ToString))
        LoggingModule.Write(String.Format("Run In Shell: {0}", RunInShell.ToString))
        LoggingModule.Write(String.Format("Redirect Standard Input: {0}", Redirect.ToString))
        LoggingModule.Write(String.Format("Load User Profile: {0}", LoadUserProfile.ToString))
        LoggingModule.Write(String.Format("Validated Process: {0}", IsValidated.ToString))

        Dim ext As String = Path.GetExtension(command)

        If ext = "" Then
            command &= ".exe"
        End If

        Dim procID As Integer
        Dim newProc As Process
        Dim procInfo As New ProcessStartInfo(command)
        procInfo.Arguments = parameters
        procInfo.WindowStyle = WindowStyle
        procInfo.UseShellExecute = RunInShell
        procInfo.RedirectStandardInput = Redirect
        procInfo.Verb = "runas"
        procInfo.LoadUserProfile = LoadUserProfile
        Dim retVal As Integer = 0

        Try
            If File.Exists(command) Then
                If IsValidated Then
                    LoggingModule.Write("Starting " & command)
                    newProc = Process.Start(procInfo)
                    procID = newProc.Id
                    If wait Then
                        LoggingModule.Write("Wait is set to true. Setting process WaitForExit() flag.")
                        newProc.WaitForExit()
                        Dim procEC As Integer = -1
                        If newProc.HasExited Then
                            LoggingModule.Write("Process has exited properly.")
                            procEC = newProc.ExitCode
                            LoggingModule.Write(String.Format("Exit code: {0}", procEC.ToString))
                        End If
                        retVal = procEC
                    End If
                Else
                    LoggingModule.Write(String.Format("Executable was not validated. {0} will not be launched.", command))
                End If
            Else
                retVal = -1
                LoggingModule.Write(String.Format("{0} does not exist. Unable to launch process.", command), MyColor:=ConsoleColor.Red)
            End If

        Catch ex As Exception
            LoggingModule.Write(String.Format("ShellRun Error: {0}", ex.Message), MyColor:=ConsoleColor.Red)
            Return -1
        End Try

        LoggingModule.Write(String.Format("Exiting ShellRun routine ({0})", retVal))
        Return retVal
    End Function

    Private Function ConvertLongToBinaryFlagList(ByVal value As Long) As String
        Dim Result As New StringBuilder()
        For Counter As Integer = 31 To 0 Step -1
            If (value And CLng(2 ^ Counter)) <> 0 Then
                Result.Append("1")
            Else
                Result.Append("0")
            End If
        Next
        Return Result.ToString()
    End Function

    Private Function ConvertBinaryFlagListToLong(ByVal Bin As String) As Long 'function to convert a binary number to decimal
        Dim dec As Long = 0
        Dim power As Integer = 0
        For x As Integer = Bin.Length() - 1 To 0 Step -1
            If Bin.Substring(x, 1) <> "0" Then
                dec += CLng(2 ^ power)
            End If
            power += 1
        Next

        Return dec
    End Function

    Private Function DeleteQDEXFiles() As Integer
        Dim SuccessfulFileCount As Integer = 0

        Try
            For Each FileName As String In IO.Directory.GetFiles("C:\Office\PumpSrv\QDX", "*.qdx")
                If Not IsDebug Then IO.File.Delete(FileName)
                LoggingModule.Write(String.Format("Deleted {0}\{1}.", "C:\Office\PumpSrv\QDX", FileName))
                SuccessfulFileCount += 1
            Next
        Catch ex As Exception
            LoggingModule.Write(String.Format("QDEX Error: Unable to delete file in {0}.", "C:\Office\PumpSrv\QDX"), MyColor:=ConsoleColor.Red)
        End Try
        Return (SuccessfulFileCount)
    End Function

    Private Function LaunchProcess(ByVal MyDir As String, ByVal ProcessParams As String, ByVal Application As String,
                                   Optional ByVal Table As String = "", Optional ByVal wait As Boolean = True,
                                   Optional ByVal WindowStyle As ProcessWindowStyle = ProcessWindowStyle.Hidden,
                                   Optional ByVal RunInShell As Boolean = False, Optional ByVal Redirect As Boolean = False,
                                   Optional ByVal ChangeWorkingDirectory As Boolean = True, Optional ByVal Validated As Boolean = False) As Integer
        'Launch fuel process using ShellRun function to allow parameters and wait/no wait functionality
        Dim StartTime As DateTime = Now()
        Dim retVal As Integer = 0

        LoggingModule.Write(String.Format("Launching {0}\{1} {2}", MyDir, Application, ProcessParams))

        If ChangeWorkingDirectory Then
            Try
                'Set the current directory.
                LoggingModule.Write(String.Format("Setting current directory to {0}.", MyDir))
                Directory.SetCurrentDirectory(MyDir)
            Catch e As DirectoryNotFoundException
                LoggingModule.Write(String.Format("Launch Process Error: Error setting current directory({0}).", e.ToString), MyColor:=ConsoleColor.Red)
            End Try
        End If

        If Not IsDebug Then
            retVal = ShellRun(MyDir & "\" & Application, ProcessParams, wait, WindowStyle, RunInShell, Redirect, IsValidated:=Validated)
        Else
            LoggingModule.Write(String.Format("Debug mode, {0} will not be called.", MyDir & "\" & Application))
        End If

        'Check return value to see if process completed without errors
        If retVal = -1 Then
            LoggingModule.Write(String.Format("Error encountered ({0}).", retVal), MyColor:=ConsoleColor.Red)
        Else
            LoggingModule.Write(String.Format("Completed successfully ({0}).", retVal))
        End If
        Dim EndTime As DateTime = Now()
        Dim TotTimeSeconds = DateDiff(DateInterval.Second, StartTime, EndTime)
        LoggingModule.Write(String.Format("Process completed in {0} seconds.", TotTimeSeconds))
        Return retVal
    End Function

    Public Function GetFileName(ByVal filepath As String) As String

        Dim slashindex As Integer = filepath.LastIndexOf("\")
        Dim dotindex As Integer = filepath.LastIndexOf(".")

        GetFileName = filepath.Substring(slashindex + 1, dotindex - slashindex - 1)
    End Function

    Private Sub CreateBatchFiles(ByVal WhichBatFiles As CreateBatFiles)
        'Define default dates for files to allow version control
        Dim dtCreation As Date = #12/6/2017 12:00:00 PM#
        Dim dtModified As Date = #12/6/2017 12:00:00 PM#
        'QDX\Drv32 files
        Dim SrvStart As String = "C:\office\pumpsrv\qdx\drv32\SrvStart.bat"
        Dim SrvStop As String = "C:\office\pumpsrv\qdx\drv32\SrvStop.bat"
        Dim QEmpty As String = "C:\office\pumpsrv\qdx\drv32\QEmpty.bat"
        Dim QSetup As String = "C:\office\pumpsrv\qdx\drv32\QSetup.bat"
        'FCC files
        Dim FCCSrvStart As String = "C:\office\FCC\FCCSrvStart.bat"
        Dim FCCSrvStop As String = "C:\office\FCC\FCCSrvStop.bat"
        Dim FCCEmptyQdx As String = "C:\office\FCC\FCCEmptyQdx.bat"

        Dim StartFile, StopFile, EmptyFile, startParms, stopParms, emptyParms As String

        If WhichBatFiles = CreateBatFiles.QDEX Then
            If Not UseQdx Then
                LoggingModule.Write("Ceate batch files called for QDEX but system is SQL. No QDEX batch files will be created.")
                Exit Sub
            End If
            StartFile = SrvStart
            StopFile = SrvStop
            EmptyFile = QEmpty
            startParms = "/load"
            stopParms = "/close"
            emptyParms = "/empty /e /nowait"
        Else
            StartFile = FCCSrvStart
            StopFile = FCCSrvStop
            EmptyFile = FCCEmptyQdx
            startParms = "/load /fcc"
            stopParms = "/close /fcc"
            emptyParms = "/empty /z /nowait /fcc"
        End If

        'Create QDX\Drv32 files
        Try
            LoggingModule.Write(String.Format("Updating {0} for new FuelLoader.", StartFile))
            Using writer As New StreamWriter(StartFile)
                writer.WriteLine("@ECHO OFF")
                writer.WriteLine("CLS")
                If WhichBatFiles = CreateBatFiles.FCC Then writer.WriteLine("cd C:\office\PumpSrv\QDX")
                writer.WriteLine(String.Format("C:\office\pumpsrv\FuelLoader.exe {0}", startParms))
                writer.WriteLine("C:\office\pumpsrv\FuelLoader.exe /startbroker")
                If WhichBatFiles = CreateBatFiles.FCC Then writer.WriteLine("cd C:\office\FCC")
                writer.Close()
            End Using
            LoggingModule.Write(String.Format("Updating {0} for new FuelLoader.", StopFile))
            Using writer As New StreamWriter(StopFile)
                writer.WriteLine("@ECHO OFF")
                writer.WriteLine("CLS")
                If WhichBatFiles = CreateBatFiles.FCC Then writer.WriteLine("cd C:\office\PumpSrv\QDX")
                writer.WriteLine("C:\office\pumpsrv\FuelLoader.exe /stopbroker")
                writer.WriteLine(String.Format("C:\office\pumpsrv\FuelLoader.exe {0}", stopParms))
                If WhichBatFiles = CreateBatFiles.FCC Then writer.WriteLine("cd C:\office\FCC")
                writer.Close()
            End Using
            LoggingModule.Write(String.Format("Updating {0} for new FuelLoader.", EmptyFile))
            Using writer As New StreamWriter(QEmpty)
                writer.WriteLine("@ECHO OFF")
                writer.WriteLine("CLS")
                If WhichBatFiles = CreateBatFiles.FCC Then writer.WriteLine("cd C:\office\PumpSrv\QDX")
                writer.WriteLine(String.Format("C:\office\pumpsrv\FuelLoader.exe {0}", emptyParms))
                If WhichBatFiles = CreateBatFiles.FCC Then writer.WriteLine("cd C:\office\FCC")
                writer.Close()
            End Using
            If WhichBatFiles = CreateBatFiles.FCC Then
                LoggingModule.Write(String.Format("Updating {0} for new FuelLoader.", QSetup))
                Using writer As New StreamWriter(QSetup)
                    writer.WriteLine("@ECHO OFF")
                    writer.WriteLine("CLS")
                    writer.WriteLine("cd C:\office\pumpsrv\qdx")
                    writer.WriteLine("drv32\q-Setup")
                    writer.Close()
                End Using
                IO.File.SetCreationTime(QSetup, dtCreation)
                IO.File.SetLastWriteTime(QSetup, dtModified)
            End If

            IO.File.SetCreationTime(SrvStart, dtCreation)
            IO.File.SetLastWriteTime(SrvStart, dtModified)
            IO.File.SetCreationTime(SrvStop, dtCreation)
            IO.File.SetLastWriteTime(SrvStop, dtModified)
            IO.File.SetCreationTime(QEmpty, dtCreation)
            IO.File.SetLastWriteTime(QEmpty, dtModified)

        Catch ex As Exception
            LoggingModule.Write("CreateBatchFilesError: " & ex.Message, MyColor:=ConsoleColor.Red)
        End Try
    End Sub



    Private Function RegisterApp() As Integer

        SetFuelRegistryKeys()
        SetProgLoaderRegistryKey()
        RemoveLegacyFuelLoaderBat()
        CreateBatchFiles(CreateBatFiles.QDEX)
        CreateBatchFiles(CreateBatFiles.FCC)

        Return (1)
    End Function

    Private Function GetRegBasePath() As String
        If Environment.Is64BitOperatingSystem Then
            Return "Software\WOW6432Node"
        Else
            Return "Software"
        End If
    End Function

    Private Function ProcessAppLdrKeys() As Boolean
        Dim AppLdrKey As RegistryKey
        Dim ProgKey As RegistryKey
        Dim SubKeys() As String
        Dim basePath As String = GetRegBasePath()
        Dim Active, Description, Parameters, ProcessKey, ProgramPath, Seconds, StartIn, View, WinExecandWait As String
        Try
            AppLdrKey = Registry.LocalMachine.OpenSubKey(String.Format("{0}\Positive\ProgLoader\Programs", basePath), True)
            SubKeys = AppLdrKey.GetSubKeyNames()
            If SubKeys Is Nothing Or SubKeys.Length <= 0 Then

            Else
                For Each keyname As String In SubKeys
                    ProgKey = AppLdrKey.OpenSubKey(keyname)
                    Active = CStr(ProgKey.GetValue("Active", "0"))
                    If Active = "1" Then
                        Description = CStr(ProgKey.GetValue("Description", ""))
                        Parameters = CStr(ProgKey.GetValue("Parameters", ""))
                        ProcessKey = CStr(ProgKey.GetValue("ProcessKey", ""))
                        ProgramPath = CStr(ProgKey.GetValue("ProgramPath", ""))
                        Seconds = CStr(ProgKey.GetValue("Seconds", "1"))
                        StartIn = CStr(ProgKey.GetValue("StartIn", ""))
                        View = CStr(ProgKey.GetValue("View", "0"))
                        WinExecandWait = CStr(ProgKey.GetValue("WinExecandWait", ""))
                    End If
                    ProgKey.Close()
                Next
            End If
            AppLdrKey.Close()
        Catch ex As Exception

        End Try
    End Function

    Private Sub SetFuelRegistryKeys()
        Dim DbKey, TimerKey As RegistryKey
        Dim basePath As String = GetRegBasePath()

        Try
            DbKey = Registry.LocalMachine.OpenSubKey(String.Format("{0}\PointOfSale\PumpSrv\Database", basePath), True)
            If DbKey Is Nothing Then
                DbKey.CreateSubKey(String.Format("{0}\PointOfSale\PumpSrv\Database", basePath))
                LoggingModule.Write(String.Format("{0}\PointOfSale\PumpSrv\Database\DatabaseList value does not exist in the registry.", basePath), MyColor:=ConsoleColor.Red)
            End If
            DbKey.SetValue("DatabaseList", DecimalFlag, RegistryValueKind.DWord)
            LoggingModule.Write(String.Format("Setting Database list default value = {0}.", DecimalFlag), MyColor:=ConsoleColor.Yellow)
            DbKey.Close()
        Catch ex As Exception
            LoggingModule.Write(String.Format("SetFuelRegistryKeys Error: Unable to read {0}\PointOfSale\PumpSrv\Database registry key.", basePath), MyColor:=ConsoleColor.Red)
        End Try

        Try
            TimerKey = Registry.LocalMachine.OpenSubKey(String.Format("{0}\PointOfSale\FuelCompMng\GeneralParam", basePath), True)
            If TimerKey Is Nothing Then
                TimerKey.CreateSubKey(String.Format("{0}\PointOfSale\FuelCompMng\GeneralParam", basePath))
                LoggingModule.Write(String.Format("{0}\PointOfSale\FuelCompMng\GeneralParam key does not exist in the registry.", basePath), MyColor:=ConsoleColor.Red)
            End If
            TimerKey.SetValue("UseTcpCom", 1, RegistryValueKind.DWord)
            TimerKey.SetValue("UseVP", 1, RegistryValueKind.DWord)
            TimerKey.SetValue("SleepTime(MSec)", SleepTimeMS, RegistryValueKind.DWord)
            LoggingModule.Write("Setting UseTcpCom default value = 1.", MyColor:=ConsoleColor.Yellow)
            LoggingModule.Write("Setting UseVP default value = 1.", MyColor:=ConsoleColor.Yellow)
            LoggingModule.Write(String.Format("Setting SleepTime(MSec) default value = {0}.", SleepTimeMS), MyColor:=ConsoleColor.Yellow)
            TimerKey.Close()
        Catch ex As Exception
            LoggingModule.Write(String.Format("SetFuelRegistryKeys Error: Unable to read the {0}\PointOfSale\FuelCompMng\GeneralParam registry key.", basePath), MyColor:=ConsoleColor.Red)
        End Try
    End Sub

    Private Sub SetProgLoaderRegistryKey()
        Dim ProgLoaderKey As RegistryKey
        Dim basePath As String = GetRegBasePath()
        Try
            ProgLoaderKey = Registry.LocalMachine.OpenSubKey(String.Format("{0}\Positive\ProgLoader\Programs\01", basePath), True)
            If ProgLoaderKey Is Nothing Then
                ProgLoaderKey.CreateSubKey(String.Format("{0}\Positive\ProgLoader\Programs\01", basePath))
                LoggingModule.Write(String.Format("{0}\Positive\ProgLoader\Program\01 key does not exist in the registry. Creating key.", basePath), MyColor:=ConsoleColor.Red)
            End If
            ProgLoaderKey.SetValue("ProgramPath", "C:\Office\PumpSrv\FuelLoader.exe")
            ProgLoaderKey.SetValue("Parameters", "/Start")
            LoggingModule.Write("Setting ProgramPath default value = C:\Office\PumpSrv\FuelLoader.exe.", MyColor:=ConsoleColor.Yellow)
            LoggingModule.Write("Setting Parameters default value = /Start.", MyColor:=ConsoleColor.Yellow)
            ProgLoaderKey.Close()
        Catch ex As Exception
            LoggingModule.Write(String.Format("Unable to read the {0}\Positive\ProgLoader\Program\01 registry key.", basePath), MyColor:=ConsoleColor.Red)
        End Try
    End Sub

    Private Sub RemoveLegacyFuelLoaderBat()
        If IO.File.Exists("C:\Office\PumpSrv\FuelLoader.bat") Then
            Try
                LoggingModule.Write("Deleting C:\Office\PumpSrv\FuelLoader.bat", LogFile, True, True)
                IO.File.SetAttributes("C:\Office\PumpSrv\FuelLoader.bat", FileAttributes.Normal)
                IO.File.Delete("C:\Office\PumpSrv\FuelLoader.bat")
            Catch ex As Exception
                LoggingModule.Write("RemoveLegacyFuelLoaderBatch Error: Unable to delete C:\Office\PumpSrv\FuelLoader.bat", MyColor:=ConsoleColor.Red)
            End Try
        End If
    End Sub
    Private Sub RegisterFuel(ByVal IsValidated As Boolean)
        If IsValidated Then
            Dim RegFileList As String = ReadSetting("RegisterAppsList", "C:\Office\PumpSrv\PumpSrv.Exe,C:\Office\PumpSrv\Cl2PumpSrv.Exe,C:\Office\PumpSrv\FuelMobileSrv.Exe,C:\Office\FCC\FCC.Exe,C:\Office\PumpSrv\Generator.exe")
            Dim RegFile, CurrDir, CurrFile As String
            Dim RegFiles As String() = RegFileList.Split(New Char() {","c})
            For Each RegFile In RegFiles
                RegFile = Trim(RegFile)
                Dim FileInfo As New FileInfo(RegFile)
                CurrFile = FileInfo.Name
                CurrDir = FileInfo.DirectoryName
                If CurrFile <> "" AndAlso CurrDir <> "" Then

                    LaunchProcess(CurrDir, "/RegServer", CurrFile, "", False, ProcessWindowStyle.Normal, True, False, ChangeWorkingDirectory:=False, Validated:=IsValidated)
                End If
                FileInfo = Nothing
            Next
        Else
            LoggingModule.Write("ERROR: Register Apps list is invalid. No apps will be registered.")
        End If

    End Sub

    Private Function BuildFlags(ByVal CmdFlags As String) As String
        Dim NewBinaryFlag As String = "00000000000000000000000000000000"
        'Build binary flag based on current databased defined in the list
        Dim NewFlags As String() = CmdFlags.Split(","c)
        Dim Flag As String
        Dim NewIndex As Integer = -1
        Dim TableIndex As Integer = -1
        Dim StartInteger As Integer = 0
        Dim EndInteger As Integer = 0


        For Each Flag In NewFlags
            If Not IsNumeric(Flag) Then
                Dim RangeFlags As String() = Flag.Split("-"c)
                If RangeFlags.Count() = 2 Then
                    Try
                        StartInteger = CInt(RangeFlags(0))
                        EndInteger = CInt(RangeFlags(1))
                        For Looper As Integer = StartInteger To EndInteger
                            NewIndex = 32 - CInt(Looper)
                            If ((NewIndex >= 0) AndAlso (NewIndex <= 31)) Then
                                NewBinaryFlag = NewBinaryFlag.Remove(NewIndex, 1).Insert(NewIndex, "1")
                            End If
                        Next
                        Flag = "0"

                    Catch ex As Exception
                        LoggingModule.Write("BuildFlagsError: Error processing ranges: " & ex.Message, MyColor:=ConsoleColor.Red)
                    End Try

                Else
                    TableIndex = CInt(tables.IndexOfValue(Flag.ToLower))
                    If TableIndex > -1 Then 'Table name found in defined list
                        LoggingModule.Write(Flag & " table is defined. Retrieving index.")
                        Flag = CStr(tables.Keys(TableIndex)) 'Return key for index = DB #. Replace parameter 2.
                        LoggingModule.Write("Index retrieved = " & Flag)
                    Else
                        LoggingModule.Write(Flag & " table is NOT defined.", MyColor:=ConsoleColor.Red)
                        Flag = "0"
                    End If
                End If
            End If

            Try
                NewIndex = 32 - CInt(Flag)
                If ((NewIndex >= 0) AndAlso (NewIndex <= 31)) Then
                    NewBinaryFlag = NewBinaryFlag.Remove(NewIndex, 1).Insert(NewIndex, "1")
                End If

            Catch ex As Exception
                LoggingModule.Write("BuildFlagsError: Error processing binary flags: " & ex.Message, MyColor:=ConsoleColor.Red)
            End Try

        Next
        Return (NewBinaryFlag)
    End Function

    Private Function CollectConsoleInput(ByVal CaptureData As Boolean) As String
        Dim info As ConsoleKeyInfo
        Dim response As String = ""

        If CaptureData Then
            LoggingModule.Write("Enter at least 2 initials (or your name) and press ENTER", IncludeTimeStamp:=False, LogToLog:=False)
            LoggingModule.Write("to continue, or just press ENTER or ESC to terminate: ", IncludeTimeStamp:=False, LogToLog:=False)
            response = Console.ReadLine()
            Return response
        Else
            info = Console.ReadKey(True)
            response = ""
        End If
        Return (response)
    End Function

    Private Function HandleProcesses(ByVal MyProcessName As String, ByVal PerformKill As Boolean, Optional ByVal IsValidated As Boolean = False) As Boolean
        If IsValidated Then
            Dim psList() As Process
            Try
                psList = Process.GetProcessesByName(GetFileName(MyProcessName))
                'psList = Process.GetProcesses()
                For Each p As Process In psList
                    LoggingModule.Write(String.Format("{0} process was running ({1}).", MyProcessName, p.Id), MyColor:=ConsoleColor.Yellow)
                    If myProcessId = p.Id Then
                        LoggingModule.Write(String.Format("{0} process ID is current application ID. Cannot terminate self.", myProcessId), MyColor:=ConsoleColor.Yellow)
                        PerformKill = False
                    End If
                    If PerformKill AndAlso IsValidated Then
                        LoggingModule.Write(String.Format("Terminating {0} ({1}).", MyProcessName, p.Id), MyColor:=ConsoleColor.Red)
                        If Not IsDebug Then p.Kill()
                    End If
                Next p

            Catch ex As Exception
                LoggingModule.Write("HandleProcessesError: " & ex.Message, MyColor:=ConsoleColor.Red)
            End Try
            Return Process.GetProcessesByName(GetFileName(MyProcessName)).Count > 0
        Else
            Return False
        End If

    End Function

    Private Sub LoadDatabaseFromRegistry()
        Dim dectmp As Long = 0 'Temp decimal flag to read registry value if exists
        Dim PathsKey As RegistryKey
        Dim basePath As String = GetRegBasePath()

        Try
            PathsKey = Registry.LocalMachine.OpenSubKey(String.Format("{0}\PointOfSale\PumpSrv\Database", basePath), True)
            If PathsKey Is Nothing Then
                PathsKey.CreateSubKey(String.Format("{0}\PointOfSale\PumpSrv\Database", basePath))
            End If
            dectmp = CLng(PathsKey.GetValue("DatabaseList", 0))
            If dectmp = 0 Then
                PathsKey.SetValue("DatabaseList", DecimalFlag, RegistryValueKind.DWord)
                LoggingModule.Write("Database list value does not exist in the registry.", , MyColor:=ConsoleColor.Yellow)
                LoggingModule.Write(String.Format("Setting default value = {0}.", DecimalFlag), MyColor:=ConsoleColor.Yellow)
            Else
                DecimalFlag = dectmp
                LoggingModule.Write(String.Format("Loading database flag list from the registry. Value = {0}.", DecimalFlag))
            End If
            PathsKey.Close()
        Catch ex As Exception
            LoggingModule.Write(String.Format("LoadDatabaseFromRegistry Error: Unable to read the registry. Using default database list ({0}).", DecimalFlag), MyColor:=ConsoleColor.Yellow)

        End Try
        BinaryFlags = ConvertLongToBinaryFlagList(DecimalFlag)
    End Sub



    Private Sub EmptyQdx(ByVal LaunchType As String)
        Dim DbType As String
        If UseQdx Then
            DbType = "QDEX database"
        Else
            DbType = "Fuel SQL database"
        End If

        Dim DbBitFlag As Integer = -1
        Dim UserInput As String
        UserInput = PromptEmpty(DbType)

        If UserInput.Length > 1 Then
            If UserInput <> "-1" Then LoggingModule.Write(String.Format("User entered a valid value ({0}). Continuing with empty {1} process.", UserInput, DbType))

            If UseQdx Then
                'Try to read second command line argument if exists
                If EmptyList <> "" Then
                    BinaryFlags = BuildFlags(EmptyList)
                Else
                    LoggingModule.Write("Database table parameter was not provided.")
                    LoggingModule.Write("Loading enabled tables from the registry.")
                    LoadDatabaseFromRegistry()
                End If

                LoggingModule.Write(LaunchType & " QDX in progress...")
                LoggingModule.Write("QDX table binary flags = " & BinaryFlags)
                Dim Counter As Integer = 0
                'Step through binary flag list to launch enabled databases.
                For j = 31 To 0 Step -1
                    Counter += 1
                    DbBitFlag = CInt(BinaryFlags.Substring(j, 1))
                    If DbBitFlag = 1 Then
                        LoggingModule.Write(String.Format("Binary bit {0} = 1. Retrieving table name for record {1}.", Counter, Counter))
                        LoggingModule.Write(String.Format("Emptying {0} table: ", tables(Counter).ToString))
                        LaunchProcess("C:\Office\PumpSrv\QDX", "/" & EmptyType & Counter, "drv32\q-Empt32", Validated:=True)

                    Else
                        LoggingModule.Write(String.Format("Binary bit {0} = 0. Record {0} will not be emptied.", Counter, Counter))
                    End If
                Next
                EmptyFuelDbTables()
            Else
                EmptyFuelDbTables()
            End If

        Else
            LoggingModule.Write(String.Format("User did not enter at least two initials({0}). Fuel empty {1} process will not be completed.", UserInput, DbType), MyColor:=ConsoleColor.Red)
        End If
        If Not DoNotWait Then LoggingModule.Write(String.Format("FuelLoader empty {0} process has completed.", DbType))

    End Sub

    Private Sub KillFuel(ByVal LaunchType As String)
        Dim UserInput As String = ""
        If Not IsForced Then
            'If killing fuel was not forced with the proper flag as intended from a script, prompt the user
            'to enter their initials, or any two characters to make sure fuel is not killed by accident.
            LoggingModule.Write("You are about to terminate fuel abruptly. Please verify no one is fueling", , MyColor:=ConsoleColor.Red)
            LoggingModule.Write("before continuing! This process is only suggested if you plan to delete ", , MyColor:=ConsoleColor.Red)
            LoggingModule.Write("QDEX. Are you sure you want to continue? Please type your initials and press", , MyColor:=ConsoleColor.Red)
            LoggingModule.Write("ENTER to continue or just press ENTER or ESC to terminate without killing fuel.", , MyColor:=ConsoleColor.Red)
            LoggingModule.Write("*** NOTE: At least (2) initials must be provided to continue! ***", , MyColor:=ConsoleColor.Red)
            LoggingModule.Write("", IncludeTimeStamp:=False, LogToLog:=False)
            UserInput = CollectConsoleInput(True)
        Else
            LoggingModule.Write("Force parameter was found. Process will continue in forced mode.", , MyColor:=ConsoleColor.Yellow)
            UserInput = "-1"
        End If

        If UserInput.Length > 1 Then
            If UserInput <> "-1" Then LoggingModule.Write(String.Format("User entered a valid value ({0}). Continuing with Fuel Kill process.", UserInput))
            Dim proc As String = ""
            Dim id As Integer = -1
            'Loop through the process list in order by index id base 0
            HandleProcesses("FuelLoader.exe", True, True) 'Kill other instances of FuelLoader if already running
            For Index = 0 To fuelprocess.Count - 1
                proc = fuelprocess.GetByIndex(Index).ToString
                If proc.Substring(0, 1) = "*" Then
                    StartStopService(proc.Substring(1), DoToService.StopIt, FuelProcessListValidated)
                Else
                    id = CInt(fuelprocess.GetKey(Index))
                    HandleProcesses(proc, True, FuelProcessListValidated)
                    Threading.Thread.Sleep(100)
                End If

            Next
            If FuelProcessListValidated = False Then LoggingModule.Write("ERROR: Fuel process list was invalid No processes or services were terminated or stopped!")
            LoggingModule.Write("Completed processing fuel terminate process list.")
        Else
            LoggingModule.Write(String.Format("User did not enter at least two initials({0}). Fuel Kill process will not be completed.", UserInput), MyColor:=ConsoleColor.Red)
        End If
        If Not DoNotWait Then LoggingModule.Write("FuelLoader Kill process has completed.")
    End Sub

    Private Sub StopStartBroker(ByVal DoWhat As DoToService, ByVal IsValidated As Boolean)
        If IsValidated Then
            Dim proc, myproc, mypath, what As String
            Dim id As Integer = -1
            If DoWhat = DoToService.StartIt Then
                what = "Starting"
            Else
                what = "Stopping"
            End If
            LoggingModule.Write(String.Format("{0} fuel broker.", what), MyColor:=ConsoleColor.Yellow)
            'Loop through the process list in order by index id base 0
            For Index = 0 To brokerprocess.Count - 1
                proc = brokerprocess.GetByIndex(Index).ToString
                If proc.Substring(0, 1) = "*" Then
                    StartStopService(proc.Substring(1), DoWhat, IsValidated)
                Else
                    myproc = Path.GetFileName(proc)
                    mypath = Path.GetDirectoryName(proc)
                    If DoWhat = DoToService.StartIt Then
                        'LoggingModule.Write(String.Format("Launching {0}.", proc), MyColor:=ConsoleColor.Green)
                        LaunchProcess(mypath, "", myproc, "", WindowStyle:=ProcessWindowStyle.Normal, wait:=False, ChangeWorkingDirectory:=False, RunInShell:=True, Validated:=IsValidated)
                    Else
                        'id = CInt(brokerprocess.GetKey(Index))
                        HandleProcesses(myproc, True, IsValidated)
                        Threading.Thread.Sleep(100)
                    End If

                End If

            Next
            LoggingModule.Write(String.Format("Completed {0} fuel broker.", what.ToLower))
        Else
            LoggingModule.Write("ERROR: Broker process list was not valid. No changes made to broker processes or services.")
        End If

    End Sub

    Private Function DeleteQdx(ByVal LaunchType As String) As Boolean
        If UseQdx Then
            LoggingModule.Write("Please wait for fuel to delete QDEX database ...")
            IsQDEXRunning = False
            LoggingModule.Write("Verifying Q-DEX process is not running.")
            IsQDEXRunning = HandleProcesses("Q-DEX32.exe", False)
            Dim UserInput As String = ""
            'Verify QDEX is not running before attempting to delete the QDEX database files
            If IsQDEXRunning Then
                LoggingModule.Write("ERROR! You cannot delete QDEX files while the QDEX Engine is running.", MyColor:=ConsoleColor.Red)
                LoggingModule.Write("Please unload the QDEX engine before continuing!")
                Threading.Thread.Sleep(5000)
                Return False
            Else
                'Delete all files ending with QDX and LOG in the QDEX folder. 
                Dim FileCnt As Integer = DeleteQDEXFiles()
                LoggingModule.Write(String.Format("Deleted {0} file(s).", FileCnt), MyColor:=ConsoleColor.Red)
                LoggingModule.Write("FuelLoader delete QDEX process has completed.")
                EmptyFuelDbTables()
                Return True
            End If
        Else
            Dim DbType As String = "Fuel SQL database"
            Dim UserInput As String
            LoggingModule.Write("QDEX is not supported on 64 bit OS. Empyting Fuel SQL database instead.")

            UserInput = PromptEmpty(DbType)

            If UserInput.Length > 1 Then
                If UserInput <> "-1" Then LoggingModule.Write(String.Format("User entered a valid value ({0}). Continuing with empty {1} process.", UserInput, DbType))
                EmptyFuelDbTables()
            End If

        End If

    End Function

    Private Function PromptEmpty(ByVal DbType As String) As String

        LoggingModule.Write(String.Format("Please wait for fuel to empty {0} ...", DbType))
        IsPumpSrvRunning = False
        LoggingModule.Write("Verifying PumpSrv.exe process is not running.")
        IsPumpSrvRunning = HandleProcesses("PumpSrv.exe", False)
        Dim UserInput As String = ""
        If IsPumpSrvRunning Then
            LoggingModule.Write(String.Format("WARNING: You are about to empty {0} while fuel is running!", DbType), MyColor:=ConsoleColor.Red)
            If Not IsForced Then
                LoggingModule.Write("Are you sure you want to continue? Please type your initials and press", MyColor:=ConsoleColor.Red)
                LoggingModule.Write(String.Format("ENTER to continue or just press ENTER or ESC to terminate without emptying {0}.", DbType), MyColor:=ConsoleColor.Red)
                LoggingModule.Write("*** NOTE: At least (2) initials must be provided to continue! ***", MyColor:=ConsoleColor.Red)
                LoggingModule.Write("", IncludeTimeStamp:=False, LogToLog:=False)
                UserInput = CollectConsoleInput(True)
            Else
                LoggingModule.Write("Force parameter was found. Process will continue in forced mode.", MyColor:=ConsoleColor.Yellow)
                UserInput = "-1"
            End If
        Else

            LoggingModule.Write(String.Format("PumpSrv was not running. Continuing with empty {0}.", DbType))
            UserInput = "-1"
        End If
        Return UserInput
    End Function

    Private Sub ProcessRunList(ByVal runlist As SortedList)
        Dim LaunchType As String
        If runlist.Count > 0 Then
            For CommandLoop As Integer = 0 To runlist.Count - 1
                LaunchType = runlist.GetByIndex(CommandLoop).ToString
                MyTitle = "Retalix FuelLoader ver. " & MyVersion & "     *** " & LaunchType & " Fuel ***"
                If IsDebug Then MyTitle = "DEBUG MODE - " & MyTitle
                Console.Title = MyTitle
                'Launch fuelcontrol with the correct parameter based on what was passed in the command line.
                If IsCalledByFCC Then
                    LoggingModule.Write("*** Process called by FCC ***")
                End If

                Select Case LaunchType

                    Case "Start"
                        LoggingModule.Write(String.Format(StartMessage, LaunchType))
                        LaunchProcess("C:\Office\PumpSrv", "/LaunchFuel", "FuelControl.exe", "", True, ProcessWindowStyle.Normal, False, True, Validated:=True)
                        LoggingModule.Write(String.Format(EndMessage, LaunchType))

                    Case "Stop"
                        LoggingModule.Write(String.Format(StartMessage, LaunchType))
                        LaunchProcess("C:\Office\PumpSrv", "/ShutdownFuel", "FuelControl.exe", "", True, ProcessWindowStyle.Normal, False, True, Validated:=True)
                        LoggingModule.Write(String.Format(EndMessage, LaunchType))

                    Case "Load"  'Add logic here
                        'Load fuel is used specificly for SrvStart.bat to load QDEX in a new window. 
                        LoggingModule.Write(String.Format(StartMessage, LaunchType))

                        If UseQdx Then
                            LoggingModule.Write("Launching QDEX Engine.")
                            LaunchProcess("C:\Office\PumpSrv\QDX", "\NoWindow", "drv32\Q-DEX32.exe", "", False, ProcessWindowStyle.Hidden, True, False, Validated:=True)
                            'Load QDEX tables
                            LoggingModule.Write("*** Preparing to load QDEX tables ***")
                            Threading.Thread.Sleep(3000)
                            LaunchProcess("C:\Office\PumpSrv\QDX", "", "drv32\Q-Load32.exe", "", True, ProcessWindowStyle.Normal, False, False, Validated:=True)
                        Else
                            LoggingModule.Write("QDEX is not supported on 64 bit OS. QDEX will not be loaded.", MyColor:=ConsoleColor.Red)
                        End If
                        'Launch QDEX Engine

                        LoggingModule.Write(String.Format(EndMessage, LaunchType))

                    Case "Close"
                        'Close fuel is called directly from SrvStop.bat to stop fuel. Used by FCC.
                        LoggingModule.Write(String.Format(StartMessage, LaunchType))
                        If UseQdx Then
                            LaunchProcess("C:\Office\PumpSrv\QDX", "-2", "drv32\Q-Clos32.exe", "", True, ProcessWindowStyle.Normal, False, False, Validated:=True)
                        Else
                            LoggingModule.Write("QDEX is not supported on 64 bit OS. QDEX will not be closed.", MyColor:=ConsoleColor.Red)
                        End If

                        LoggingModule.Write(String.Format(EndMessage, LaunchType))

                    Case "Setup"
                        LoggingModule.Write(String.Format(StartMessage, "launch QDEX Setup"))
                        If UseQdx Then
                            LaunchProcess("C:\Office\PumpSrv\QDX", "", "drv32\Q-Setup.exe", "", True, ProcessWindowStyle.Normal, False, True, Validated:=True)
                        Else
                            LoggingModule.Write("QDEX is not supported on 64 bit OS. QDEX Setup will not be launched.", MyColor:=ConsoleColor.Red)
                        End If

                        LoggingModule.Write(String.Format(EndMessage, LaunchType))

                    Case "Register"
                        LoggingModule.Write("Registering FuelLoader application ...")
                        RegisterApp()
                        LoggingModule.Write("FuelLoader Register process has completed.")

                    Case "Regall"
                        LoggingModule.Write("Registering Fuel only applications ...")
                        RegisterFuel(RegFileListValidated)
                        LoggingModule.Write("Fuel Register process has completed.")

                    Case "RegallFuel"
                        LoggingModule.Write("Registering Fuel applications including FuelLoader ...")
                        RegisterApp()
                        RegisterFuel(RegFileListValidated)
                        LoggingModule.Write("Fuel and FuelLoader Register process has completed.")

                    Case "Appldr"
                        LoggingModule.Write("Launching AppLdr ...")
                        'LaunchProcess("C:\Office\Exe", "", "AppLdr.exe", "", False, ProcessWindowStyle.Normal, True, False, ChangeWorkingDirectory:=True, Validated:=AppLdrValidated)
                        ShellRun("C:\Office\Exe\AppLdr.exe", "", False, ProcessWindowStyle.Normal, False, False, True, AppLdrValidated)
                        LoggingModule.Write("FuelLoader is terminating.")
                        LoggingModule.Write("", IncludeTimeStamp:=False)
                        Console.ForegroundColor = ConsoleColor.White 'Return console to default white color font.
                        End
                        'LoggingModule.Write("AppLdr has completed.")

                    Case "Pause"
                        LoggingModule.Write("Pausing application ...")
                        Threading.Thread.Sleep(5000)

                    Case "Delete"
                        WaitForFunction = DeleteQdx(LaunchType)

                    Case "FMS"
                        FMSStartStop(fmsParam)

                    Case "Help"
                        DisplayHelp(False)
                        WaitForFunction = True

                    Case "HelpLog"
                        DisplayHelp(True)
                        WaitForFunction = True

                    Case "Empty"
                        EmptyQdx(LaunchType)
                        WaitForFunction = True

                    Case "Kill"
                        KillFuel(LaunchType)
                        WaitForFunction = True

                    Case "StopBroker"
                        StopStartBroker(DoToService.StopIt, BrokerProcessListValidated)

                    Case "StartBroker"
                        StopStartBroker(DoToService.StartIt, BrokerProcessListValidated)

                    Case "Reconfigure"
                        ReconfigureFuel(reconfigureParam)

                    Case Else
                        LoggingModule.Write(String.Format("FuelLoader {0} processes is not supported.", LaunchType), MyColor:=ConsoleColor.Red)
                        Threading.Thread.Sleep(2000)

                End Select

                If (WaitForFunction AndAlso Not DoNotWait) OrElse IsDebug Then
                    LoggingModule.Write("Press any key to continue...", IncludeTimeStamp:=False, LogToLog:=False)
                    Console.ReadKey()
                Else
                    Threading.Thread.Sleep(1000)
                End If
                WaitForFunction = False
            Next
        Else
            LoggingModule.Write("No parameters were provided.", MyColor:=ConsoleColor.Red)
            Threading.Thread.Sleep(3000)
        End If
    End Sub

    Private Function ReadSetting(key As String, defaultval As String) As String
        Try
            Dim appSettings = ConfigurationManager.AppSettings

            Dim result As String = appSettings(key)
            If IsNothing(result) Then
                AddUpdateAppSettings(key, defaultval)
                Return defaultval
            Else
                Return result
            End If
        Catch e As ConfigurationErrorsException
            LoggingModule.Write(String.Format("ReadSettings Error: Error reading app settings: {0}", key), MyColor:=ConsoleColor.Red)
            Return defaultval
        End Try

    End Function

    Private Function AddUpdateAppSettings(key As String, value As String) As Boolean

        Try
            Dim configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
            Dim settings = configFile.AppSettings.Settings
            If IsNothing(settings(key)) Then
                settings.Add(key, value)
            Else
                settings(key).Value = value
            End If
            configFile.Save(ConfigurationSaveMode.Modified)
            ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name)
            Return True
        Catch e As ConfigurationErrorsException
            LoggingModule.Write(String.Format("AddUpdateAppSettings Error: Error writing app setting: {0}", key), MyColor:=ConsoleColor.Red)
            LoggingModule.Write(String.Format("Value: {0}", value), MyColor:=ConsoleColor.Red)
            Return False
        End Try

    End Function

    Private Sub StartStopService(ByVal ServiceName As String, ByVal DoWhat As DoToService, ByVal IsValidated As Boolean)
        If IsValidated Then
            Try
                Dim service As ServiceController = New ServiceController(ServiceName)

                If DoWhat = DoToService.StartIt Then
                    If ((service.Status.Equals(ServiceControllerStatus.Stopped)) Or (service.Status.Equals(ServiceControllerStatus.StopPending)) Or (service.Status.Equals(ServiceControllerStatus.Paused))) Then
                        LoggingModule.Write(String.Format("{0} service was not running.", ServiceName), MyColor:=ConsoleColor.Yellow)
                        service.Start()
                        LoggingModule.Write(String.Format("{0} service was started.", ServiceName), MyColor:=ConsoleColor.Green)
                    Else
                        LoggingModule.Write(String.Format("{0} service was found, but is already running.", ServiceName), MyColor:=ConsoleColor.Yellow)
                    End If
                Else
                    If service.Status.Equals(ServiceControllerStatus.Running) Then
                        LoggingModule.Write(String.Format("{0} service was running.", ServiceName), MyColor:=ConsoleColor.Yellow)
                        service.Stop()
                        LoggingModule.Write(String.Format("{0} service stopped.", ServiceName), MyColor:=ConsoleColor.Red)
                    Else
                        LoggingModule.Write(String.Format("{0} service was found, but is not currently running.", ServiceName), MyColor:=ConsoleColor.Yellow)
                    End If
                End If
            Catch ex As Exception
                LoggingModule.Write(String.Format("StopStartService Error: Error processing service: {0}", ServiceName), MyColor:=ConsoleColor.Red)
                LoggingModule.Write(ex.Message, MyColor:=ConsoleColor.Red)
            End Try
        End If

    End Sub
    Private Function IsValidReconfig(ByVal Value As String) As Boolean
        If GetParam(Value, 1) <> -2 AndAlso GetParam(Value, 2) <> -2 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function GetParam(ByVal Parm As String, ByVal ParmIndex As Integer, Optional ByVal IsNumeric As Boolean = False) As Integer

        Try
            Dim parms As String() = Parm.Split(New Char() {","c})
            If ParmIndex <= parms.Length Then
                Return CInt(parms(ParmIndex - 1))
            Else
                If ParmIndex > 2 OrElse IsNumeric Then
                    Return 0 'allow default of 0 for third and fourth parameters if exist
                Else
                    Return -2
                End If
            End If

        Catch ex As Exception
            Return -2
        End Try

    End Function

    Private Sub ReconfigureFuel(ByVal table As String)
        LoggingModule.Write("Entering Reconfigure Fuel routine.")
        LoggingModule.Write(String.Format("Table = {0}", table))
        Dim configVars As String()
        Try
            Dim pumpsrv = New PUMPSRVLib.Ctrl()
            Dim value As String = ""
            If table = "all" Then
                LoggingModule.Write("All tables are selected. Looping through fuel tables to reconfigure.")
                For Each kvp As KeyValuePair(Of String, String) In reconfigTable 'Loop through all tables defined in the reconfigTable list
                    LoggingModule.Write(String.Format("Returning PumpSrv parameters for {0}.", kvp.Key))
                    LoggingModule.Write(String.Format("Value returned =  {0}.", kvp.Value))
                    If kvp.Value <> "" Then
                        If IsValidReconfig(kvp.Value) Then
                            pumpsrv.ReConfigure2(GetParam(kvp.Value, 1), GetParam(kvp.Value, 2), 0, 0)
                            LoggingModule.Write(String.Format("Sent pumpsrv.Reconfigure2({0},{1},{2},{3})to PumpSrv COM", GetParam(kvp.Value, 1), GetParam(kvp.Value, 2), 0, 0))
                        Else
                            LoggingModule.Write(String.Format("Reconfigure table {0} did not return valid integer values.", kvp.Key), MyColor:=ConsoleColor.Red)
                        End If
                    Else
                        LoggingModule.Write(String.Format("Table value returned is null. Reconfigure for {0} will not occur.", table), MyColor:=ConsoleColor.Red)
                    End If

                    Threading.Thread.Sleep(500)
                Next kvp
            Else
                Dim tableIsNumeric As Boolean = False
                configVars = table.Split(","c)
                Dim param1, param2, param3, param4 As Integer
                If IsNumeric(configVars(0)) Then tableIsNumeric = True

                LoggingModule.Write(String.Format("Returning PumpSrv parameters for {0}.", configVars(0)))
                If tableIsNumeric Then
                    value = table
                    LoggingModule.Write(String.Format("Value appears to be numeric. Using {0} for reconfigure unchanged.", value))
                Else
                    reconfigTable.TryGetValue(configVars(0), value) 'Attempt to read back params for table 
                    LoggingModule.Write(String.Format("Value returned =  {0}.", value))
                End If
                param1 = GetParam(value, 1, tableIsNumeric)
                param2 = GetParam(value, 2, tableIsNumeric)
                param3 = GetParam(value, 3, tableIsNumeric)
                param4 = GetParam(value, 4, tableIsNumeric)

                'override params if passed with command line, primarily used as overrides with table name + param
                If configVars.Length = 2 Then
                    Integer.TryParse(configVars(1), param2)
                ElseIf configVars.Length = 3 Then
                    Integer.TryParse(configVars(1), param2)
                    Integer.TryParse(configVars(2), param3)
                ElseIf configVars.Length = 4 Then
                    Integer.TryParse(configVars(1), param2)
                    Integer.TryParse(configVars(2), param3)
                    Integer.TryParse(configVars(3), param4)
                End If

                If value <> "" Then 'If params are returned and not -2 which is unsupported
                    If IsValidReconfig(String.Format("{0},{1},{2},{3}", param1, param2, param3, param4)) Then
                        pumpsrv.ReConfigure2(param1, param2, param3, param4)
                        LoggingModule.Write(String.Format("Sent pumpsrv.Reconfigure2({0},{1},{2},{3})to PumpSrv COM", param1, param2, param3, param4))
                    Else
                        LoggingModule.Write(String.Format("Reconfigure table {0} did not return valid integer values.", configVars(0)), MyColor:=ConsoleColor.Red)
                    End If
                Else
                    LoggingModule.Write(String.Format("Table index was not found. Reconfigure for {0} will not occur.", configVars(0)), MyColor:=ConsoleColor.Red)
                End If
            End If
            pumpsrv = Nothing

        Catch ex As Exception
            LoggingModule.Write("Error loading PumpSrv object. Unable to reconfigure.", MyColor:=ConsoleColor.Red)
        End Try

    End Sub



    Private Function FMSStartStop(ByVal StartStop As String) As Integer
        LoggingModule.Write("Entering FMS Start/Stop routine.")
        LoggingModule.Write(String.Format("Mode = {0}", StartStop))
        Try
            Dim pumpsrv = New PUMPSRVLib.Ctrl()
            Dim resp As Integer
            If StartStop.ToLower = "stop" Then
                LoggingModule.Write("Stopping FuelMobileSrv process via PumpSrv.")
                resp = pumpsrv.StopService(2)
            ElseIf StartStop.ToLower = "start" Then
                LoggingModule.Write("Starting FuelMobileSrv process via PumpSrv.")
                resp = pumpsrv.StartService(2)
            Else
                LoggingModule.Write("Unsupported FMS command.")
            End If

            pumpsrv = Nothing
        Catch ex As Exception
            LoggingModule.Write("FMSStartStop Error: Error loading PumpSrv object. Unable to reconfigure.", MyColor:=ConsoleColor.Red)
            Return -1
        End Try

    End Function

    Private Function GetFuelDbTables() As List(Of String)
        Dim GetTables As String = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_CATALOG='FUEL'"
        Dim Tables As New List(Of String)
        Dim CurrentTable As String
        myCmd = myConn.CreateCommand

        myCmd.CommandText = GetTables
        myCmd.Connection = myConn
        'Open the connection.
        Try
            myConn.Open()

            myReader = myCmd.ExecuteReader()
            Do While myReader.Read()
                CurrentTable = myReader.GetString(0)
                Tables.Add(CurrentTable)
            Loop
            myReader.Close()
            myConn.Close()
        Catch ex As Exception
            LoggingModule.Write("GetFuelDbTables Error: Unable to generate fuel database SQL tables list.", MyColor:=ConsoleColor.Red)
            LoggingModule.Write(String.Format("Message: {0}", ex.Message), MyColor:=ConsoleColor.Red)
        End Try
        Tables.TrimExcess()

        Return Tables

    End Function

    Private Sub EmptyFuelDbTables()
        Dim Tables As New List(Of String)
        Tables = GetFuelDbTables()
        Dim CurrentCommand As String
        Dim RecAffected As Integer

        Dim TableCount As Integer = Tables.Count

        If TableCount > 0 Then
            myCmd = myConn.CreateCommand
            myCmd.Connection = myConn
            Try
                'Open the connection.
                myConn.Open()

                For Each Table In Tables
                    CurrentCommand = "DELETE FROM " & Table
                    myCmd.CommandText = CurrentCommand
                    RecAffected = myCmd.ExecuteNonQuery()
                    LoggingModule.Write(String.Format("{0} ({1} rows affected)", CurrentCommand, RecAffected))
                Next

                myConn.Close()

            Catch ex As Exception
                LoggingModule.Write("EmptyFuelDbTables Error: Unable to empty fuel SQL database tables.", MyColor:=ConsoleColor.Red)
                LoggingModule.Write(String.Format("Message: {0}", ex.Message), MyColor:=ConsoleColor.Red)
            End Try


        Else
            LoggingModule.Write("No SQL database tables were returned.", MyColor:=ConsoleColor.Red)
        End If

    End Sub

End Module
