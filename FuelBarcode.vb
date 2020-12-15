Imports System.IO
Imports Microsoft.Win32
Imports System.Data.Odbc

Public Class FuelBarcode

    Public dbDrive As String
    Public dbExt As String
    Public dbDriver As String
    Private Function GetRegBasePath() As String
        If Environment.Is64BitOperatingSystem Then
            Return "Software\WOW6432Node"
        Else
            Return "Software"
        End If
    End Function

    Private Sub SetDbVars()
        If Environment.Is64BitOperatingSystem Then
            dbDrive = "C"
            dbExt = "ib"
            dbDriver = "Interbase ODBC driver"
        Else
            dbDrive = "D"
            dbExt = "gdb"
            dbDriver = "Firebird/Interbase(r) driver"
        End If
    End Sub

    Private Sub LoadGrades()
        Dim gradeKey, wsKey As RegistryKey
        Dim basePath As String = GetRegBasePath()
        Dim sGrade, sFullName, sShortName3, sShortName5, sAdditionalCode As String
        Dim iCode, iConvertCode, iWSGrade, iWSBarcodeGrade, iValid As Integer
        For Grade As Integer = 1 To 76
            sGrade = Grade.ToString.PadLeft(2, "0"c)
            Try
                gradeKey = Registry.LocalMachine.OpenSubKey(String.Format("{0}\PointOfSale\PumpSrv\Grades\Grade{1}", basePath, sGrade), True)
                If gradeKey Is Nothing Then
                    gradeKey.CreateSubKey(String.Format("{0}\PointOfSale\PumpSrv\Grades\Grade{1}", basePath, sGrade))
                End If
                iValid = CInt(gradeKey.GetValue("ValidEntry", 0))
                If iValid = 0 Then Continue For
                iCode = CInt(gradeKey.GetValue("Code", 0))
                iConvertCode = CInt(gradeKey.GetValue("ConvertCode", 0))
                sFullName = CStr(gradeKey.GetValue("FullName", ""))
                sShortName3 = CStr(gradeKey.GetValue("ShortName3", ""))
                sShortName5 = CStr(gradeKey.GetValue("ShortName5", ""))
                sAdditionalCode = CStr(gradeKey.GetValue("AdditionalCode", ""))

                gradeKey.Close()

            Catch ex As Exception
                LoggingModule.Write(String.Format("Unable to read {0}\PointOfSale\PumpSrv\Grades\Grade{1}", basePath, sGrade), MyColor:=ConsoleColor.Red)
            End Try

            Try
                wsKey = Registry.LocalMachine.OpenSubKey(String.Format("{0}\Positive\Terminal99\WeStock", basePath), True)
                iWSGrade = CInt(wsKey.GetValue(String.Format("Grade{0}", Grade), 8000))
                iWSBarcodeGrade = CInt(wsKey.GetValue(String.Format("Barcodegrade{0}", Grade), 8000))
                wsKey.Close()

            Catch ex As Exception
                LoggingModule.Write(String.Format("Unable to read {0}\Positive\Terminal99\WeStock registry key.", basePath), MyColor:=ConsoleColor.Red)
            End Try
        Next

    End Sub

    Private Sub LoadWashes()
        Dim washKey, wsKey As RegistryKey
        Dim basePath As String = GetRegBasePath()
        Dim sFullName, sWash As String
        Dim iValid, iPrice, iWSWash, iWSWashBarcode As Integer
        For Wash As Integer = 1 To 8
            sWash = Wash.ToString.PadLeft(2, "0"c)
            Try
                washKey = Registry.LocalMachine.OpenSubKey(String.Format("{0}\PointOfSale\PumpSrv\WashPrograms\WashProgram{1}", basePath, sWash), True)
                If washKey Is Nothing Then
                    washKey.CreateSubKey(String.Format("{0}\PointOfSale\PumpSrv\WashPrograms\WashProgram{1}", basePath, sWash))
                End If
                iValid = CInt(washKey.GetValue("ValidEntry", 0))
                If iValid = 0 Then Continue For
                iPrice = CInt(washKey.GetValue("Price", 0))
                sFullName = CStr(washKey.GetValue("FullName", ""))

                washKey.Close()

            Catch ex As Exception
                LoggingModule.Write(String.Format("Unable to read {0}\PointOfSale\PumpSrv\WashPrograms\WashProgram{1}", basePath, sWash), MyColor:=ConsoleColor.Red)
            End Try

            Try
                wsKey = Registry.LocalMachine.OpenSubKey(String.Format("{0}\Positive\Terminal99\WeStock", basePath), True)
                iWSWash = CInt(wsKey.GetValue(String.Format("WashProgram{0}", sWash), 8000))
                iWSWashBarcode = CInt(wsKey.GetValue(String.Format("BarcodeWashProgram{0}", sWash), 8000))
                wsKey.Close()

            Catch ex As Exception
                LoggingModule.Write(String.Format("Unable to read {0}\Positive\Terminal99\WeStock registry key.", basePath), MyColor:=ConsoleColor.Red)
            End Try
        Next

    End Sub

    Private Sub QueryTenders()
        Dim query As String = "select * from EMPLOYEE"
        Dim cnn As New Odbc.OdbcConnection()
        Dim estring As New Odbc.OdbcConnectionStringBuilder("DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA;PWD=masterkey;DBNAME=128.1.7.81:C:\office\db\office.gdb;")
        cnn = New OdbcConnection(estring.ToString)

        Dim da As New OdbcDataAdapter("select * from EMPLOYEE", estring.ToString)
        Dim ds As New DataSet
        Dim dt As New DataTable

        Try
            cnn.Open()
            da.Fill(dt)
            cnn.Close()
            cnn.Dispose()
        Catch

        End Try

    End Sub

End Class
