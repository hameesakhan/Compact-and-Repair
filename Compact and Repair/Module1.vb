Module Module1
    Public DBTempPath, DBTempFile, DBPass As String
    Sub Main()
        Try
            My.Computer.Registry.ClassesRoot.CreateSubKey("*").CreateSubKey("shell").CreateSubKey("Compact using CnR").CreateSubKey("command").SetValue("", My.Application.Info.DirectoryPath & "\CompactandRepair.exe %1")
            My.Computer.Registry.ClassesRoot.CreateSubKey("*").CreateSubKey("shell").CreateSubKey("Compact using CnR").SetValue("icon", "(" & My.Application.Info.DirectoryPath & "\CompactandRepair.exe,0)")
        Catch
            Console.WriteLine("Cannot register for Context Menu. Please try running application as Administrator.")
        End Try

        Dim DBPath As String
        Console.WriteLine("This tool only works for accdb and mdb jet databases. Please backup your databases before running this utility on them.")
        Console.WriteLine("DISCLAIMER: There is no warranty of this product.")
        Console.WriteLine("Press Ctrl-C to quit.")
        Console.WriteLine("================================================================================")
        Console.WriteLine()
        Console.WriteLine("No of arguments: " & My.Application.CommandLineArgs.Count)
        Dim i As Integer
        DBPath = ""
        Do While i < My.Application.CommandLineArgs.Count
            DBPath = DBPath + " " + My.Application.CommandLineArgs.Item(i)
            i = i + 1
        Loop
        If My.Computer.FileSystem.FileExists(DBPath) Then GoTo SkipREPL

        '         REPL - Read Evaluate Print Loop
REPL:
        Console.Write("Enter path to database: ")
        DBPath = Console.ReadLine()

        If Not My.Computer.FileSystem.FileExists(DBPath) Then Console.WriteLine("We are unable to find the specified database: " + DBPath) : GoTo REPL

SkipREPL:
        Console.WriteLine("Database: " & DBPath)
        DBTempFile = My.Computer.FileSystem.GetName(My.Computer.FileSystem.GetTempFileName()) & My.Computer.FileSystem.GetFileInfo(DBPath).Extension
        KillTempDB(DBTempFile)
        If SendToTempFolder(DBPath) Then

            DBPass = ReturnPassOfDB(My.Computer.FileSystem.GetName(DBPath))
            DoLog("Step 1 of 5")
            RemovePassFromDB(DBTempPath, DBPass)

            DoLog("Step 2 of 5")
            CompactAndRepairDB(DBTempPath)

            DoLog("Step 3 of 5")
            SetPassToDB(DBTempPath, DBPass)

            DoLog("Step 4 of 5")
            ReceiveFromTempFolder(DBPath)

            DoLog("Step 5 of 5")
            KillTempDB(DBTempFile)
        End If
        Console.WriteLine("Please <Enter> to quit.")
        Console.Read()
    End Sub
    Function CompactAndRepairDB(ByVal DBPath As String) As String
        On Error GoTo EndErr
        Dim cnn As New Microsoft.Office.Interop.Access.Dao.DBEngine()
        cnn.CompactDatabase(DBPath, DBPath & "BACKUP")
        Kill(DBPath)
        Rename(DBPath & "BACKUP", DBPath)
        Exit Function
EndErr:
        Console.WriteLine("CAR: " & Err.Description)
    End Function
    Function RemovePassFromDB(ByVal DbPath As String, ByVal DbPassword As String) As String
        On Error GoTo EndErr
        Dim cnn As New ADODB.Connection
        cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & DbPath & "';"
        If DbPassword <> "" Then cnn.ConnectionString = cnn.ConnectionString & " Jet OLEDB:Database Password=" & DbPassword & ";"
        cnn.Mode = ADODB.ConnectModeEnum.adModeShareExclusive
        cnn.Open()
        If DbPassword = "" Then DbPassword = "NULL" Else DbPassword = "[" & DbPassword & "]"
        cnn.Execute("ALTER DATABASE PASSWORD NULL " & DbPassword & ";")
        cnn.Close()
        Exit Function
EndErr:
        If cnn.State <> 0 Then cnn.Close()
        Console.WriteLine("RPTD: " & Err.Description)
    End Function
    Function SetPassToDB(ByVal DbPath As String, ByVal DbPassword As String)
        On Error GoTo EndErr
        Dim cnn As New ADODB.Connection
        cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & DbPath & "';"
        If DbPassword <> "" Then cnn.ConnectionString = cnn.ConnectionString & " Jet OLEDB:Database Password=" & DbPassword & ";"
        cnn.Mode = ADODB.ConnectModeEnum.adModeShareExclusive
        cnn.Open()
        If DbPassword = "" Then DbPassword = "NULL" Else DbPassword = "[" & DbPassword & "]"
        cnn.Execute("ALTER DATABASE PASSWORD " & DbPassword & " NULL;")
        cnn.Close()
        Exit Function
EndErr:
        If cnn.State <> 0 Then cnn.Close()
        Console.WriteLine("SPTB: " & Err.Description)
    End Function
    Function DoLog(ByVal Text As String)
        Console.WriteLine(Text)
    End Function
    Function ReturnPassOfDB(ByVal DBName As String) As String
        Dim DbFileName As String = My.Computer.FileSystem.GetName(DBName)
        DbFileName = DbFileName.ToLower
        'Use this Select-Case to predefine passwords for your databases.
        '
        Select Case DbFileName
            Case "SampleDBName"
                Return "ThisPassword"
            Case Else
                Console.Write("Unable to find database password. Please enter database password. You may leave it blank for non-password databases:")
                Return Console.ReadLine
        End Select
    End Function
    Function SendToTempFolder(ByVal PathStr As String) As Boolean
        On Error GoTo err_occured
        DBTempPath = My.Computer.FileSystem.SpecialDirectories.Temp & "\" & DBTempFile
        My.Computer.FileSystem.CopyFile(PathStr, DBTempPath, True)
        Return True

        Exit Function
err_occured:
        Console.WriteLine(Err.Number & ": " & Err.Description)
        Return False
    End Function
    Function ReceiveFromTempFolder(ByVal PathStr As String) As Boolean
        On Error GoTo err_occured

        My.Computer.FileSystem.CopyFile(DBTempPath, PathStr, True)
        Return True

        Exit Function
err_occured:
        Console.WriteLine(Err.Number & ": " & Err.Description)
    End Function
    Function KillTempDB(ByVal FileName As String)
        If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.SpecialDirectories.Temp & "\" & FileName) Then Kill(My.Computer.FileSystem.SpecialDirectories.Temp & "\" & FileName)
    End Function

End Module