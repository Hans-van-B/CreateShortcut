Imports System.IO, IWshRuntimeLibrary

Module Module1
    Dim Trace As Integer = 1
    Dim Translate As Boolean = False
    Dim TSet As Boolean = False
    Dim CreateDir As Boolean = False
    Dim WindowStyle As String = ""
    Dim WorkDir As String = ""
    Dim Description As String = ""
    Dim Arguments As String = ""
    Dim IconLoc As String = ""
    Dim Key As String = ""


    '---- Assign special folders ---------------------------------------------------------------------------

    Function AssignLoc(Loc As String)
        xtrace_subs("AssignLoc " & Loc)
        Dim Result As String = "Unknown"
        If Loc = "CDT" Then Result = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory)
        If Loc = "CSM" Then Result = Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu)
        If Loc = "CSU" Then Result = Environment.GetFolderPath(Environment.SpecialFolder.CommonStartup)
        If Loc = "DTD" Then Result = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        If Loc = "DT" Then Result = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        If Loc = "SM" Then Result = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)
        If Loc = "NS" Then Result = Environment.GetFolderPath(Environment.SpecialFolder.NetworkShortcuts)
        If Loc = "SU" Then Result = Environment.GetFolderPath(Environment.SpecialFolder.Startup)
        If Loc = "PS" Then Result = Environment.GetFolderPath(Environment.SpecialFolder.PrinterShortcuts)
        If Loc = "W" Then Result = Environment.GetFolderPath(Environment.SpecialFolder.Windows)

        If Loc = "CSMP" Then Result = Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu) & "\" & "Programs"
        If Loc = "SMP" Then Result = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu) & "\" & "Programs"

        If Len(Loc) > 4 Then Result = Loc
        AssignLoc = Result

        xtrace_i("Result = " & Result)
        xtrace_sube("AssignLoc")
    End Function

    '==== Main =============================================================================================

    Sub Main()
        xtrace_init()
        Read_Command_Line_Arg()
        If ExitProgram Then Exit Sub

        '---- Read Command line
        Dim args() As String = System.Environment.GetCommandLineArgs()
        Dim Name As String = "Shortcut" ' Shortcut Name
        Dim Exe As String = "Undefined"  ' Executable
        Dim LocId As String = "C:\Temp"
        Dim Loc As String
        Dim SubDir As String = ""
        Dim PN As String    ' Paramater Name
        Dim PV As String    ' Paramater Value
        Dim IsPos As Integer

        xtrace("-- Read Arguments --")
        For Each arg As String In args
            xtrace("Argument=" & arg, 2)
            IsPos = InStr(arg, "=")
            XTrace("IsPos = " & IsPos.ToString, 3)

            ' Switches (No values)
            If IsPos < 1 Then
                XTrace("Arguments 1", 3)

                If arg = "-c" Then CreateDir = True
                If arg = "-d" Then
                    XTrace("--Dryrun--", 1)
                    DryRun = True
                End If
                If arg = "-t" Then Translate = True
                If arg = "--ts" Then TSet = True
                If arg = "-v" Then Trace = 2
                If arg = "--vv" Then Trace = 3
                If arg = "-s" Then Trace = 0
            End If

            If IsPos = 2 Then
                XTrace("Arguments 2", 3)

                PN = Left(arg, 1)
                PV = Mid(arg, 3)
                XTrace(PN & " = " & PV, 3)

                If PN = "N" Then Name = PV.Replace("""", "")
                If PN = "E" Then Exe = PV.Replace("""", "")
                If PN = "L" Then LocId = PV.Replace("""", "")
                If PN = "S" Then SubDir = PV.Replace("""", "")
                If PN = "D" Then Description = PV.Replace("""", "")
                If PN = "A" Then Arguments = PV  '.Replace("""", "")      Remove quotes disabled!  2022-02-03
                If PN = "I" Then IconLoc = PV.Replace("""", "")
                If PN = "K" Then Key = PV.Replace("""", "")
                If PN = "W" Then WorkDir = PV.Replace("""", "")
            End If

            ' 4+= letter names + Value
            If IsPos > 2 Then
                XTrace("Arguments 3", 3)

                PN = Left(arg, IsPos - 1)
                PV = Mid(arg, 6)
                XTrace(PN & " = " & PV, 3)
                If PN = "--ws" Then WindowStyle = LCase(PV)
                If PN = "--wd" Then WorkDir = PV.Replace("""", "") ' Obsolete
            End If

            XTrace(" ", 3)  ' End argument
        Next

        Loc = AssignLoc(LocId)    ' Assign special folders
        If SubDir <> "" Then Loc = Loc & "\" & SubDir

        '---- Translate ---------------------------------------------------------

        If Translate Then
            XTrace(Loc, 1)
            Exit Sub
        End If

        ' This does not work, you may consider creating a bat file for external use but it is easier to use -t
        If TSet Then
            XTrace("SET " & LocId & "=" & Loc, 1)
            Environment.SetEnvironmentVariable(LocId, Loc)
            Exit Sub
        End If

        '---- Create the directory -----------------------------------------------
        XTrace("Name     = " & Name, 2)
        XTrace("Exe      = " & Exe, 2)
        XTrace("Location = " & Loc, 2)
        If WindowStyle <> "" Then XTrace("WindowStyle  = " & WindowStyle, 2)
        If Description <> "" Then XTrace("Description  = " & Description, 2)
        If Arguments <> "" Then xtrace("Arguments    = " & Arguments, 2)

        If DryRun Then
            xtrace_i("Exit on dryrun")
            GoTo QUIT
        Else
            xtrace_i("Proceed")
        End If

        If (Not System.IO.Directory.Exists(Loc)) Then

            If CreateDir Then
                XTrace("Create directory: " & Loc, 1)
                If Not DryRun Then System.IO.Directory.CreateDirectory(Loc)
            Else
                XTrace("Warning: The directory """ & Loc & """ Does not exist", 0)
                Exit Sub
            End If
        End If

        '---- Create the shortcut ------------------------------------------------

        If Exe = "Undefined" Then
            XTrace("Warning: The executable path is not supplied", 0)
            Exit Sub

        End If

        Dim ShortcutPath As String
        Dim MyWshShell As WshShell = New WshShell
        Dim MyShortcut As IWshRuntimeLibrary.IWshShortcut

        ShortcutPath = Loc & "\" & Name & ".lnk"
        XTrace("ShortcutPath = " & ShortcutPath, 2)

        XTrace("Create """ & ShortcutPath & """ As """ & Exe & """", 1)
        If DryRun Then Exit Sub

        MyShortcut = CType(MyWshShell.CreateShortcut(ShortcutPath), IWshRuntimeLibrary.IWshShortcut)
        MyShortcut.TargetPath = Exe 'Specify target app full path

        If WindowStyle = "min" Then
            XTrace("Set WindowsStyle Minimized", 1)
            'MyShortcut.WindowStyle = vbMinimizedNoFocus
            MyShortcut.WindowStyle = 7  ' 1 = normal window, 3 = maximized window, 7 = minimized window
        End If

        If Description <> "" Then
            MyShortcut.Description = Description
        End If

        If Arguments <> "" Then
            MyShortcut.Arguments = Arguments
        End If

        If WorkDir <> "" Then
            MyShortcut.WorkingDirectory = WorkDir
        End If

        If IconLoc = "" Then
            MyShortcut.IconLocation = Exe
        Else
            XTrace("Set Icon", 1)
            MyShortcut.IconLocation = IconLoc
        End If

        If Key <> "" Then
            XTrace("Set Shortcut Key", 1)
            MyShortcut.Hotkey = Key
        End If

        MyShortcut.Save()

QUIT:
        If WaitForKey Then AnyKey()
    End Sub

End Module
