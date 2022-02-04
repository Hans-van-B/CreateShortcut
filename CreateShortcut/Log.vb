Module Log
    Sub xtrace_init()
        ' Show the command-line string at the top of the log file
        My.Computer.FileSystem.WriteAllText(LogFile, Environment.CommandLine & vbNewLine, False)
        xtrace("xtrace_init")
        xtrace(AppName & " V" & AppVer)
        xtrace_TimeStamp()

        xtrace_i("AppRoot   = " & AppRoot)
        'xtrace_i("Ini       = " & IniFile)
        xtrace_i("Log level to con     = " & CTrace.ToString, 2)
        xtrace_i("Log level to logfile = " & LTrace.ToString, 2)
    End Sub

    '---- xtrace ----
    Sub xtrace_root(Msg As String, TV As Integer)
        Dim Nr As Int16
        Dim Tab As String = ""

        ' If subroutines are Not logged then tabbing is also disabeled
        If LTrace >= G_TL_Sub Then
            ' If exiting main or sublevel maint error,
            ' this if makes it more clear
            If SubLevel >= 0 Then Tab = "|"

            For Nr = 1 To SubLevel
                Tab = Tab + "  |"
            Next
        End If

        If TV <= CTrace Then
            Console.Write(Msg & vbCrLf)
        End If

        If TV <= LTrace Then
            My.Computer.FileSystem.WriteAllText(LogFile, Tab + Msg & vbCrLf, True)
        End If
    End Sub

    Sub xtrace(Msg As String)
        xtrace_root(" " & Msg, 2)
    End Sub

    Sub xtrace(Msg As String, TV As Integer)
        xtrace_root(" " & Msg, TV)
    End Sub

    '---- xtrace_i ----
    Sub xtrace_i(Msg As String)
        xtrace(" * " & Msg)
    End Sub

    Sub xtrace_i(Msg As String, TV As Integer)
        xtrace(" * " & Msg, TV)
    End Sub

    '---- xtrace_err
    Sub xtrace_err(Msg As String)
        xtrace("ERROR: " & Msg, 1)
        ErrorCount = ErrorCount + 1
    End Sub

    '---- xtrace_err multi-line
    Sub xtrace_err(Msg() As String)
        Dim Line1 As Boolean = True
        For Each Line As String In Msg
            If Line1 Then
                xtrace("ERROR: " & Line, 1)
                Line1 = False
            Else
                xtrace("       " & Line, 1)
            End If
        Next
        ErrorCount = ErrorCount + 1
    End Sub

    '--- xtrace TimeStamp ----
    Sub xtrace_TimeStamp()
        xtrace_i("Timestamp = " & DateTime.Now)
    End Sub

    '--- xtrace Sub ----
    Sub xtrace_subs(Msg As String)
        xtrace_root("->" & Msg & " (" & (SubLevel + 1).ToString & ")", G_TL_Sub)
        SubLevel = SubLevel + 1
    End Sub

    Sub xtrace_sube(Msg As String)
        SubLevel = SubLevel - 1
        xtrace_root("<-" & Msg & " (" & (SubLevel + 1).ToString & ")", G_TL_Sub)
    End Sub

    Sub xtrace_sube()
        xtrace_sube("<NoName>")
    End Sub
End Module
