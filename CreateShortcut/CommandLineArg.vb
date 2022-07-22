Module CommandLineArg
    Sub Read_Command_Line_Arg()
        xtrace_subs("Read_Command_Line_Arg")

        Dim SwName As String = ""
        Dim SwString As String = ""
        Dim P1 As Integer
        Dim Name As String
        Dim ValS As String

        For Each argument As String In My.Application.CommandLineArgs
            xtrace_i("Argument=" & argument)

            ' WARNING: Command line arguments are also read in MAIN !!!!!
            '---- Double-dash arguments
            If Left(argument, 2) = "--" Then
                SwName = Mid(argument, 3)
                xtrace_i("DDA:" & SwName)

                '---- Double-dash Assignment
                P1 = InStr(SwName, "=")
                If P1 > 0 Then
                    Name = Left(SwName, P1 - 1)
                    ValS = Mid(SwName, P1 + 1)
                    xtrace_i("Name = " & Name)

                    If Name = "??" Then

                    End If

                    Continue For
                End If

                '---- Double-dash No Assignment
                If SwName = "help" Then
                    ShowHelp()
                    ExitProgram = True
                End If

                If SwName = "xhelp" Then
                    xtrace("xhelp is not supported!", 0)
                    ExitProgram = True
                End If

                Continue For
            End If

            '---- Single-dash arguments
            If Left(argument, 1) = "-" Then
                ' Switch String = remaining switches
                SwString = Mid(argument, 2)

                ' for each switch in the string
                While Len(SwString) > 0
                    SwName = Left(SwString, 1)
                    SwString = Mid(SwString, 2)
                    xtrace_i("SDA:" & SwName & "," & SwString)

                    If SwName = "v" Then
                        xtrace_i("Set verbose = True")
                        Console.WriteLine(" Log file = " & LogFile)
                        Verbose = True
                    End If

                    If SwName = "w" Then
                        xtrace_i("Set WaitForKey = True")
                        WaitForKey = True
                    End If

                    If SwName = "h" Then
                        ShowHelp()
                        ExitProgram = True
                    End If

                    If SwName = "d" Then
                        xtrace_i("Set Debug = True")
                        Debug = True
                        Verbose = True
                    End If
                End While

                Continue For
            End If

            '---- Else (No dashes)
            P1 = InStr(argument, "=")
            If P1 > 0 Then
                Name = Left(argument, P1 - 1)
                ValS = Mid(argument, P1 + 1)
                xtrace_i("NDA:" & argument)

                If Name = "??" Then

                End If
            End If

            '---- Else (No dashes, No Assign)
            ' Wait <Ini>
            ' for compatibility with the older wait command
            If (argument = "/?") Or (argument = "/h") Then
                ShowHelp()
                ExitProgram = True
                Continue For
            End If

        Next

        xtrace_i("Check Environment Variables")

        If Environment.GetEnvironmentVariable("DEBUG") = "TRUE" Then
            xtrace_i("Set by environment Debug = True")
            Debug = True
            Verbose = True
        End If

        ValS = Environment.GetEnvironmentVariable(DryRunStr)    ' Also see the -d switch
        If ValS = "TRUE" Then
            xtrace(DryRunStr & " = TRUE", 1)
            DryRun = True
        Else
            xtrace_i(DryRunStr & " = """ & ValS & """")
        End If

        xtrace_sube("Read_Command_Line_Arg")
    End Sub

End Module
