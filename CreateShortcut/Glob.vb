Module Glob
    ' Public GitHub
    Public AppName As String = "CreateShortcut"
    Public AppVer As String = "0.01.03"

    Public AppRoot As String = IO.Path.GetDirectoryName(Diagnostics.Process.GetCurrentProcess().MainModule.FileName)
    Public CD As String = My.Computer.FileSystem.CurrentDirectory

    ' Log
    Public Temp As String = Environment.GetEnvironmentVariable("TEMP")
    Public LogFile As String = Temp & "\" & AppName & ".log"
    Public Verbose As Boolean = False
    Public CTrace As Integer = 1                       ' Console Trace Level
    Public LTrace As Integer = 2
    Public G_TL_Sub As Integer = 2  ' Trace level for Subroutines

    Public SubLevel As Integer = 0
    Public ErrorCount As Integer = 0
    Public ExitProgram As Boolean = False

    ' Defaults
    Public Debug As Boolean = False
    Public WaitForKey As Boolean = False

    Public DryRunStr As String = "CreateShortCutDryrun"
    Public DryRun As Boolean = False
End Module
