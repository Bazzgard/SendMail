Imports System.Configuration
Public Class LogUtil
    Shared sLogPath As String = ConfigurationManager.AppSettings("LogPath")

    Public Shared Sub logWrite(s As String)
        Using w As New System.IO.StreamWriter(sLogPath, True)
            w.WriteLine("[" & DateTime.Now.ToString() & "]" & s)
        End Using
    End Sub

    Public Shared Sub logError(ex As Exception)
        Using w As New System.IO.StreamWriter(sLogPath, True)
            w.WriteLine("[" & DateTime.Now.ToString() & "]" & ex.Message)
            w.WriteLine(ex.StackTrace)
        End Using
    End Sub

End Class
