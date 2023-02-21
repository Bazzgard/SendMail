Imports System.Configuration

Module Module1

    Sub Main()
        Dim sMailSettingsFilePath = ConfigurationManager.AppSettings("MailSettingsFilePath")
        Try

            Dim mailSender As New MailSender()
            mailSender.SendMail(sMailSettingsFilePath)

            LogUtil.logWrite("メールが送信されました。")
        Catch ex As Exception
            LogUtil.logError(ex)
        End Try

    End Sub

End Module
