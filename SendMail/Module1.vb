Module Module1

    Sub Main()
        Dim mailSender As New MailSender()
        Dim result As Boolean = mailSender.SendMail("mail.xml")
        If result Then
            Console.WriteLine("Email sent successfully.")
        Else
            Console.WriteLine("Failed to send email.")
        End If
    End Sub

End Module
