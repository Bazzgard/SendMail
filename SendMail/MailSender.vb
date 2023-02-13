Imports System.Net.Mail
Imports System.Xml

Public Class MailSender

    Public Function SendMail(xmlFilePath As String) As Boolean
        Try
            Dim xmlDoc As New Xml.XmlDocument()
            xmlDoc.Load(xmlFilePath)
            Dim emailNode As Xml.XmlNode = xmlDoc.SelectSingleNode("/Email")

            Dim fromAddress As String = emailNode.SelectSingleNode("From").InnerText
            Dim toAddresses As String() = emailNode.SelectNodes("To").OfType(Of Xml.XmlNode)().Select(Function(x) x.InnerText).ToArray()
            Dim ccAddresses As String() = emailNode.SelectNodes("CC").OfType(Of Xml.XmlNode)().Select(Function(x) x.InnerText).ToArray()
            Dim bccAddresses As String() = emailNode.SelectNodes("BCC").OfType(Of Xml.XmlNode)().Select(Function(x) x.InnerText).ToArray()
            Dim subject As String = emailNode.SelectSingleNode("Subject").InnerText
            Dim body As String = emailNode.SelectSingleNode("Body").InnerText

            Dim server As String = emailNode.SelectSingleNode("smtp/server").InnerText
            Dim port As Integer = Integer.Parse(emailNode.SelectSingleNode("smtp/port").InnerText)
            Dim username As String = emailNode.SelectSingleNode("smtp/username").InnerText
            Dim password As String = emailNode.SelectSingleNode("smtp/password").InnerText

            ' メール作成
            Using client As New SmtpClient(server, port)
                client.Credentials = New Net.NetworkCredential(username, password)
                client.EnableSsl = True
                Using message As New MailMessage()
                    message.From = New MailAddress(fromAddress)
                    For Each toAddress As String In toAddresses
                        message.To.Add(toAddress)
                    Next
                    For Each ccAddress As String In ccAddresses
                        message.CC.Add(ccAddress)
                    Next
                    For Each bccAddress As String In bccAddresses
                        message.Bcc.Add(bccAddress)
                    Next
                    message.Subject = subject
                    message.Body = body
                    ' メール送信
                    client.Send(message)
                End Using
            End Using

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
End Class
