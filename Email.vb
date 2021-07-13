Imports System.Net.Mail





Module Email


    Function SendEmail(eToEmail As String, eSubject As String, eBody As String, eFrom As String, eToMob As String) As Boolean
        Dim i As Integer = 1

        SendEmail = False
        eToMob = eToMob.Replace(" ", String.Empty)

        If frmMain.chkRunTest.Checked = True Then
            eToMob = frmMain.txtTestMobile.Text
            eToMob = eToMob.Replace(" ", String.Empty)
        End If

        Dim sSubject As String = "Shipment Notice:  " & eSubject
        Dim sToMob As String = eToMob & "@e2s.messagemedia.com"

        'sToMob = "0404644548@e2s.messagemedia.com"


        ' Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        ' Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "niagara-com-au.Mail.protection.outlook.com"
        ' Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        ' Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
        ' Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendtls") = False
        ' Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True

        Dim Smtp_Server As New SmtpClient
        Dim e_mail As New MailMessage()

        Smtp_Server.UseDefaultCredentials = False
        Smtp_Server.Credentials = New Net.NetworkCredential(frmMain.smtpUser, frmMain.smtpPswd)
        Smtp_Server.DeliveryMethod = SmtpDeliveryMethod.Network
        Smtp_Server.Port = frmMain.smtpPort
        Smtp_Server.EnableSsl = frmMain.smtpSSL
        Smtp_Server.Host = frmMain.smtpHost

        ' SMS
        Try
            e_mail = New MailMessage()
            e_mail.From = New MailAddress(frmMain.smtpUser)
            e_mail.To.Add(sToMob)
            e_mail.Subject = sSubject
            e_mail.IsBodyHtml = True
            e_mail.Body = eBody
            Smtp_Server.Send(e_mail)
            SendEmail = True





        Catch ex As Exception
            SendEmail = False




        End Try


    End Function




End Module
