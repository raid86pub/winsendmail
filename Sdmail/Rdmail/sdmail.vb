Imports System.Threading.Thread
Imports System.IO

Module sdmail
        Const IDX_TOTO = 0
        Const IDX_FILE = 1
        Const IDX_GREP = 2
        Const IDX_HOST = 3
        Const IDX_USER = 4
        Const IDX_PSWD = 5
        Dim gMs() As String = {"rdstk01@sina.com", "ABCDE", "VBCMD LOGIN", "smtp.sina.com", "rdstk01@sina.com", "Lrf@12345"}

        '  The mailman object is used for receiving (POP3)
        '  and sending (SMTP) email.
        Dim mailman As New Chilkat.MailMan()

        Dim gBody(10) As String
        Dim r As Random = New Random

        Sub InitBody()
                gBody(0) = "Hi Guys," & vbCrLf
                gBody(0) &= "Stand ye calm and resolute," & vbCrLf
                gBody(0) &= "Like a forest close and mute," & vbCrLf
                gBody(0) &= "Thanks," & vbCrLf
                gBody(0) &= "-Ried"

                gBody(1) = "Hi Larry," & vbCrLf
                gBody(1) &= "With folded arms and looks which are" & vbCrLf
                gBody(1) &= "Weapons of unvanquished war." & vbCrLf
                gBody(1) &= "Thanks," & vbCrLf
                gBody(1) &= "-Ried"

                gBody(2) = "Hi Man," & vbCrLf
                gBody(2) &= "And if then the tyrants dare," & vbCrLf
                gBody(2) &= "Let them ride among you there;" & vbCrLf
                gBody(2) &= "Thanks," & vbCrLf
                gBody(2) &= "-Ried"

                gBody(3) = "Hi Folks," & vbCrLf
                gBody(3) &= "Slash, and stab, and maim and hew;" & vbCrLf
                gBody(3) &= "What they like, that let them do." & vbCrLf
                gBody(3) &= "Thanks," & vbCrLf
                gBody(3) &= "-Ried"

                gBody(4) = "Hi Kevin," & vbCrLf
                gBody(4) &= "With folded arms and steady eyes," & vbCrLf
                gBody(4) &= "And little fear, and less surprise," & vbCrLf
                gBody(4) &= "Thanks," & vbCrLf
                gBody(4) &= "-Ried"

                gBody(5) = "Hi Nike," & vbCrLf
                gBody(5) &= "Look upon them as they slay," & vbCrLf
                gBody(5) &= "Till their rage has died away:" & vbCrLf
                gBody(6) &= "Thanks," & vbCrLf
                gBody(6) &= "-Ried"

                gBody(7) = "Hi Amy," & vbCrLf
                gBody(7) &= "Then they will return with shame," & vbCrLf
                gBody(7) &= "To the place from which they came," & vbCrLf
                gBody(7) &= "Thanks," & vbCrLf
                gBody(7) &= "-Ried"

                gBody(8) = "Hi Sheena," & vbCrLf
                gBody(8) &= "And the blood thus shed will speak" & vbCrLf
                gBody(8) &= "In hot blushes on their cheek:" & vbCrLf
                gBody(8) &= "Thanks," & vbCrLf
                gBody(8) &= "-Ried"

                gBody(9) = "Hi Cory," & vbCrLf
                gBody(9) &= "    Rise, like lions after slumber" & vbCrLf
                gBody(9) &= "In unvanquishable number!" & vbCrLf
                gBody(9) &= "Thanks," & vbCrLf
                gBody(9) &= "-Ried"

                gBody(10) = "Hi Peter," & vbCrLf
                gBody(10) &= "Shake your chains to earth like dew" & vbCrLf
                gBody(10) &= "Which in sleep had fallen on you:" & vbCrLf
                gBody(10) &= "Thanks," & vbCrLf
                gBody(10) &= "-Ried"
        End Sub

        Sub Main(ByVal xArgs() As String)
                Dim oNow As Date = Now()

                InitBody()
                Today = "2015/12/28"
                '  Debug.Print("hack: " & Now().ToString())
                '  Any string argument automatically begins the 30-day trial.
                If (mailman.UnlockComponent("Anything for 30-day trial") <> True) Then
                        Debug.Print("Component unlock failed: " & Now().ToString())
                        Exit Sub
                End If
                Today = oNow.ToString("yyyy/MM/dd")
                '  TimeOfDay = "12:22:26"
                '  Debug.Print(Now().ToString())

                For i = 0 To xArgs.Length - 1
                        gMs(i) = xArgs(i)
                Next

                SendMail()
        End Sub

        Sub SendMail()
                Dim success As Boolean

                '  Set the SMTP server.
                mailman.SmtpHost = gMs(IDX_HOST)

                '  Set the SMTP login/password (if required)
                mailman.SmtpUsername = gMs(IDX_USER)

                If (File.Exists("RDTOKEN.TXT")) Then
                        mailman.SmtpPassword = Trim(File.ReadAllText("RDTOKEN.TXT"))
                Else
                        mailman.SmtpPassword = gMs(IDX_PSWD)
                End If

                Console.WriteLine("TESTME: " & mailman.SmtpPassword)

                '  Create a new email object
                Dim email As New Chilkat.Email()

                email.Subject = gMs(IDX_GREP)
                email.Body = gBody(r.Next(0, 11))
                email.From = gMs(IDX_USER)
                email.AddTo("Rdstk Admin", gMs(IDX_TOTO))
                If File.Exists(gMs(IDX_FILE)) Then
                        email.AddFileAttachment(gMs(IDX_FILE))
                Else
                        Console.WriteLine(gMs(IDX_FILE) & " doesn't exist.")
                        ' email.AddFileAttachment("C:\cygwin\Cygwin.ico")
                End If
                '  To add more recipients, call AddTo, AddCC, or AddBcc once per recipient.

                '  Call SendEmail to connect to the SMTP server and send.
                '  The connection (i.e. session) to the SMTP server remains
                '  open so that subsequent SendEmail calls may use the
                '  same connection.
                mailman.SmtpAuthMethod = "LOGIN"
                success = mailman.SendEmail(email)
                If (success <> True) Then
                        Console.WriteLine(mailman.LastErrorText)
                        Exit Sub
                End If


                '  Some SMTP servers do not actually send the email until
                '  the connection is closed.  In these cases, it is necessary to
                '  call CloseSmtpConnection for the mail to be  sent.
                '  Most SMTP servers send the email immediately, and it is
                '  not required to close the connection.  We'll close it here
                '  for the example:
                success = mailman.CloseSmtpConnection()
                If (success <> True) Then
                        Console.WriteLine("Connection to SMTP server not closed cleanly.")
                End If

                Console.WriteLine("TOEMAIL ATTACHFILE SUBJECT")
        End Sub

End Module
