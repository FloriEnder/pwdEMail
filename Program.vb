Imports System.DirectoryServices
Imports System.Xml
Imports System.Net.Mail
Public Class Var
    Public Shared adpath As String
    Public Shared ldapfilter As String
    Public Shared mailserver As String
    Public Shared mailserverport As Integer
    Public Shared mailuser As New System.Net.NetworkCredential
    Public Shared mailsender As String
    Public Shared mailRplyTo As String
    Public Shared mailSubject As String
    Public Shared PathToLogo As String
    Public Shared sendadminmail As Boolean
    Public Shared adminmail As String
    Public Shared mailSubjectAdmin As String
    Public Shared ErrorInfos As String
    Public Shared mailssendtouser As String
    Public Shared debug As Boolean
End Class
Module Program
    Sub Main()
        GetConfig()
        GetAdUser()
        If Var.sendadminmail Then
            Sendmail(Var.adminmail, Today, adminmail:=Var.sendadminmail)
        End If
    End Sub
    Sub GetConfig()
        If (IO.File.Exists(".\Config.xml")) Then
            Dim XMLReader As New XmlTextReader(".\Config.xml")
            Try
                While XMLReader.Read
                    If XMLReader.Name = "ADPath" Then
                        Var.adpath = XMLReader.ReadElementString()
                    ElseIf XMLReader.Name = "ldapfilter" Then
                        Var.ldapfilter = XMLReader.ReadElementString()
                    ElseIf XMLReader.Name = "mailserver" Then
                        Var.mailserver = XMLReader.ReadElementString()
                    ElseIf XMLReader.Name = "mailserverport" Then
                        Var.mailserverport = XMLReader.ReadElementString()
                    ElseIf XMLReader.Name = "mailuser" Then
                        Var.mailuser.UserName = XMLReader.ReadElementString()
                    ElseIf XMLReader.Name = "mailuserpwd" Then
                        Var.mailuser.Password = XMLReader.ReadElementString()
                    ElseIf XMLReader.Name = "mailsender" Then
                        Var.mailsender = XMLReader.ReadElementString()
                    ElseIf XMLReader.Name = "mailRplyTo" Then
                        Var.mailRplyTo = XMLReader.ReadElementString()
                    ElseIf XMLReader.Name = "mailSubject" Then
                        Var.mailSubject = XMLReader.ReadElementString()
                    ElseIf XMLReader.Name = "PathToLogo" Then
                        Var.PathToLogo = XMLReader.ReadElementString()
                    ElseIf XMLReader.Name = "sendadminmail" Then
                        Var.sendadminmail = Convert.ToBoolean(XMLReader.ReadElementString())
                    ElseIf XMLReader.Name = "adminmail" Then
                        Var.adminmail = XMLReader.ReadElementString()
                    ElseIf XMLReader.Name = "mailSubjectAdmin" Then
                        Var.mailSubjectAdmin = XMLReader.ReadElementString()
                    ElseIf XMLReader.Name = "debug" Then
                        Var.debug = Convert.ToBoolean(XMLReader.ReadElementString())
                    End If
                End While
                XMLReader.Close()
            Catch ex As Exception
                Console.WriteLine(ex.ToString)
                Environment.Exit(1)
            End Try
        Else
            Console.WriteLine(".\Config.xml wurde nicht gefunden!")
            Environment.Exit(1)
        End If
    End Sub
    Sub Sendmail(receiver As String, ablaufdatum As Date, Optional ablaufintagen As Integer = 0, Optional name As String = "", Optional adminmail As Boolean = False)
        Dim Msg As New MailMessage
        Dim mySmtpsvr As New SmtpClient()
        mySmtpsvr.EnableSsl = True
        mySmtpsvr.Host = Var.mailserver
        mySmtpsvr.Port = Var.mailserverport
        mySmtpsvr.UseDefaultCredentials = False
        mySmtpsvr.Credentials = Var.mailuser
        Try
            Msg.From = New MailAddress(Var.mailsender)
            Dim txtmail As String = ""
            If adminmail Then
                Msg.IsBodyHtml = False
                Msg.To.Add(receiver)
                Msg.Subject = Var.mailSubjectAdmin
                txtmail = "E-Mail an " & Var.mailssendtouser.ToString & " gesendet." & vbCrLf & Var.ErrorInfos & vbCrLf & "Ende!"
            Else
                Msg.IsBodyHtml = True
                If Var.debug Then
                    Msg.To.Add(Var.adminmail)
                Else
                    Msg.To.Add(receiver)
                End If
                Dim txtmailALL As String() = IO.File.ReadAllLines(".\mail.html")
                Msg.Subject = Var.mailSubject
                Msg.ReplyTo = New MailAddress(Var.mailRplyTo)
                For Each str As String In txtmailALL
                    Dim replstr As String = str
                    replstr = Replace(replstr, "@Benutzer", name) & vbCrLf
                    If ablaufintagen < 0 Then
                        replstr = Replace(replstr, "@Passwortablauf", "lief vor " & (ablaufintagen * -1).ToString) & vbCrLf
                    Else
                        replstr = Replace(replstr, "@Passwortablauf", "läuft in " & ablaufintagen.ToString) & vbCrLf
                    End If
                    replstr = Replace(replstr, "@Ablaufdatum", ablaufdatum.ToString) & vbCrLf
                    Dim tmplogopath() As String = Split(Var.PathToLogo, "\")
                    replstr = Replace(replstr, "@Logo", tmplogopath(tmplogopath.Length() - 1)) & vbCrLf
                    txtmail += replstr & vbCrLf
                Next
                Msg.Attachments.Add(New Attachment(Var.PathToLogo))
            End If
            Msg.Body = txtmail
            If Var.debug Then
                Console.WriteLine("Mail send to " & name & " - " & receiver)
            End If
            mySmtpsvr.Send(Msg)
            Var.mailssendtouser += receiver + vbCrLf
        Catch ex As Exception
            If Var.debug Then
                Console.WriteLine("Fehler Mail versand " & vbCrLf & ex.ToString)
                Var.ErrorInfos += "Fehler Mail versand " & vbCrLf & ex.ToString & vbCrLf
            Else
                Var.ErrorInfos += "Fehler Mail versand " & vbCrLf & ex.ToString & vbCrLf
            End If
        End Try
    End Sub
    <CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Plattformkompatibilität überprüfen", Justification:="<Ausstehend>")>
    Sub GetAdUser()
        Dim ADPathShort As String = ""
        Dim tmparray() As String = Split(Var.adpath, ",")
        For Each tmpstring In tmparray
            If Mid(tmpstring, 1, 1) = "D" Then
                ADPathShort += tmpstring & ","
            End If
        Next
        ADPathShort = ADPathShort.Substring(0, ADPathShort.Length - 1)
        Dim de As New DirectoryEntry("LDAP://" & ADPathShort)
        Dim ds As New DirectorySearcher(de) With {.Filter = ("(maxPwdAge=*)")}
        Try
            Dim srcResult As SearchResult = ds.FindOne
            Dim maxpwdage As Int64 = Convert.ToInt64(srcResult.Properties("maxPwdAge").Item(0))
            maxpwdage = maxpwdage / 60 / 60 / 24 / 10000000 * -1
            If Var.debug Then
                Console.WriteLine("maxpwdage -> " & maxpwdage)
            End If
        Catch ex As Exception
            Console.WriteLine("Fehler GetMaxPWDAge" & vbCrLf & ex.ToString)
            Environment.Exit(1)
        End Try
        Dim de1 As New DirectoryEntry("LDAP://" & Var.adpath)
        Dim ds1 As New DirectorySearcher(de1) With {.Filter = Var.ldapfilter}
        ds1.PropertiesToLoad.Add("cn")
        ds1.PropertiesToLoad.Add("pwdLastSet")
        ds1.PropertiesToLoad.Add("mail")
        ds1.PropertiesToLoad.Add("userAccountControl")
        Dim src As SearchResultCollection = ds1.FindAll()
        For Each sr As SearchResult In src
            Try
                If Convert.ToInt32(sr.Properties("userAccountControl").Item(0)) = 512 And sr.Properties("userAccountControl").Item(0).ToString <> "" Then
                    Dim pwdLastSet As Date
                    If sr.Properties("pwdLastSet").Count <> 0 Then
                        pwdLastSet = Date.FromFileTimeUtc(sr.Properties("pwdLastSet").Item(0))
                    End If
                    Select Case DateDiff(DateInterval.Day, Date.Today, pwdLastSet.AddDays(maxpwdage))
                        Case 14
                            Sendmail(sr.Properties("mail").Item(0).ToString, pwdLastSet.AddDays(maxpwdage), 14, sr.Properties("cn").Item(0).ToString)
                        Case 7
                            Sendmail(sr.Properties("mail").Item(0).ToString, pwdLastSet.AddDays(maxpwdage), 7, sr.Properties("cn").Item(0).ToString)
                        Case 5
                            Sendmail(sr.Properties("mail").Item(0).ToString, pwdLastSet.AddDays(maxpwdage), 5, sr.Properties("cn").Item(0).ToString)
                        Case 4
                            Sendmail(sr.Properties("mail").Item(0).ToString, pwdLastSet.AddDays(maxpwdage), 4, sr.Properties("cn").Item(0).ToString)
                        Case 3
                            Sendmail(sr.Properties("mail").Item(0).ToString, pwdLastSet.AddDays(maxpwdage), 3, sr.Properties("cn").Item(0).ToString)
                        Case 2
                            Sendmail(sr.Properties("mail").Item(0).ToString, pwdLastSet.AddDays(maxpwdage), 2, sr.Properties("cn").Item(0).ToString)
                        Case 1
                            Sendmail(sr.Properties("mail").Item(0).ToString, pwdLastSet.AddDays(maxpwdage), 1, sr.Properties("cn").Item(0).ToString)
                        Case 0
                            Sendmail(sr.Properties("mail").Item(0).ToString, pwdLastSet.AddDays(maxpwdage), 0, sr.Properties("cn").Item(0).ToString)
                        Case < 0
                            Sendmail(sr.Properties("mail").Item(0).ToString, pwdLastSet.AddDays(maxpwdage), Convert.ToInt32(DateDiff(DateInterval.Day, Date.Today, pwdLastSet.AddDays(maxpwdage))), sr.Properties("cn").Item(0).ToString)
                        Case Else
                            If Var.debug Then
                                Console.WriteLine("in " & DateDiff(DateInterval.Day, Date.Today, pwdLastSet.AddDays(maxpwdage)).ToString & "-> " & sr.Properties("cn").Item(0).ToString)
                            End If
                    End Select
                End If
            Catch ex As Exception
                If Var.debug Then
                    Console.WriteLine("Fehler GetADuser" & vbCrLf & ex.ToString)
                    Var.ErrorInfos += "Fehler GetADuser" & vbCrLf & ex.ToString & vbCrLf
                Else
                    Var.ErrorInfos += "Fehler GetADuser" & vbCrLf & ex.ToString & vbCrLf
                End If
            End Try
        Next
    End Sub
End Module
