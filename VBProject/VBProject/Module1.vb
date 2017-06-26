Imports OpenPop.Pop3
Imports OpenPop.Mime
Imports System.Data
Imports System.Console
Imports System.IO
Imports System.Security.Permissions
Imports System.Security
Imports System.Security.AccessControl
Imports System.Text

Module Module1

    'Dim file paths
    Dim pathF As String = "C:\000-POP3-FILES"
    Dim pathA As String = "C:\000-POP3-FILES\WWW_TEST_IN_PROGRESS.txt"
    Dim pathB As String = "C:\000-POP3-FILES\WWW_ERROR_MESSAGE.txt"
    Dim pathC As String = "C:\000-POP3-FILES\WWW_BODY.txt"
    Dim pathD As String = "C:\000-POP3-FILES\WWW_BODY_COMPLETE.txt"
    Dim pathE As String = "C:\POP3ini.txt"
    'These will be read from the POP3ini file
    Dim hostname As String
    Dim port As Integer
    Dim ssl As Boolean
    Dim tsl As Boolean
    Dim username As String
    Dim password As String
    'Dim encod_ As String
    Dim encod As Encoding = Encoding.Default
    Dim plaintext As String

    'If main is called with parameter "getMessage", executes Write_Files and proceeds accordingly
    'If main is called with parameter "deleteMessage", executes Delete_Mails and proceeds accordingly
    Sub Main(ByVal args As String())

        InitiateFolder()
        InitiateFiles()
        'WriteLine("Initiating Session...{0}{0}", vbCrLf)
        Try
            ReadFileE() 'get info for the connection
            If My.Application.CommandLineArgs(0).Equals("deleteMessage") Then
                Delete_Mails()
            ElseIf My.Application.CommandLineArgs(0).Equals("getMessage") Then
                Write_Files()
            End If
        Catch Exp As Exception
            'If any errors occur,creates the error log file
            CreateF_B(Exp.Message)
            CreateF_D(plaintext)
        End Try
        'WriteLine()
        'WriteLine()
        'WriteLine("Press any key to continue")
        'ReadKey()

    End Sub

    'Just in case folder does not exist, create it
    'In that case it will do nothing of course, windows security configuration and permissions will block this action
    Sub InitiateFolder()

        If (Not Directory.Exists(pathF)) Then
            Directory.CreateDirectory(pathF)
        End If

    End Sub

    'Just in case the required files already exist, delete them
    'Does not delete other files in the folder such as pre-existing attachements
    Sub InitiateFiles()

        If File.Exists(pathA) Then
            File.Delete(pathA)
        End If

        If File.Exists(pathB) Then
            File.Delete(pathB)
        End If

        If File.Exists(pathC) Then
            File.Delete(pathC)
        End If

        If File.Exists(pathD) Then
            File.Delete(pathD)
        End If

    End Sub

    'Creates a connection with the info from POP3ini file
    'Gets the requested email's fields
    'Reads and -if necessary- trims the mail, then writes/downloads all the requested files
    Sub Write_Files()

        Try
            'Creates a connection
            Dim pop3Client As Pop3Client
            pop3Client = New Pop3Client
            pop3Client.Connect(hostname, port, ssl)
            pop3Client.Authenticate(username, password, AuthenticationMethod.UsernameAndPassword)

            Dim count As Integer = pop3Client.GetMessageCount
            'number in the next line indicates the message to be read - for debbugging purposes, must be used without the following comparison
            'Dim message As Message = pop3Client.GetMessage(10)
            Dim message1 As Message = pop3Client.GetMessage(1)
            Dim message2 As Message = pop3Client.GetMessage(count)

            'Find the oldest message by comparing the dates of the first and the last message in the inbox
            Dim message As Message
            Dim num As String
            If message1.Headers.Date > message2.Headers.Date Then
                message = message2
                num = count
            Else
                message = message1
                num = 1
            End If

            'Start reading the mail's info
            'Gets sender
            Dim from As String = message.Headers.From.Address
            'Gets sender displayed name
            Dim fromdis As String = message.Headers.From.DisplayName
            'Gets sent addresses
            Dim sentto As String = ""
            For Each sentto_ In message.Headers.Cc
                Dim s1 As String = sentto_.Address
                sentto = String.Concat(sentto, String.Concat(s1, "; "))
            Next
            'Gets cc'ed addresses
            Dim cc As String = ""
            For Each cc_ In message.Headers.Cc
                Dim c1 As String = cc_.Address
                cc = String.Concat(cc, String.Concat(c1, "; "))
            Next
            'Gets subject
            Dim subject As String = message.Headers.Subject
            'Gets attachement count
            Dim attach As String = message.FindAllAttachments.ToArray.Length
            'Gets date
            Dim recdate As String = message.Headers.Date

            'Attachement Handling
            'If there are attachements "downloades" them to the folder
            If attach > 0 Then
                'WriteLine()
                'WriteLine("Saving attachements...")
                'WriteLine()
                For Each msgpart As MessagePart In message.FindAllAttachments
                    Dim thefile = msgpart.FileName
                    Dim filetype = msgpart.ContentType
                    Dim contentid = msgpart.ContentId
                    File.WriteAllBytes(pathF & "\" & thefile, msgpart.Body)
                Next
            End If

            'Creating Files A,C..
            If (count > 0) Then
                CreateF_A(count, num, from, fromdis, sentto, cc, subject, attach, recdate) '..passing all the info
                If (message.FindFirstPlainTextVersion IsNot Nothing) Then '..if message is not empty
                    If Not (message.FindFirstPlainTextVersion.IsAttachment) Then '..if message body is not an attachement
                        Dim plaintext As String = message.FindFirstPlainTextVersion.GetBodyAsText '..gets the text from the message
                        Dim bcount As Integer = encod.GetByteCount(plaintext)
                        'If text file is over 800bytes, cut it
                        If bcount > 800 Then
                            Dim txtcut As Byte() = encod.GetBytes(plaintext, 0, 800)
                            CreateF_C(encod.GetString(txtcut))
                        Else
                            CreateF_C(plaintext)
                        End If
                    End If
                End If
            End If

        Catch Exp As Exception
            'If any errors occur,creates the error log file
            CreateF_B(Exp.Message)
            CreateF_D(plaintext)
        End Try

    End Sub

    'Creates the file WWW_TEST_IN_PROGRESS with all the necessary info that have already been processed in Write_Files
    Sub CreateF_A(ByVal count As Integer, ByVal num As String, ByVal from As String, ByVal fromdis As String, ByVal sentto As String, ByVal cc As String, ByVal subject As String, ByVal attach As String, ByVal recdate As String)

        'WriteLine("Creating... WWW_TEST_IN_PROGRESS.txt")
        File.Create(pathA).Dispose()

        'This string will contain the necessary info for file A in the correct format
        Dim info As String = String.Format("totMessagesInInbox  = {1}{0}curMessageNum       = {2}{0}senderMailX         = {3}{0}senderMailDisplayX  = {4}{0}senderMailTO        = {5}{0}senderMailCC        = {6}{0}subjectX            = {7}{0}totFilesAttachedX   = {8}{0}dateReceivedX       = {9}{0}", vbCrLf, count, num, from, fromdis, sentto, cc, subject, attach, recdate)

        Dim fs As FileStream = File.Create(pathA, 1024)
        Dim sw As StreamWriter = New StreamWriter(fs, encod)
        sw.Write(info, encod)
        sw.Write(vbCrLf)
        sw.Close()

    End Sub

    'Creates the file with the error message log if any errors occur
    Sub CreateF_B(ByVal Err As String)

        'WriteLine("Creating... WWW_ERROR_MESSAGE.txt")
        File.Create(pathB).Dispose()

        Dim fs As FileStream = File.Create(pathB, 1024)
        Dim sw As StreamWriter = New StreamWriter(fs, Text.Encoding.Default)
        sw.Write(Err)
        sw.Write(vbCrLf)
        sw.Close()

    End Sub

    'Creates the file with the message contents that do not exceed 800bytes - this is already checked in Write_Files
    Sub CreateF_C(ByVal plaintext As String)

        'WriteLine("Creating... WWW_BODY.txt")
        File.Create(pathC).Dispose()

        Dim fs As FileStream = File.Create(pathC, 800)
        Dim sw As StreamWriter = New StreamWriter(fs, encod)
        sw.Write(plaintext)
        sw.Write(vbCrLf)
        sw.Close()

    End Sub

    'Creates the file with the whole message text in case of errors - if the message was read
    Sub CreateF_D(ByVal plaintext As String)

        'WriteLine("Creating... WWW_BODY_COMPLETE.txt")
        File.Create(pathD).Dispose()

        Dim fs As FileStream = File.Create(pathD, 1024)
        Dim sw As StreamWriter = New StreamWriter(fs, encod)
        sw.Write(plaintext)
        sw.Write(vbCrLf)
        sw.Close()

    End Sub

    'Reads the info from the file pop3ini so it can initiate a session
    'Checks if file exists - if it doesn't the program will just create the error log file
    Sub ReadFileE()
        Dim line As String
        Dim port_ As String
        Dim ssl_ As String
        Dim tsl_ As String
        Dim i As Integer = 0
        Dim objReader As New StreamReader(pathE)
        'WriteLine("Reading... POP3ini.txt")
        Do While objReader.Peek() <> -1
            line = objReader.ReadLine()
            i = 1 + i
            'getting the info while counting lines and using substring while looping the input
            If (i.Equals(1)) Then
                hostname = line.Substring(15, line.Length - 15)
                'WriteLine("Hostname read :" + hostname)
            ElseIf (i.Equals(2)) Then
                port_ = line.Substring(15, line.Length - 15)
                Integer.TryParse(port_, port)
                'WriteLine("Port read :" + port_)
            ElseIf (i.Equals(3)) Then
                ssl_ = line.Substring(15, line.Length - 15)
                If ssl_.Equals("Yes") Then
                    ssl = True
                Else
                    ssl = False
                End If
                'WriteLine("Ssl read :" + ssl_)
            ElseIf (i.Equals(4)) Then
                tsl_ = line.Substring(15, line.Length - 15)
                If tsl_.Equals("Yes") Then
                    tsl = True
                Else
                    tsl = False
                End If
                'WriteLine("Tsl read :" + tsl_)
            ElseIf (i.Equals(5)) Then
                username = line.Substring(15, line.Length - 15)
                'WriteLine("User read :" + username)
            ElseIf (i.Equals(6)) Then
                password = line.Substring(15, line.Length - 15)
                'WriteLine("Password read :" + password)
            End If
        Loop

    End Sub

    'Creates a connection, finds the oldest mail and deletes it from inbox
    'Disconnects to commit the change
    Sub Delete_Mails()

        'WriteLine("Deleting oldest message...")
        Try
            'Create a connection
            Dim pop3Client As Pop3Client
            pop3Client = New Pop3Client
            pop3Client.Connect(hostname, port, ssl)
            pop3Client.Authenticate(username, password, AuthenticationMethod.UsernameAndPassword)

            Dim count As Integer = pop3Client.GetMessageCount
            Dim message1 As Message = pop3Client.GetMessage(1)
            Dim message2 As Message = pop3Client.GetMessage(count)

            'Find the oldest message by comparing the dates of the first and the last message in the inbox
            If message1.Headers.Date > message2.Headers.Date Then
                pop3Client.DeleteMessage(count)
            Else
                pop3Client.DeleteMessage(1)
            End If
            'It is necessary to disconnect to commit the change
            pop3Client.Disconnect()
        Catch Exp As Exception
            'If any errors occur,creates the error log file
            CreateF_B(Exp.ToString())
        End Try

    End Sub

End Module
