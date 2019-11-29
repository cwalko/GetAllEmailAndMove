Imports System
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

Module Module1
    Dim sqlcon As SqlConnection
    Dim Location As String = Where()

    Sub Main
        OpenSqlConnection(Location) ' find location of computer via its OS

        ConvertToOrder()

    End Sub
    Sub ConvertToOrder()
        'Create Directories - sourceDirectory and  

        Dim sourcePath As String = GetFolderDate()
        sourcePath = "C:\eCommerseIntegration\email_Data\2019-10-16\" '****FOR TESTING ONLY*****
        Dim archivePath As String = sourcePath + "archive\"
        Dim archiveDirectory As DirectoryInfo = Directory.CreateDirectory(sourcePath + "archive\")

        Try
            Dim txtFiles = Directory.EnumerateFiles(sourcePath, "*.txt")
            Dim FileCount As Integer = txtFiles.Count
            Msg("Count = ", FileCount)
            For Each currentFile As String In txtFiles
                Dim fileName = currentFile.Substring(sourcePath.Length)
                Dim fileContent As String() = IO.File.ReadAllLines(sourcePath + fileName)
                Msg("line 2 = ", fileContent(1))
                'Directory.Move(currentFile, Path.Combine(archivePath, fileName))
            Next
        Catch e As System.Exception
            Msg("Error Msg = ", e.Message)
        End Try

    End Sub

    Function NextFriday(ByRef Tdate) As String
        Tdate = Tdate.AddDays(1)
        Do While Tdate.DayOfWeek <> DayOfWeek.Friday
            Tdate = Tdate.AddDays(1)
        Loop
        Return Tdate
    End Function
    Function GetFolderDate() As String
        Dim fldDate As Date = Now()
        If fldDate.DayOfWeek <> DayOfWeek.Friday Then
            NextFriday(fldDate)
        End If
        Dim myYear As Integer = Year(fldDate)  ' Get the year
        Dim myMonth As Integer = Month(fldDate)  ' Get the month
        Dim myDate As String = Format(fldDate, "dd")   ' "ddMMMyyyy"  Get the Date
        Dim folderName As String = myYear.ToString() & "-" & myMonth.ToString & "-" & myDate & "\"  ' build the Destination "file name" (date)
        Dim email_Data_Folder As String = ConfigurationManager.AppSettings.Get("email_Data_Dir") + folderName
        Dim sourcePath As String = email_Data_Folder
        Return sourcePath
    End Function

    Sub OpenSqlConnection(Location) ' use the Location to read the correct Connection String from commonSettings.config
        Dim MyConnectionString As String = ConfigurationManager.AppSettings.Get("Seafood" + Location)
        ' MyConnectionString = ConfigurationManager.ConnectionStrings("conString").ConnectionString   ' from App.config
        sqlcon = New SqlConnection(MyConnectionString)
        Try
            sqlcon.Open()
        Catch ex As System.Exception
            Msg("error opening Connection, error = " + vbCrLf, ex.ToString)
        End Try
    End Sub

    Sub Msg(text As String, var As String)
        Console.Write(text + " " + var)
        Console.ReadLine()
    End Sub
    Function Where() ' Determine where the code is being run so the correct Connection String can be loaded
        Dim Location As String = ""
        Select Case My.Computer.Info.OSFullName
            Case "Microsoft Windows Server 2008 R2 Standard" ' code is running a Church
                Location = "Church"
            Case "Microsoft Windows 10 Home" ' Code is running at Home
                Location = "Home"
            Case Else ' code is not at home or church, we will not have Connection String
                Dim errTxt As String = "CreateOrdersFromText program SHUTTING DOWN.  Can not find correct Location/Computer code - Location unkown!!" + vbCrLf + "Check, did someone change the OS of the computer?  This Program is shutting down!"
                Dim OrderId As String = "N/A"
                Dim eMailContentFileName As String = ConfigurationManager.AppSettings.Get("SetupDir") + "Error - ComputerLocationUnknown.txt"
                WriteErrorFile(errTxt, OrderId, eMailContentFileName)
                Msg(errTxt, "")
                End
        End Select
        Return Location
    End Function
    Sub WriteErrorFile(errTxt As String, OrderId As String, eMailContentFileName As String)
        Try
            Dim errFilePath As String = ConfigurationManager.AppSettings.Get("ErrorFile")
            IO.File.AppendAllText(errFilePath, vbCrLf + Now.ToUniversalTime + vbCrLf + errTxt + vbCrLf)
            eMail_Setup(eMailContentFileName, OrderId)  'Prepare email that this program is shutting down, most likely because Outlook is NOT running 
        Catch ex As IOException
            eMailContentFileName = "unknown error.txt" + ex.ToString
            Module1.eMail_Setup(eMailContentFileName, OrderId)  'Prepare email that this program is shutting down, most likely because Outlook is NOT running 
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Sub eMail_Setup(eMailContentFileName As String, OrderId As String)

        ' Dim EmailDirectory() As String = IO.File.ReadAllLines("C:\eCommerseIntegration\Setup Files\Ecommerse Setup Directory.txt")
        'Dim Directory As String = EmailDirectory(0)
        Dim emailContent As String() = IO.File.ReadAllLines(eMailContentFileName)
        Dim sSubject As String = emailContent(0) + "   ----  OrderId = " + OrderId
        Dim sBody As String = emailContent(2) + "   ----  OrderId = " + OrderId
        Dim sTo As String = emailContent(3) ' add Rick Jeakle and Greg Dohner
        Dim sCC As String = emailContent(4) ' seafood@jaeklegroup.net
        Dim sFilename As String = emailContent(5)
        Dim sDisplayname As String = emailContent(6)
        sEmailSend(sSubject, sBody, sTo, sCC, sFilename, sDisplayname)

    End Sub
    Sub sEmailSend(sSubject As String, sBody As String,
                             sTo As String, sCC As String,
                             sFilename As String, sDisplayname As String)
        Dim oApp As Microsoft.Office.Interop.Outlook._Application
        oApp = New Microsoft.Office.Interop.Outlook.Application

        Dim oMsg As Microsoft.Office.Interop.Outlook._MailItem
        oMsg = oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
        oMsg.Subject = sSubject
        oMsg.Body = sBody
        oMsg.To = sTo
        oMsg.CC = sCC

        Dim strS As String = sFilename
        Dim strN As String = sDisplayname
        If sFilename <> "" Then
            Dim sBodyLen As Integer = Int(sBody.Length)
            Dim oAttachs As Microsoft.Office.Interop.Outlook.Attachments = oMsg.Attachments
            Dim oAttach As Microsoft.Office.Interop.Outlook.Attachment
            oAttach = oAttachs.Add(strS, , sBodyLen, strN)
        End If
        oMsg.Send()
        ' MsgBox("Email Sent")

    End Sub
End Module
