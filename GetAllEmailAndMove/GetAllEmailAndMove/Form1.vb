
'get all the mail from the Inbox and move it to the Destination folder
'
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
'Imports System.Drawing.Printing ' used to print List Box
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Text

Public Class Form1

    ' ReadOnly dt As New DataTable
    Dim outLookApp As Application = New Application()
    Dim outLookNS As [NameSpace] = outLookApp.GetNamespace("MAPI")
    Dim olSourceFolder As MAPIFolder = Nothing
    Dim mi As MailItem
    Dim olDestinationFolder As MAPIFolder = Nothing
    Dim olBadCreditFolder As MAPIFolder = Nothing
    '   Dim dr As DataRow = Nothing
    Dim isread As String = "Unread"
    Dim i As Integer
    Dim subject As String
    Public PayTransID As String
    Dim myYear As Integer
    Dim myMonth As Integer
    Dim myDate As String
    ' Dim File As String = ""
    Dim SQLDB As String = "eCom_Email"
    Dim CreationTime, SentOn, ReceivedTime As Date
    Dim SenderName, SenderEmailAddress As String
    Dim emailBody As String
    Public fileDate As String
    Dim IP As String
    Dim CustomerName As String

    Dim rerun As Integer = 0 ' 0 mean first time running; 1 mean we are running again
    ReadOnly separators As Char() = New Char() {Chr(13), Chr(10)} 'This is a CR/LineFeed to find the "each line" of the email body

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ScrollBox("Text")
        olSourceFolder = outLookNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) ' where new eCommerse emails are received - the inbox of "seafood@JaekleGroup.com"
        Dim outLookApp As Application = New Application()
        Label1.Text = "Unknown"
        Me.Show()
    End Sub
    Public Sub _tMonitorOutLook()
        ' AddHandler outLookApp.NewMailEx, AddressOf outLookApp_NewMailEx
        Dim fldDate As Date = Now()

        If fldDate.DayOfWeek <> DayOfWeek.Friday Then
            NextFriday(fldDate)
        End If
        myYear = Year(fldDate)  ' Get the year
        myMonth = Month(fldDate)  ' Get the month
        myDate = Format(fldDate, "dd")   ' "ddMMMyyyy"  Get the Date
        fileDate = myYear.ToString() & "-" & myMonth.ToString & "-" & myDate  ' build the Destination "file name" (date)
        '   MsgBox("fileDate = " + fileDate)
        If olSourceFolder.Items.Count <> 0 Then 'There is email in Inbox
            Cursor = Cursors.Default
            ProcessEmail() ' Process the  new email
        End If
        Try

            While olSourceFolder.Items.Count = 0
                Delay(10)
                'TextBox1.Text = TextBox1.Text + ("Listening for New emails!") + vbCrLf
                ScrollBox("Text")
                'Thread.Sleep(2000) ' Sleep for 2 seconds
            End While
        Catch
            Dim eMailContentFileName As String = "Outlook Down.txt"
            Module1.eMail_Setup(eMailContentFileName, "")  'Prepare email that this program is shutting down, most likely because Outlook is NOT running 
        End Try
        '  MultiBeep(356)
        Me.Show()
        Me.Refresh()
        'TextBox1.ForeColor = Color.Red
        TextBox1.Text = TextBox1.Text + ("Inbox is empty - Listening for New emails!") + vbCrLf
        ScrollBox("Text")
        ProcessEmail()
    End Sub
    Sub Delay(ByVal dblSecs As Double)
        Cursor = Cursors.WaitCursor 'and some various me.Cursor / current.cursor 
        Const OneSec As Double = 1.0# / (1440.0# * 60.0#)
        Dim dblWaitTil As Date
        Now.AddSeconds(OneSec)
        dblWaitTil = Now.AddSeconds(OneSec).AddSeconds(dblSecs)
        Do Until Now > dblWaitTil
            System.Windows.Forms.Application.DoEvents() ' Allow windows messages to be processed
        Loop

    End Sub

    Sub MultiBeep(numbeeps)
        For counter = 1 To numbeeps
            Beep()
        Next counter
    End Sub
    Sub CreateFolder()
        Try
            olDestinationFolder = outLookNS.Folders("seafood@JaekleGroup.com").Folders("done").Folders(fileDate)
        Catch
            olDestinationFolder = outLookNS.Folders("seafood@JaekleGroup.com").Folders("done").Folders.Add(fileDate, Outlook.OlDefaultFolders.olFolderInbox)
        End Try
        Try
            olBadCreditFolder = outLookNS.Folders("seafood@JaekleGroup.com").Folders("done").Folders(fileDate).Folders("BadCredit") 'Existing folder
        Catch
            olBadCreditFolder = olDestinationFolder.Folders.Add("BadCredit")  'outLookNS.Folders(olDestinationFolder.Folders.Add("BadCredit", Outlook.OlDefaultFolders.olFolderInbox))
        End Try

    End Sub

    Sub ProcessEmail()  'BtnProcessNew_Click(sender As Object, e As EventArgs) Handles BtnProcessNew.Click
        Dim EnteredBy As String = "eCommerce Middleware Program"
        Dim OrderType As String = "e-Commerce"
        Dim PaymentType As String = "Credit Card"
        Dim PaymentStatus As String = ""
        Dim myYear As Integer = Now.Year
        Dim IP As String = ""
        Dim CustomerName As String = ""
        Dim StrAddress As String = ""
        Dim ZipCode As String = ""
        Dim CSZ As String = ""
        Dim Phone As String = ""
        Dim eMail As String = ""
        Dim DeliveryType As String = ""
        Dim requestedHour As String = ""
        Dim TimeOrderTaken As String = ""
        Dim nFriday As String = ""
        Dim PaymentMethod As String = ""
        Dim Approved As String = ""
        Dim Auth_Info As String = ""
        Dim Comments_1 As String = "" ' Comments may have several lines; assuming max of 4
        Dim Comments_2 As String = ""
        Dim Comments_3 As String = ""
        Dim Comments_4 As String = ""
        Dim Comment As String = "" ' All 4 comment lines merged togther later in code
        Dim TotalPrice As String = ""
        Dim orderStartsAt As Integer = 20
        Dim ECommerce_DataPath As String = ""
        Dim ECommerce_DataFileId As String = ""
        Dim eComOrder_data As String = ""
        Dim Order_Header As String = "Qty" + vbTab + "Sku" + vbTab + "Description" + vbTab + "Price" + vbTab + "Total"
        ' Dim StopB4Move As String = TextBox2.Text
        Dim eMailContentFileName As String

        If rerun = 0 Then
            CreateFolder()
            ' CreateDataTable()
            rerun = 1
        End If

        ' For Each mail item As Object In olSourceFolder.Items
        If olSourceFolder.Items.Count = 0 Then
            Cleanup()
        End If

        For i = olSourceFolder.Items.Count To 1 Step -1
            Try
                mi = TryCast(mi, Outlook.MailItem)
                mi = olSourceFolder.Items(i)

                'Parse out the eCommerse Cart Order Number
                subject = mi.Subject
                Dim substring As String

                Dim dIndex = subject.IndexOf("#")
                If (dIndex > -1) Then
                End If
                substring = subject.Substring(Str(dIndex) + 1)
                PayTransID = substring
                CreationTime = mi.CreationTime
                Dim SentOn As Date = mi.SentOn
                ReceivedTime = mi.ReceivedTime
                SenderName = mi.SenderName
                SenderEmailAddress = mi.SenderEmailAddress
                emailBody = mi.Body
                ' MsgBox("sqldb = " + SQLDB)
                Call Module1.Insert_to_SQL(SQLDB, PayTransID, CreationTime, SentOn, ReceivedTime, SenderName, SenderEmailAddress, emailBody)
                'MsgBox(SQLDB & " - " & PayTransID  & " - " & CreationTime & " - " & SentOn & " - " & ReceivedTime & " - " & SenderName & " - " & SenderEmailAddress)


                'Parse each Email Line
                Dim emailBodyLines As String() = emailBody.Split(separators, StringSplitOptions.RemoveEmptyEntries) ' an array to hold each 'line' of the email
                If emailBodyLines(0) <> "Check if Credit Card Approved!!" Then ' if not a seafood order move it out of inbox and send an email
                    eMailContentFileName = "NON Ecommerse Email.txt"
                    Module1.eMail_Setup(eMailContentFileName, "")  'Prepare to send email that an email has come in that is NOT an Ecommerse Order

                    mi.Move(olBadCreditFolder)
                    _tMonitorOutLook() ' go back to watching for new email
                End If

                Dim emailLastLineIndex As Integer = emailBodyLines.Count - 1 ' Count of lines in email - use -1 to get last line
                'MsgBox("Last line number = " + emailLastLineIndex.ToString)
                'MsgBox("emailBody Last Line = " + emailBodyLines(emailLastLineIndex))

                Dim line As Integer = -1
                'Dim Comments_1 As String
                For Each s As String In emailBodyLines 'loop through each line of the email
                    line += 1
                    ListBox1.Items.Add(line & ") " & (s))
                    ScrollBox("List")
                    ' ListBox1.SelectedItem = ListBox1.Items.Count - 1 ' scrolls the list box  so you can see line being processed

                    Select Case line
                        Case 3 ' IP Adress Of Customer
                            IP = s
                            'MsgBox(s)
                            If Not IP.IndexOf("IP: ") = -1 Then
                                Dim Start As Integer = IP.IndexOf("IP: ") + 4
                                ' MsgBox("start =" + Start.ToString)
                                Dim Take As Integer = IP.Length - Start
                                IP = IP.Substring(Start, Take)
                            Else
                                IP = ""
                            End If
                      '  MsgBox("New IP:" & IP)
                        Case 5  ' Name of Customer
                            CustomerName = s
                            Dim Take As Integer = CustomerName.Length - 2
                            CustomerName = CustomerName.Substring(2, Take)
                    '   MsgBox("Customer name: " & CustomerName)
                        Case 6 ' Street Address of Customer
                            StrAddress = s
                     '  MsgBox("Street:" & StrAddress)
                        Case 7 ' City State Zip
                            CSZ = s
                            ZipCode = s.Substring(s.Length - 5, 5)
                       '  MsgBox("zip = " + ZipCode)
                       '  MsgBox("CSZ" & CSZ)
                        Case 8 ' Phone Number of Customer
                            'Dim Phone As String
                            Phone = Regex.Replace(s, "\D", "") 'replaces any non-numeric data with nothing
                  '     MsgBox("phone:" & Phone)
                        Case 9 ' email Address of customer
                            eMail = s
                            Dim Start As Integer
                            If Not eMail.IndexOf(": ") = -1 Then
                                Start = eMail.IndexOf(": ") + 2
                                Dim Take As Integer = eMail.Length - Start
                                eMail = eMail.Substring(Start, Take)
                            Else
                                eMail = ""
                            End If
                    '   MsgBox("email:" & eMail)
                        Case 10 ' Delivery Type - Eat in Take Out
                            DeliveryType = s
                            Dim Start As Integer
                            If Not DeliveryType.IndexOf(": ") = -1 Then
                                Start = DeliveryType.IndexOf(": ") + 2
                                Dim Take As Integer = DeliveryType.Length - Start
                                DeliveryType = DeliveryType.Substring(Start, Take)
                            Else
                                DeliveryType = ""
                            End If
                  '     MsgBox("DeliveryType:" & DeliveryType)
                        Case 11  ' Requested Time of Dinner Pickup
                            'CreationTime As Date = CreationTime
                            'sgBox("CreationTime: " & CreationTime)
                            requestedHour = s
                            Dim Start As Integer
                            Dim Tdate As Date = CreationTime
                            TimeOrderTaken = CreationTime.ToString("yyyy-MM-ddThh:mmZ")
                            ' MsgBox("CreationTime: " & TimeOrderTaken)
                            If Not requestedHour.IndexOf(": ") = -1 Then
                                Start = requestedHour.IndexOf(": ") + 2
                                Dim Take As Integer = requestedHour.Length - Start
                                requestedHour = requestedHour.Substring(Start, Take)
                                'Dim RequestedTime As String = (RequestedTime)
                                requestedHour = requestedHour.Insert(1, ":")
                            Else
                                requestedHour = ""
                            End If

                            'http://net-informations.com/q/faq/stringdate.html
                            'CONVERT STRING TO DATA?Time

                            If Tdate.DayOfWeek <> DayOfWeek.Friday Then
                                NextFriday(Tdate)
                            End If
                            nFriday = Tdate.ToString("yyyy-MM-ddThh:mmZ")
                            Dim endTake As Integer = nFriday.IndexOf("T")
                            nFriday = nFriday.Substring(0, [endTake])
                            nFriday = nFriday & "T" & requestedHour & "Z"
                          '  MsgBox("Requested Time: " & nFriday)
                        Case 14 ' Payment Method
                            PaymentMethod = s
                            'MsgBox("index= " & PaymentMethod.IndexOf(":"))
                            If Not PaymentMethod.IndexOf(":") = -1 Then
                                Dim Start As Integer = PaymentMethod.IndexOf(":") + 2
                                Dim Take As Integer = PaymentMethod.Length - Start
                                PaymentMethod = PaymentMethod.Substring(Start, Take)
                            Else
                                PaymentMethod = ""
                            End If
                       ' MsgBox("PaymentMethod: " & PaymentMethod)
                        Case 15 'Approved Yes/No
                            Approved = s
                            If Not Approved.IndexOf(":") = -1 Then
                                Dim Start As Integer = Approved.IndexOf(":") + 2
                                Dim Take As Integer = Approved.Length - Start
                                'MsgBox("Start/Take: " & Start & " / " & Take)
                                Approved = Approved.Substring(Start, Take)
                                Approved = Approved.Replace(vbTab, "")

                                If chkBox_Testing.Checked = True Then
                                    chkBox_Testing.Text = "Testing"
                                    Approved = "Yes" ' JUST FOR TESTING, if Testing Check Box is checked
                                Else
                                    chkBox_Testing.Text = "Not Testing"
                                End If

                                If Approved = "Yes" Then
                                    PaymentStatus = "Paid"
                                    '  MsgBox("Payment Status =" + PaymentStatus)
                                Else
                                    PaymentStatus = "Not Paid"

                                    eMailContentFileName = "CREDIT CARD WAS NOT ACCEPTED.txt"
                                    Module1.eMail_Setup(eMailContentFileName, PayTransID)  'Prepare email that an order has come in with a rejected CC 
                                    mi.Move(olBadCreditFolder)

                                    'Exit For 'Loop - Get next email
                                    GoTo _Skip ' Skip the Rest of this email

                                End If
                            Else
                                Approved = ""
                                PaymentStatus = "Not Paid" ' Need to add a SUB to alart someone that the order is not paid
                            End If
                            If chkBox_Testing.Checked = True Then
                                PaymentStatus = "Testing" ' JUST FOR TESTING, if Testing Check Box is checked
                            End If
                       ' MsgBox("Approved: " & PaymentStatus)
                        Case 16 ' Authorization Info
                            Auth_Info = s
                            If Not Auth_Info.IndexOf(":") = -1 Then
                                Dim Start As Integer = Auth_Info.IndexOf(":") + 2
                                Dim Take As Integer = Auth_Info.Length - Start
                                ' MsgBox("Start/Take: " & Start & " / " & Take)
                                Auth_Info = Auth_Info.Substring(Start, Take)
                            Else
                                Auth_Info = ""
                            End If
                       ' MsgBox("Auth_Info: " & Auth_Info)
                        Case 17 ' Comments_1  Rick Jaekle says there MAY be upto 4 comment lines
                            Comments_1 = s
                            Comments_1 = Replace(Comments_1, vbTab, "")
                            'MsgBox("Comments_1: " & Comments_1)
                            'MsgBox("Index = " + Comments_1.IndexOf(":").ToString)
                            If Not Comments_1.IndexOf(":") = -1 Then
                                Dim Start As Integer = Comments_1.IndexOf(":") + 1
                                Dim Take As Integer = Comments_1.Length - Start
                                ' MsgBox("Start/Take: " & Start & " / " & Take)
                                Comments_1 = Comments_1.Substring(Start, Take)
                                ' MsgBox("Legth =" + Comment.Length.ToString)
                            Else
                                Comments_1 = "None"
                            End If
                            Comment = Comments_1
                            '  MsgBox("Legth =" + Comment.Length.ToString)
                            If Comment.Equals("") Then
                                Comment = "None"
                            End If
                            '  MsgBox("Comments_1: " & Comment)

                            Dim input As String = Comment
                            RemoveCommaQuote(input)
                            Comment = input
                       ' MsgBox("after remove comma - Comment =" + Comment)
                        Case 18 'If more comment lines then Comments_2
                            Comments_2 = s
                            If s = "Order" Then ' No more Comment lines
                                '  MsgBox("s=Comments")
                                Comments_2 = "None"
                                orderStartsAt = 18
                                '   MsgBox("Line 18 - Not more comments, 'Order' found")
                                Exit For
                            Else
                                Comment = Comment + " - " + Comments_2
                                Dim input As String = Comment
                                RemoveCommaQuote(input)
                                Comment = input
                                '  MsgBox("Comments: " & Comment)
                            End If

                        Case 19 'If 3rd comment line then Comments_3
                            Comments_3 = s
                            If s = "Order" Then ' No more Comment lines
                                Comments_3 = ""
                                orderStartsAt = 19
                                Exit For
                            Else
                                Comment = Comment + " - " + Comments_3
                                Dim input As String = Comment
                                RemoveCommaQuote(input)
                                Comment = input
                                'MsgBox("Line 19 - Comments 3: " & Comment)
                            End If
                        Case 20  'If 4rd comment line then Comments_4
                            Comments_4 = s
                            If s = "Order" Then ' No more Comment lines
                                Comments_4 = ""
                                orderStartsAt = 20
                                Exit For
                            Else
                                Comment = Comment + " - " + Comments_4
                                Dim input As String = Comment
                                RemoveCommaQuote(input)
                                Comment = input
                                '   MsgBox("Line 20 - Comments4: " & Comment)
                            End If
                            'At this point assuming no more comments

                    End Select

                Next

                Dim TotalLine As String = emailBodyLines(emailLastLineIndex)
                ' MsgBox("Toal Line:" + TotalLine)
                If Not TotalLine.IndexOf(":") = -1 Then
                    Dim Start As Integer = TotalLine.IndexOf("$") + 1
                    Dim Take As Integer = TotalLine.Length - Start
                    TotalLine = TotalLine.Substring(Start, Take)
                    TotalLine = TotalLine.Replace(vbTab, String.Empty)
                Else
                    TotalLine = ""
                End If

                Do Until TotalLine.IndexOf(vbTab) = -1 ' remove extra Tabs
                    MsgBox("Index = " + TotalLine.IndexOf(vbTab))
                    TotalLine = TotalLine.Trim(vbTab)
                Loop

                ECommerce_DataFileId = PayTransID + ".txt"
                Dim Friday_Date As Date = Date.Now
                If Friday_Date.DayOfWeek <> DayOfWeek.Friday Then  ' if today is not Friday then use fuction to find next Friday date
                    NextFriday(Friday_Date)  'use fuction to find next Friday date
                End If
                ' MsgBox("friday Date =" + Friday_Date.ToString("yyyy-MM-dd"))
                Dim DataDirectory() As String = IO.File.ReadAllLines("C:\eCommerseIntegration\Setup Files\eCommerseIntegration email_Data Directory.txt")
                ECommerce_DataPath = DataDirectory(0) + Friday_Date.ToString("yyyy-MM-dd") + "\"
                ' MsgBox("Data Path = " + ECommerce_DataPath)
                If Not Directory.Exists(ECommerce_DataPath) Then
                    Directory.CreateDirectory(ECommerce_DataPath)
                End If

                Dim sFileName As String = ECommerce_DataPath + ECommerce_DataFileId

                ' Add Id of eCommerse Id file to check for duplicates
                Dim DataDir As String = DataDirectory(0)
                Dim Id_DataPath As String = DataDirectory(0) + myYear.ToString + ".txt"

                '  MsgBox("Data Path = " + Id_DataPath)
                If Not Directory.Exists(DataDir) Then
                    Directory.CreateDirectory(DataDir)
                End If

                If File.Exists(Id_DataPath) = False Then
                    Dim createText() As String = {"Id#, Customer, Date Processed, Folder Date", "     "}
                    IO.File.WriteAllLines(Id_DataPath, createText)
                End If
                Dim IdFile() As String = IO.File.ReadAllLines(Id_DataPath)
                If Array.Find(IdFile, Function(x) (x.StartsWith(PayTransID))) <> "" Then ' Dupe email, already in list Id's we ahve processed in the IdFile
                    ' MsgBox("Id found - Dupe email")
                    eMailContentFileName = "DUPELICATE Ecommerse Order.txt"
                    TextBox1.Text = TextBox1.Text + ("Id# " + PayTransID + " Is a Dupe - Not processing") + vbCrLf
                    ScrollBox("Text")
                    Module1.eMail_Setup(eMailContentFileName, PayTransID)  'Prepare email that an order has come in that has already been processed, a Dupe 
                    mi.Move(olBadCreditFolder)

                Else ' - a good new order
                    '  MsgBox("Id is NOT found - adding it to file ")
                    IO.File.AppendAllText(Id_DataPath, PayTransID + ", " + CustomerName + ", " + Now.ToShortDateString + ", " + Friday_Date.ToShortDateString + Environment.NewLine) ' add Id to IdFile for future checking on dupes
                    Dim Header As String = "PayTransID" + vbTab + "CustomerName" + ControlChars.Tab + "ZipCode" + vbTab + "EnteredBy" + vbTab + "TimeOrderTaken" + vbTab + "TimeExpected" + vbTab + "OrderType" + vbTab + "DeliveryType" + vbTab + "PaymentType" + vbTab
                    Header += "PaymentStatus" + vbTab + "TotalPrice" + vbTab + "SenderEmailAddress" + vbTab + "Comments"
                    Dim LineData As String = PayTransID + vbTab + CustomerName + vbTab + ZipCode + vbTab + EnteredBy + vbTab + TimeOrderTaken + vbTab + nFriday _
        + vbTab + OrderType + vbTab + DeliveryType + vbTab + PaymentType + vbTab + PaymentStatus + vbTab + TotalLine + vbTab + SenderEmailAddress
                    ' MsgBox("Combined Comments = " + Comment)
                    If Comment.Equals("") Then
                        Comment = "None"
                    End If
                    ' MsgBox("Comments = " + Comment)
                    LineData = LineData + vbTab + Comment
                    Dim sTextFile As New StringBuilder
                    sTextFile.AppendLine(Header)
                    sTextFile.AppendLine(LineData)
                    sTextFile.AppendLine(Order_Header)

                    Dim numberofOrderLines As Integer = emailLastLineIndex - orderStartsAt - 3 ' the number of lines in th eOrder section with food (sku) being ordered
                    ' MsgBox("number Of order lines = " + numberofOrderLines.ToString)
                    orderStartsAt += 2 'the email line number at which the first line of the order is to be found
                    ' emailLastLineIndex ' the last line of the email
                    For order_line As Integer = orderStartsAt To (numberofOrderLines + orderStartsAt - 1) ' looping though each line that has an order
                        Dim Order_Detail As String = emailBodyLines(order_line) ' a single line of the detail order
                        Dim desc As String = Order_Detail
                        'Order_Detail = Order_Detail.Replace(vbTab, ",") ' take out the tabs and replace with a comma
                        Order_Detail = Order_Detail.Replace("$", "")
                        'MsgBox("Indexof ',' in Order Detail :" + Order_Detail.IndexOf(",").ToString)
                        'Order_Detail = Order_Detail.Insert(Order_Detail.IndexOf(vbTab), (vbTab + "Sku#")) ' REMOVE THIS LINE ONE THE SKU NUMBE IS ADDED O DATA BY RICK
                        Dim lgth As Integer = Order_Detail.Length
                        Order_Detail = Order_Detail.Substring(0, lgth - 1)

                        sTextFile.AppendLine(Order_Detail)
                    Next

                    IO.File.AppendAllText(sFileName, sTextFile.ToString)

                    mi.Move(olDestinationFolder) ' move mail item to todays folder in Outlook
                End If


            Catch ex As System.IO.IOException
                ' Code that reacts to IOException.
            Catch ex As NullReferenceException
                Module1.eMail_Setup("System Down Unknown Error - restart system", "")  'We have found an error not recognized so we are shutting down and sending an email.  Send email warning someone the program has stopped

                MessageBox.Show("NullReferenceException: " & ex.Message)
                MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

                'If mi Is Nothing Then
                Module1.eMail_Setup("System Down Bad Mail Object.txt", "")  'We have found an error not recognized so we are shutting down and sending an email.  Send email warning someone the program has stopped
                'Me.Close()
                System.Windows.Forms.Application.Exit()
                ' MsgBox("mi (Mail Item) = 'Nothing', shutting down")
                End
                'End If
            End Try
_Skip:
            Me.TextBox1.Text = Me.TextBox1.Text + "Processed " + PayTransID + " for customer " + CustomerName + vbCrLf
            ScrollBox("Text")
            Me.Show()
            Me.Refresh()
        Next 'get next email
        TextBox1.Text = TextBox1.Text + "Listening for New emails!" + vbCrLf
        ScrollBox("Text")
        _tMonitorOutLook()
        ' Cleanup()

    End Sub

    Public Shared Function RemoveCommaQuote(ByRef input As String) As String
        '  MsgBox("input" + input)
        If input.Contains(",") = True Then
            input = input.Replace(",", " ")
        End If

        Dim ch As Char
        ch = ChrW(34) 'the " (double quote) character
        If input.Contains(ch) = True Then
            input = input.Replace("ch", " ")
        End If
        '  MsgBox("input" + input)
        Return (input)
    End Function

    'Public Shared Function GetSku(ByVal desc As String)
    '    'MsgBox("desc = " + desc)
    '    Dim Sku As String = ""

    '    Dim connString As String = "Data Source = .\SQLEXPRESS; Database = Seafood; Integrated Security=SSPI"
    '    Dim conn As New SqlConnection(connString)
    '    If conn.State = 0 Then
    '        conn.Open()

    '    End If
    '    '  // Create New DataAdapter
    '    Dim a As New SqlDataAdapter("SELECT sku, Price, Description From ItemTypes Where (DisplayAttrib = 1) Order By DisplayOrder", conn)
    '    ' // Use DataAdapter to fill DataTable
    '    Dim comm As New SqlCommand("SELECT sku, Price, Description From ItemTypes Where (DisplayAttrib = 1) Order By DisplayOrder", conn)
    '    Dim dt As New DataTable
    '    a.Fill(dt)


    '    If Not desc.IndexOf(vbTab) = -1 Then
    '        Dim Start As Integer = desc.IndexOf(vbTab) + 1
    '        Dim Take As Integer = desc.Length - Start
    '        desc = desc.Substring(Start, Take)
    '        Dim end_of_Desc As Integer = desc.IndexOf(vbTab)
    '        desc = desc.Substring(0, end_of_Desc)
    '        '  MsgBox("desc = " + desc)
    '    Else
    '        desc = ""
    '    End If
    '    'Dim temp As String = "Description = Beer Battered Filets"

    '    ' MsgBox("desc = " + desc)

    '    Dim Result() As DataRow = dt.Select("Description = " + desc)
    '    '("Description = 'Beer Battered Filets'")
    '    For Each row As DataRow In Result
    '        If (row.IsNull("Description")) Then
    '            Form1.TextBox1.Text = Form1.TextBox1.Text + " Result Is Empty - No Description in Order line (program line 590)" + vbCrLf
    '            Form1.ScrollBox("Text")
    '        End If
    '        Sku = row("Sku").ToString

    '        '    MsgBox("Sku = " + Sku)

    '    Next row
    '    conn.Close()
    '    Return (Sku)
    'End Function
    'Public Sub RefillInbox()
    '    If rerun = 0 Then
    '        CreateFolder()
    '    End If
    '    For i = olDestinationFolder.Items.Count To 1 Step -1
    '        mi = olDestinationFolder.Items(i)
    '        mi = TryCast(mi, MailItem)
    '        mi.Move(olSourceFolder)
    '    Next
    '    MsgBox("Inbox refilled")
    '    Cleanup()

    'End Sub

    'Private Sub BtnRefillInbox_Click(sender As Object, e As EventArgs)
    '    RefillInbox()
    'End Sub

    Private Sub BtnProcessNew_Click(sender As Object, e As EventArgs) Handles BtnProcessNew.Click
        'MessageBox.Show("Start listening to Outlook mail!")
        TextBox1.Text = TextBox1.Text + "Program is Designed to Run Continuously." + vbCrLf
        TextBox1.Text = TextBox1.Text + "Listening for New emails!" + vbCrLf
        ScrollBox("Text")
        BtnProcessNew.Enabled = False
        BtnProcessNew.Text = "Running"
        _tMonitorOutLook() ' watch for new email
    End Sub
    'Private Sub btnExit_MouseEnter(sender As Object, e As EventArgs) Handles btnExit.MouseEnter
    '    Cursor = Cursors.Default
    '    ' BtnProcessNew.FontSize = 10
    'End Sub

    'Private Sub BtnExit_MouseLeave(sender As Object, e As EventArgs) Handles btnExit.MouseLeave
    '    If Cursor <> Cursors.Default Then
    '        Cursor = Cursors.WaitCursor
    '    End If
    '    ' BtnProcessNew.FontSize = 8
    'End Sub

    'Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    'End Sub

    Private Sub chkBox_Testing_CheckedChanged(sender As Object, e As EventArgs) Handles chkBox_Testing.CheckedChanged
        If chkBox_Testing.Checked = True Then
            chkBox_Testing.BackColor = Color.Red
            Label2.BackColor = Color.Red
            chkBox_Testing.Text = "Testing"
        Else
            chkBox_Testing.BackColor = Control.DefaultBackColor
            Label2.BackColor = Control.DefaultBackColor
            chkBox_Testing.Text = "Not Testing"
        End If

    End Sub

    'Private Sub Email_TxtBox_TextChanged(sender As Object, e As EventArgs)
    '    '    Email_TxtBox.Text = Email_TxtBox.Text
    'End Sub

    Private Sub Cleanup()

        'Dim result As DialogResult = MessageBox.Show("Close App?", "Refill?", MessageBoxButtons.YesNo)
        ' If (result = DialogResult.Yes) Then
        outLookApp = Nothing
        outLookNS = Nothing
        olSourceFolder = Nothing
        mi = Nothing
        olDestinationFolder = Nothing
        Me.Dispose()
        System.Windows.Forms.Application.Exit()
        End ' End the program
        ' End If

    End Sub

    Private Sub BtnCleanup_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Cleanup()
        Me.Dispose()
        System.Windows.Forms.Application.Exit()
        End ' End the program
    End Sub

    'Private Sub PrintListBox_Click(sender As Object, e As EventArgs) Handles PrintListBox.Click
    '    MsgBox("ListBox Items = " & ListBox1.Items.Count) 'there are 353 items in listbox1 - only 52 are printing
    '    Dim PrintDialog1 As New PrintDialog
    '    Dim result As DialogResult = PrintDialog1.ShowDialog()
    '    If result = DialogResult.OK Then DocToPrint.Print()
    'End Sub

    'Private Sub PrintDocument1_PrintPage(sender As System.Object, e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
    '    Dim mRow As Integer = 0
    '    Dim newpage As Boolean = True
    '    With DataGridView1
    '        Dim fmt As StringFormat = New StringFormat(StringFormatFlags.LineLimit)
    '        fmt.LineAlignment = StringAlignment.Center
    '        fmt.Trimming = StringTrimming.EllipsisCharacter
    '        Dim y As Single = e.MarginBounds.Top
    '        Do While mRow < .RowCount
    '            Dim row As DataGridViewRow = .Rows(mRow)
    '            Dim x As Single = e.MarginBounds.Left
    '            Dim h As Single = 0
    '            For Each cell As DataGridViewCell In row.Cells
    '                Dim rc As RectangleF = New RectangleF(x, y, cell.Size.Width, cell.Size.Height)
    '                e.Graphics.DrawRectangle(Pens.Black, rc.Left, rc.Top, rc.Width, rc.Height)
    '                If (newpage) Then
    '                    e.Graphics.DrawString(DataGridView1.Columns(cell.ColumnIndex).HeaderText, .Font, Brushes.Black, rc, fmt)
    '                Else
    '                    e.Graphics.DrawString(DataGridView1.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.Black, rc, fmt)
    '                End If
    '                x += rc.Width
    '                h = Math.Max(h, rc.Height)
    '            Next
    '            newpage = False
    '            y += h
    '            mRow += 1
    '            If y + h > e.MarginBounds.Bottom Then
    '                e.HasMorePages = True
    '                mRow -= 1
    '                newpage = True
    '                Exit Sub
    '            End If
    '        Loop
    '        mRow = 0
    '    End With
    'End Sub

    'Private Sub BtnPrint_Click(sender As Object, e As EventArgs) Handles BtnPrint.Click
    '    PrintPreviewDialog1.Document = PrintDocument1
    '    PrintPreviewDialog1.ShowDialog()
    'End Sub

    '  Public WithEvents DocToPrint As New PrintDocument ' using to print ListBox

    'Public Sub Document_PrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs) Handles DocToPrint.PrintPage
    '    Dim printFont As New Font("Arial", 10, System.Drawing.FontStyle.Regular)
    '    Dim YPosition As Integer = 20
    '    For Each eItem As String In ListBox1.Items 'there are 353 items in listbox1 - only 52 are printing 
    '        e.Graphics.DrawString(eItem, printFont, System.Drawing.Brushes.Black, 25, YPosition)
    '        YPosition += 20
    '    Next

    'End Sub

    Public Sub ScrollBox(ByVal Box As String)
        If Box = "Text" Then
            TextBox1.SelectionStart = TextBox1.Text.Length
            TextBox1.ScrollToCaret()
        End If
        If Box = "List" Then
            ListBox1.TopIndex = ListBox1.Items.Count - 1
        End If
    End Sub

End Class
