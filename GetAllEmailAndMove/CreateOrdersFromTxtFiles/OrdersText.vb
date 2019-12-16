'Convert Txt Files to Seafood Orders in Database
Imports System
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Imports CommonSubs

Module OrdersText
    Dim sqlcon As SqlConnection
    Dim Location As String = Where() 'Find out if code is running at home or at the Church, change Connection string using "commonSettings.config" and Location var
    Dim New_ID As Integer
    Dim new_Seq As Nullable(Of Integer)
    Dim tTransaction As SqlTransaction

    Sub Main()
        ' Msg("common Sub Where = ", Location)
        ' MsgBox("holding just for testing")
        OpenSqlConnection(Location) ' find location of computer via its OS and open a connection to database
        ' Msg("Connection state =", sqlcon.State.ToString)
        GetWeekNumber()
        ConvertToOrder()
    End Sub

    Sub ConvertToOrder()

        'Create Directories - sourceDirectory and  

        Dim sourcePath As String = GetFolderDate()
        ' sourcePath = "C:\eCommerseIntegration\email_Data\2019-12-06\" '****FOR TESTING ONLY*****
        Dim archivePath As String = sourcePath + "archive\"
        Dim archiveDirectory As DirectoryInfo = Directory.CreateDirectory(sourcePath + "archive\")

        ' Read each file.txt in directory

        Dim txtFiles = Directory.EnumerateFiles(sourcePath, "*.txt")
        For Each currentFile As String In txtFiles
            Dim fileName As String = currentFile.Substring(sourcePath.Length)
            Dim fileLines As String() = IO.File.ReadAllLines(sourcePath + fileName)

            '-------------------------------------------------------------------------------------------------------------------------
            'Fill the OrderHearer Data Variables

            Dim arr As String() = SplitWords(fileLines(1)) ' Line 1 is the actual OrderHeader DAT; Use the Subroutine to break apart on TAB char.
            Dim PayTransID As String = arr(0)
            Dim CustomerName As String = arr(1)
            Dim ZipCode As String = arr(2)
            Dim EnteredBy As String = "Middleware"
            Dim TimeOrderTaken As String = arr(4)
            Dim TimeExpected As String = arr(5)
            Dim OrderType As String = arr(6)
            Dim DeliveryType As String = arr(7)
            Dim PaymentType As String = arr(8)
            Dim PaymentStatus As String = arr(9)
            Dim TotalPrice As String = arr(10)
            Dim SenderEmailAddress As String = arr(11) ' not used in database today
            Dim Comments As String = arr(12) ' not used in database today
            Dim TimeRequested As String = arr(5)
            Dim Discount As Integer = 0
            Dim Week As Integer = GetWeekNumber()
            Dim myYear As Integer = DatePart("yyyy", Now())

            'Insert the data into a new OrderHeader record

            Dim tTransaction As SqlTransaction = sqlcon.BeginTransaction()
            Try
                ' Create the OrderHearder Record
                Dim sqlcmd1 As New SqlCommand With {
                    .CommandText = "INSERT INTO OrderHeader " _
            & "(PayTransID,CustomerName,ZipCode,EnteredBy,TimeOrderTaken,TimeExpected,OrderType,DeliveryType, PaymentType,PaymentStatus,TotalPrice,TimeRequested,Discount,Year,Week)" _
            & "Values (@PayTransID, @CustomerName, @ZipCode, @EnteredBy, @TimeOrderTaken, @TimeExpected, @OrderType, @DeliveryType, @PaymentType, @PaymentStatus, @TotalPrice, @TimeRequested, @Discount, @myYear, @Week)"
                }
                With sqlcmd1.Parameters
                    .AddWithValue("@PayTransID", PayTransID)
                    .AddWithValue("@CustomerName", CustomerName)
                    .AddWithValue("@ZipCode", ZipCode)
                    .AddWithValue("@EnteredBy", EnteredBy)
                    .AddWithValue("@TimeOrderTaken", TimeOrderTaken)
                    .AddWithValue("@TimeExpected", TimeExpected)
                    .AddWithValue("@OrderType", OrderType)
                    .AddWithValue("@DeliveryType", DeliveryType)
                    .AddWithValue("@PaymentType", PaymentType)
                    .AddWithValue("@PaymentStatus", PaymentStatus)
                    .AddWithValue("@TotalPrice", TotalPrice)
                    .AddWithValue("@TimeRequested", TimeRequested)
                    .AddWithValue("@Discount", Discount)
                    .AddWithValue("@myYear", myYear)
                    .AddWithValue("@Week", Week)
                End With
                sqlcmd1.Connection = sqlcon
                sqlcmd1.Transaction = tTransaction
                sqlcmd1.ExecuteNonQuery()

                ' Get the next Order number 
                Dim cmdGetID As SqlCommand = New SqlCommand("SELECT MAX(OrderId) FROM [Seafood].[dbo].[OrderHeader]") With {
                    .Connection = sqlcon,
                    .Transaction = tTransaction
                     }
                New_ID = cmdGetID.ExecuteScalar()
                '-------------------------------------------------------------------------------------------------------------------------
                ' Get next "Seq" number, starting at 1 each week
                Dim SeqText As String = "SELECT MAX([Seq])FROM [Seafood].[dbo].[OrderHeader] Where Year = " & myYear & " And Week = " & Week
                Dim cmdGetSeq As SqlCommand = New SqlCommand(SeqText) With {
                        .Connection = sqlcon,
                        .Transaction = tTransaction
                    }
                'Dim new_Seq As Nullable(Of Integer) '= cmdGetSeq.ExecuteScalar()
                If (IsDBNull(cmdGetSeq.ExecuteScalar())) Then  'The Seq will be Null if it is the first Order of the new week
                    new_Seq = 1
                Else
                    new_Seq = cmdGetSeq.ExecuteScalar() + 1
                End If

                'Now Update the new Order with the Seq

                Dim sqlcmd2 As New SqlCommand()
                sqlcmd2.CommandText = ("UPDATE OrderHeader Set Seq = @Seq Where [OrderId] = @NewId")
                With sqlcmd2.Parameters
                    .AddWithValue("@Seq", new_Seq)
                    .AddWithValue("@NewId", New_ID)
                End With
                sqlcmd2.Connection = sqlcon
                sqlcmd2.Transaction = tTransaction
                sqlcmd2.ExecuteNonQuery()

                '------------------------------------------------------------------------------------------------------------------------------------------------

                ' Add the OrderNumber to the text File of the eMail, just for manual cross referance

                fileLines(0) = fileLines(0) + (vbTab + "OrderID") ' the lables for the Header
                fileLines(1) = fileLines(1) + (vbTab + New_ID.ToString) ' add the ID used to insert into Database
                System.IO.File.WriteAllLines(currentFile, fileLines)

                '----------------------------------------------------------------------------------------------------------------------------------------------------

                ' Insert individual OrderItems for this Order - what the customer ordered, line by line

                Dim OrderItems As String() 'an array for each order Item line
                For orderItemLine = 3 To (fileLines.Count - 1)  'Line3 of each text file are individual OrderItems
                    OrderItems = SplitWords(fileLines(orderItemLine))
                    'Array (0) = Qty; (1) = Sku; (2) = Descriton - not used; (3) = Unit Price; (4) = Total Price- not used; 
                    Dim OrderId = New_ID
                    If Not (IsNumeric(OrderItems(1))) Then  ' a sku is missing
                        ' Msg("Sku is a String", "")
                        tTransaction.Rollback() '*****Need email to tell someone*******
                        GoTo Move  ' move the order to archieve
                    End If
                    Dim QtyFree As Integer = 0
                    Dim Sku As Integer = OrderItems(1)
                    Dim UnitPrice As Decimal = OrderItems(3)
                    Dim QtyTotal = OrderItems(0) ' qty ordered by customer
                    Dim Description As String = OrderItems(2)

                    '---------------------------------------------------------------------------------------------------------------------------------------
                    ' Check If we have invetory and reduce it if we do
                    ' If inventory is 0 or less, set Qty to 0, send email to warn customer that item is not in stock
                    Dim cmdGetQtyInventory As SqlCommand = New SqlCommand("SELECT [QtyInventory] FROM [Seafood].[dbo].[ItemTypes] Where [Sku] = " + Sku.ToString) With {
                        .Connection = sqlcon,
                        .Transaction = tTransaction
                    }
                    Dim QtyInventory As Integer = 0
                    QtyInventory = cmdGetQtyInventory.ExecuteScalar()
                    ' Console.Write("Sku - Disc - QtyInventory: " + Sku.ToString + " - " + Description + " - " + QtyInventory.ToString)
                    ' Msg(" ", " ")
                    ' Console.ReadKey()
                    'Console.Write("Stopping it")
                    If (QtyInventory - QtyTotal) <= 0 Then
                        '****Setting the one item that is out of Inventory to 0 - place the rest of the order, SEND AN EMAIL TO STAFF
                        Dim SkuOutOfStock As String = Sku
                        Dim QtyOutOfStock As String = Decimal.Parse(QtyTotal)
                        Dim UnitPriceOutofStock As Decimal = Decimal.Parse(UnitPrice)
                        Dim CustomerCredit As Decimal = (Decimal.Parse(QtyTotal) * Decimal.Parse(UnitPrice))
                        Dim eMailText As String = "eCom# " + PayTransID + " Out Of Stock Item = " + Description.ToString + "; Qty Ordered = " + QtyOutOfStock.ToString + "; Unit Price of Out of Stock Item = $" + "; " + UnitPriceOutofStock.ToString + "; Total Customer $ Credit Owed = $" + CustomerCredit.ToString

                        CommonSubs.eMail_Setup("Items Not in Stock - Customer Credit.txt", eMailText) ' send email to staff telling them  we dhorted an order

                        QtyTotal = 0 ' change Qty the customer ordered to 0 - don;t place the order - ***Need email Alert
                        UnitPrice = 0 ' setting unit proce to zero just for this line

                        'Else ' Tony said not to adjust stock that his program will do that.
                        '    QtyInventory = QtyInventory - QtyTotal

                        '    'Try ' 'Update the The database with Qty left in Inventory
                        '    Dim sqlcmdInv As New SqlCommand()
                        '    sqlcmdInv.Connection = sqlcon
                        '    sqlcmdInv.Transaction = tTransaction
                        '    sqlcmdInv.CommandText = ("UPDATE [Seafood].[dbo].[ItemTypes] Set QtyInventory = @QtyInv Where [Sku] = " + Sku.ToString)
                        '    With sqlcmdInv.Parameters
                        '        .AddWithValue("@QtyInv", QtyInventory)
                        '    End With
                        '    sqlcmdInv.ExecuteNonQuery()
                    End If

                    '-----------------------------------------------------------------------------------------------------------------------------------------
                    'Set the Location where the item will be served from: Rack, Kitchen, Hall. Obtained from the Sku file

                    Dim Location As String = ""
                    ' Try
                    Dim cmdGetLocation As SqlCommand = New SqlCommand("SELECT [Location] FROM [Seafood].[dbo].[ItemTypes] Where [Sku] = " + Sku.ToString) With {
                        .Connection = sqlcon,
                        .Transaction = tTransaction
                    }
                    Location = cmdGetLocation.ExecuteScalar()
                    If Location = "" Then
                        Location = "Kitchen"
                    End If

                    Dim sqlcmd3 As New SqlCommand
                    sqlcmd3.CommandText = "INSERT INTO OrderItems " _
                    & "(OrderId,Sku,QtyTotal,QtyFree,UnitPrice,Location)" _
                    & "Values (@OrderId, @Sku, @QtyTotal, @QtyFree, @UnitPrice, @Location)"
                    With sqlcmd3.Parameters
                        .AddWithValue("@OrderId", OrderId)
                        .AddWithValue("@Sku", Sku)
                        .AddWithValue("@QtyTotal", QtyTotal)
                        .AddWithValue("@QtyFree", QtyFree)
                        .AddWithValue("@UnitPrice", UnitPrice)
                        .AddWithValue("@Location", Location)
                    End With
                    sqlcmd3.Connection = sqlcon
                    sqlcmd3.Transaction = tTransaction
                    sqlcmd3.ExecuteNonQuery()

                    '--------------------------------------------------------------------------------------------------------------
                Next orderItemLine  'Go process the next detail Item Line if any

                '--------------------------------------------------------------------------------------------------------------
                'All items are process, commit the orderHeader and the OrderItems

                tTransaction.Commit()
                ' Catch any errors
            Catch ex2 As System.Exception
                ' This catch block will handle any errors that may have occurred
                ' on the server that would cause the rollback to fail, such as
                ' a closed connection.
                tTransaction.Rollback()
                sqlcon.Close()
                sqlcon.Dispose()

                '********** Need to write to Error file and send email******************
                Console.WriteLine("Rollback Exception : {0}", ex2.GetType())
                Console.WriteLine("  Message: {0}", ex2.Message)
                Console.WriteLine("Error in converting Txt file to Create SQL Database records - entire Transaction cancelled - move offending ecommerce record ecomId = " + PayTransID)
                CommonSubs.WriteErrorFile("Error in Creating SQL Database records, entire Transaction cancelled - move offending ecommerce record", "ecomId = " _
                                          + PayTransID, "Error in converting Txt file.txt")
            End Try
            '-------------------------------------------------------------------------------------------------------
Move:
            ' Move the e-com order to the "archive" directory
            Directory.Move(currentFile, Path.Combine(archivePath, fileName))

        Next  'currentFile

    End Sub
    'Split line on Tab Char into individual variable
    Private Function SplitWords(ByVal s As String) As String()
        Return s.Split(New Char() {vbTab})
    End Function
    Function NextFriday(ByRef Tdate) As String
        Tdate = Tdate.AddDays(1)
        Do While Tdate.DayOfWeek <> DayOfWeek.Friday
            Tdate = Tdate.AddDays(1)
        Loop
        Return Tdate
    End Function
    Private Function GetWeekNumber() As Integer ' get week number of the year
        Dim fldDate As Date = Now()
        Dim WeekNumber As Integer = DatePart("ww", fldDate)
        Return WeekNumber
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

    Public Sub OpenSqlConnection(Location) ' use the Location to read the correct Connection String from commonSettings.config
        Dim MyConnectionString As String = ConfigurationManager.AppSettings.Get("Seafood" + Location)
        ' MyConnectionString = ConfigurationManager.ConnectionStrings("conString").ConnectionString   ' from App.config
        sqlcon = New SqlConnection(MyConnectionString)
        Try
            sqlcon.Open()
        Catch ex As System.Exception
            CommonSubs.WriteErrorFile("Tried and failed to open SQL Connection.  Error: " + ex.ToString, " ", "Error - SQL Connection.txt")
            Msg("error opening Connection, error = " + vbCrLf, ex.ToString)
            Msg("System is shutting down!!!", "")
            Cleanup(1)

        End Try
    End Sub

    'Public Sub Msg(text As String, var As String)
    '    Console.Write(text + " " + var)
    '    Console.ReadLine()
    'End Sub
    'Public Function Where() ' Determine where the code is being run so the correct Connection String can be loaded
    '    Dim Location As String = ""
    '    Select Case My.Computer.Info.OSFullName
    '        Case "Microsoft Windows Server 2008 R2 Standard" ' code is running a Church
    '            Location = "Church"
    '        Case "Microsoft Windows 10 Home" ' Code is running at Home
    '            Location = "Home"
    '        Case Else ' code is not at home or church, we will not have Connection String
    '            Dim errTxt As String = "CreateOrdersFromText program SHUTTING DOWN.  Can not find correct Location/Computer code - Location unkown!!" + vbCrLf + "Check, did someone change the OS of the computer?  This Program is shutting down!"
    '            Dim OrderId As String = "N/A"
    '            Dim eMailContentFileName As String = ConfigurationManager.AppSettings.Get("SetupDir") + "Error - ComputerLocationUnknown.txt"
    '            WriteErrorFile(errTxt, OrderId, eMailContentFileName)

    '            End
    '    End Select
    '    Return Location
    'End Function
    'Sub WriteErrorFile(errTxt As String, OrderId As String, eMailContentFileName As String)
    '    Try
    '        Dim errFilePath As String = ConfigurationManager.AppSettings.Get("ErrorFile")
    '        IO.File.AppendAllText(errFilePath, vbCrLf + Now.ToUniversalTime + vbCrLf + errTxt + vbCrLf)
    '        eMail_Setup(eMailContentFileName, OrderId)  'Prepare email that this program is shutting down, most likely because Outlook is NOT running 
    '    Catch ex As System.Exception
    '        eMailContentFileName = "Tried to write error file and failed - unknown error.txt" + vbCrLf + ex.ToString
    '        eMail_Setup(eMailContentFileName, OrderId)  'Prepare email that this program is shutting down, most likely because Outlook is NOT running 
    '    End Try
    'End Sub
    Private Sub Cleanup(exitcode As Integer)
        sqlcon.Close()
        sqlcon.Dispose()
        Msg("System is shutting down!!!", "")
        System.Environment.Exit(exitcode) ' End the program (need to document error 1 as SqlCon failed
        End
    End Sub
    'Sub eMail_Setup(eMailContentFileName As String, OrderId As String)

    '    ' Dim EmailDirectory() As String = IO.File.ReadAllLines("C:\eCommerseIntegration\Setup Files\Ecommerse Setup Directory.txt")
    '    'Dim Directory As String = EmailDirectory(0)
    '    Dim emailContent As String() = IO.File.ReadAllLines(eMailContentFileName)
    '    Dim sSubject As String = emailContent(0) + "   ----  OrderId = " + OrderId
    '    Dim sBody As String = emailContent(2) + "   ----  OrderId = " + OrderId
    '    Dim sTo As String = emailContent(3) ' add Rick Jeakle and Greg Dohner
    '    Dim sCC As String = emailContent(4) ' seafood@jaeklegroup.net
    '    Dim sFilename As String = emailContent(5)
    '    Dim sDisplayname As String = emailContent(6)
    '    sEmailSend(sSubject, sBody, sTo, sCC, sFilename, sDisplayname)

    'End Sub
    'Sub sEmailSend(sSubject As String, sBody As String,
    '                         sTo As String, sCC As String,
    '                         sFilename As String, sDisplayname As String)
    '    Dim oApp As _Application
    '    oApp = New Application

    '    Dim oMsg As Microsoft.Office.Interop.Outlook._MailItem
    '    oMsg = oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
    '    oMsg.Subject = sSubject
    '    oMsg.Body = sBody
    '    oMsg.To = sTo
    '    oMsg.CC = sCC

    '    Dim strS As String = sFilename
    '    Dim strN As String = sDisplayname
    '    If sFilename <> "" Then
    '        Dim sBodyLen As Integer = Int(sBody.Length)
    '        Dim oAttachs As Microsoft.Office.Interop.Outlook.Attachments = oMsg.Attachments
    '        Dim oAttach As Microsoft.Office.Interop.Outlook.Attachment
    '        oAttach = oAttachs.Add(strS, , sBodyLen, strN)
    '    End If
    '    oMsg.Send()
    '    ' MsgBox("Email Sent")

    'End Sub
End Module
