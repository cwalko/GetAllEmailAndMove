Imports System.Data.SqlClient
Imports System.IO

Module Module1
    Public PayTransID As String = Form1.PayTransID
    Public Sub Insert_to_SQL(SQLDB, PayTransID, CreationTime, SentOn, ReceivedTime, SenderName, SenderEmailAddress, emailBody)
        ' MsgBox(SQLDB & " - " & PayTransID  & " - " & CreationTime & " - " & SentOn & " - " & ReceivedTime & " - " & SenderName & " - " & SenderEmailAddress)
        Dim MyConnection As SqlConnection
        Dim MyConnectionString As String
        Dim cmd As New SqlCommand
        MyConnectionString = "Data Source = .\SQLEXPRESS; Database = " & SQLDB & "; Integrated Security=SSPI"
        '"Data Source = DESKTOP-11TM25R\SQLEXPRESS; AttachDbFilename=" & SQLDB & "; Integrated Security=True; Connect Timeout=30"   '; User Instance=True"
        MyConnection = New SqlConnection(MyConnectionString)
        ' MsgBox("b4 opening Connection State: " & MyConnection.State.ToString)
        Try
            MyConnection.Open()
            Form1.Label1.Text = MyConnection.State.ToString
            ' MsgBox("after opening Connection State: " & MyConnection.State.ToString)
        Catch Err As Exception
            Dim ErrorDirectory() As String = IO.File.ReadAllLines("C:\eCommerseIntegration\Setup Files\eCommerseIntegration Error Files Directory.txt")
            System.IO.File.AppendAllText(ErrorDirectory(0) + "\Error.txt", vbCrLf & vbCrLf & Now & vbCrLf & " -  Error Opening SQL Connection at line 16 of Module1 of DB =  " & SQLDB & vbCrLf & Err.ToString)
            'MyConnection.Close()
            'MyConnection.Dispose()
            '' Finally
            Dim ErrorFile As String = "Error - SQL Connection.txt"
            Call WriteErrorFile(Errorfile, PayTransID)  'Update the textbox with "Error Opening SQL Table" and error
            Module1.eMail_Setup("ErrorFile", "")  'Send email warning someone the program has stopped
            ' SHUTTING DOWN THE BRIDGE SYTEM
            MyConnection.Close()
            MyConnection.Dispose()
            Form1.Dispose()
            MsgBox("SYSTEM HAS BEEN SHUT DOWN! - SQL CONNECTION WILL NOT OPEN on Line 16 of Module 1")
            System.Windows.Forms.Application.Exit()
            End ' End the program
        End Try
        If Form1.Label1.Text = "Open" Then
            Try
                cmd.Connection = MyConnection
                'Form1.txtboxConnectionState.Text = MyConnection.State.ToString

                'A check if SQL Order [PayTransID] is already in Database

                Dim theQuery As String = "SELECT * FROM eComEmailBody WHERE PayTransID=@PayTransID "
                Dim cmd1 As SqlCommand = New SqlCommand(theQuery, MyConnection)
                cmd1.Parameters.AddWithValue("@PayTransID ", PayTransID)
                ' MsgBox(cmd1.Parameters.ToString)
                Using reader As SqlDataReader = cmd1.ExecuteReader()
                    If reader.HasRows Then
                        ' Order already exists
                        Dim ErrorDirectory() As String = IO.File.ReadAllLines("C:\eCommerseIntegration\Setup Files\eCommerseIntegration Error Files Directory.txt")
                        System.IO.File.AppendAllText(ErrorDirectory(0) + "\Error.txt",
                             vbCrLf & vbCrLf & Now & vbCrLf & " -  DUPLICATE - Tried to Add eCom# " & PayTransID & "to " & SQLDB & "IT IS ALREADY IN DATABASE" & vbCrLf & vbCrLf & Err.ToString)
                        reader.Close()
                        Dim errorFile As String = "DUPELICATE Ecommerse Order.txt"
                        Call Module1.WriteErrorFile(errorFile, PayTransID)  'Update the textbox with "IT IS ALREADY IN DATABASE" and error
                        GoTo Line65

                    Else
                        ' Order # is new (does not exist), add it
                        reader.Close()
                        cmd.CommandText = "INSERT INTO eComEmailBody " _
                            & "(PayTransID, emailBody, CreationTime,   SentOnTime,  ReceivedTime, SenderEmailAddress)" _
                            & "VALUES (@PayTransID , @emailBody, @CreationTime, @SentOnTime, @ReceivedTime, @SenderEmailAddress)"

                        With cmd.Parameters
                            .AddWithValue("@PayTransID ", PayTransID)
                            .AddWithValue("@emailBody", emailBody)
                            .AddWithValue("@CreationTime", CreationTime)
                            .AddWithValue("@SentOnTime", SentOn)
                            .AddWithValue("@ReceivedTime", ReceivedTime)
                            .AddWithValue("@SenderName", SenderName)
                            .AddWithValue("@SenderEmailAddress", SenderEmailAddress)
                        End With

                        cmd.ExecuteNonQuery()
                    End If

                End Using
Line65:

            Catch err As Exception
                Dim ErrorDirectory() As String = IO.File.ReadAllLines("C:\eCommerseIntegration\Setup Files\eCommerseIntegration Error Files Directory.txt")
                System.IO.File.AppendAllText(ErrorDirectory(0) + "\Error.txt", vbCrLf & vbCrLf & Now & vbCrLf & vbCrLf & " -  Error Adding Data to File  " & SQLDB & vbCrLf & err.ToString)
                'MessageBox.Show("Error while inserting record on table..." & err.Message, "Insert Records")
                MyConnection.Close()
                Form1.Label1.Text = MyConnection.State.ToString
                MyConnection.Dispose()
            Finally
                MyConnection.Close()
                Form1.Label1.Text = MyConnection.State.ToString
                MyConnection.Dispose()
            End Try

            'display state
            Form1.Label1.Text = MyConnection.State.ToString
            MyConnection.Close()
            MyConnection.Dispose()
            Form1.Label1.Text = MyConnection.State.ToString
            ' Call Module1.WriteErrorFile(PayTransID)  'Update the textbox with the data and error
        End If

    End Sub

    Public Sub WriteErrorFile(errorFile, PayTransID)
        Try

            Dim ErrorDirectory() As String = IO.File.ReadAllLines("C:\eCommerseIntegration\Setup Files\eCommerseIntegration Error Files Directory.txt")
            'Dim errorFile As String
            Dim readFile As System.IO.TextReader = New StreamReader(ErrorDirectory(0) + "\Error.txt")
            errorFile = readFile.ReadToEnd
            Form1.TextBox1.AppendText(errorFile)
            Form1.ScrollBox("Text")
            readFile.Close()
            readFile = Nothing
            Dim eMailContentFileName As String = errorFile
            Module1.eMail_Setup(eMailContentFileName, PayTransID)  'Prepare email that this program is shutting down, most likely because Outlook is NOT running 
        Catch ex As IOException
            Dim eMailContentFileName As String = "unknown error.txt  Error = " + ex.ToString
            Module1.eMail_Setup(eMailContentFileName, PayTransID)  'Prepare email that this program is shutting down, most likely because Outlook is NOT running 
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Sub eMail_Setup(fileName As String, PayTransID As String)

        Dim EmailDirectory() As String = IO.File.ReadAllLines("C:\eCommerseIntegration\Setup Files\Ecommerse Setup Directory.txt")
        Dim Directory As String = EmailDirectory(0)
        Dim emailContent As String() = IO.File.ReadAllLines(Directory + fileName)
        Form1.TextBox1.Text = Form1.TextBox1.Text + emailContent(1) + PayTransID + vbCrLf
        Form1.ScrollBox("Text")
        Dim sSubject As String = emailContent(0) + "   ----  Id = " + PayTransID
        Dim sBody As String = emailContent(2) + "   ----  PayTransID = " + PayTransID
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


    'https://www.aspsnippets.com/Articles/SqlBulkCopy-Bulk-Copy-data-from-DataTable-DataSet-to-SQL-Server-Table-using-C-and-VBNet.aspx
    '    Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
    '        If Not Me.IsPostBack Then
    '            Dim ds As New DataSet()
    '            ds.ReadXml(Server.MapPath("~/Customers.xml"))
    '            GridView1.DataSource = ds.Tables(0)
    '            GridView1.DataBind()
    '        End If
    '    End Sub

    '    Sub Bulk_Insert(sender As Object, e As EventArgs)
    '         Dim dt As New DataTable()
    '        dt.Columns.AddRange(New DataColumn(2) {New DataColumn("Id", GetType(Integer)), New DataColumn("Name", GetType(String)), New DataColumn("Country", GetType(String))})
    '        For Each row As GridViewRow In GridView1.Rows
    '            If TryCast(row.FindControl("CheckBox1"), CheckBox).Checked Then
    '                Dim id As Integer = Integer.Parse(row.Cells(1).Text)
    '                Dim name As String = row.Cells(2).Text
    '                Dim country As String = row.Cells(3).Text
    '                dt.Rows.Add(id, name, country)
    '            End If
    '        Next
    '        If dt.Rows.Count > 0 Then
    '            Dim consString As String = ConfigurationManager.ConnectionStrings("constr").ConnectionString
    '            Using con As New SqlConnection(consString)
    '                Using sqlBulkCopy As New SqlBulkCopy(con)
    '                    'Set the database table name
    '                    sqlBulkCopy.DestinationTableName = "dbo.Customers"

    '                    '[OPTIONAL]: Map the DataTable columns with that of the database table
    '                    sqlBulkCopy.ColumnMappings.Add("Id", "CustomerId")
    '                    sqlBulkCopy.ColumnMappings.Add("Name", "Name")
    '                    sqlBulkCopy.ColumnMappings.Add("Country", "Country")
    '                    con.Open()
    '                    sqlBulkCopy.WriteToServer(dt)
    '                    con.Close()
    '                End Using
    '            End Using
    '        End If
    '    End Sub

    Function NextFriday(ByRef Tdate) As String
        Tdate = Tdate.AddDays(1)
        Do While Tdate.DayOfWeek <> DayOfWeek.Friday
            Tdate = Tdate.AddDays(1)
        Loop
        Return Tdate
    End Function


End Module

