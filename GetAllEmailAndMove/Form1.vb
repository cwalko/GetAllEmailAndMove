

' NEED: Add a log of each attempt to run this program

'You can use the following method to get all the mail from the Inbox and move it to the Destination folder:
Imports Microsoft.Office.Interop
Imports System.Drawing.Printing ' used to print List Box
Public Class Form1
    Dim dt As New DataTable
    Dim olApp As Outlook.Application = New Outlook.Application()
    Dim ns As Outlook.NameSpace = olApp.GetNamespace("MAPI")
    Dim olSourceFolder As Outlook.MAPIFolder = Nothing
    Dim mi As Outlook.MailItem
    Dim olDestinationFolder As Outlook.MAPIFolder = Nothing
    Dim ErrorNumber As Integer = Nothing
    Dim myYear As Integer
    Dim myMonth As Integer
    Dim myDate As String
    Dim fileDate As String
    ReadOnly separators As Char() = New Char() {Chr(13), Chr(10)}
    Public WithEvents docToPrint As New PrintDocument
    Public Sub Document_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles docToPrint.PrintPage
        Dim printFont As New Font("Arial", 10, System.Drawing.FontStyle.Regular)
        Dim YPosition As Integer = 20
        'MsgBox("ListBox Items Count = " & ListBox1.Items.Count) ' just add for debugging
        For Each eItem As String In ListBox1.Items 'there are 353 items in listbox1 - only 52 are printing
            e.Graphics.DrawString(eItem, printFont, System.Drawing.Brushes.Black, 25, YPosition)
            YPosition += 20
        Next

    End Sub




    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        myYear = Year(Now)  ' Get the year
        myMonth = Month(Now)  ' Get the month
        myDate = Format(Date.Now(), "dd")   ' "ddMMMyyyy"  Get the Date
        fileDate = myYear.ToString() & "-" & myMonth.ToString & "-" & myDate  ' build the Destination "file name" (date)

        olSourceFolder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) ' where new eCommerse emails are received - the inbox of "seafood@JaekleGroup.com"


    End Sub
    Sub CreateFolder()
        olDestinationFolder = ns.Folders("seafood@JaekleGroup.com").Folders("done").Folders.Add(fileDate, Outlook.OlDefaultFolders.olFolderInbox)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            olDestinationFolder = ns.Folders("seafood@JaekleGroup.com").Folders("done").Folders.Add(fileDate, Outlook.OlDefaultFolders.olFolderInbox)

        Catch ex As Exception
            'MsgBox(ex.Message)
            olDestinationFolder = ns.Folders("seafood@JaekleGroup.com").Folders("done").Folders(fileDate)

        End Try

        'olSourceFolder = ns.Folders(8).Folders(9).Folders(16) - this is  the INDEX of each folder
        'olDestinationFolder = ns.Folders("seafood@JaekleGroup.com").Folders("done").Folders("2019-06-16") - this is example of hard coded done folder

        ' Add a log of each attempt to run this program

        Me.Label1.Text = "Inbox: Total " & olSourceFolder.Items.Count.ToString() & " Mails"
        Me.Label1.Refresh()
        dt.Columns.Add("1", GetType(String))
        dt.Columns.Add("2", GetType(String))
        dt.Columns.Add("3", GetType(DateTime))
        dt.Columns.Add("4", GetType(String))
        dt.Columns.Add("5", GetType(String))
        Dim dr As DataRow = Nothing
        Dim isread As String = "Unread"
        Dim i As Integer
        Dim eCommerseSeperator As String = "#"
        Dim subject As String
        Dim eCommerseOrderNumber As String
        'MsgBox("Count = " & olSourceFolder.Items.Count)
        ' For Each item As Object In olSourceFolder.Items
        For i = olSourceFolder.Items.Count To 1 Step -1
            mi = olSourceFolder.Items(i)
            mi = TryCast(mi, Outlook.MailItem)
            If mi.UnRead = False Then
                isread = "Read"
            End If

            'Parse out the eCommerse Cart Order Number
            subject = mi.Subject
            Dim substring As String
            eCommerseSeperator = "#"
            Dim dIndex = subject.IndexOf("#")
            If (dIndex > -1) Then
            End If
            substring = subject.Substring(Str(dIndex) + 1)
            eCommerseOrderNumber = substring

            'Fill Data Table dt
            dr = dt.NewRow()
            dr("1") = i.ToString()
            dr("2") = eCommerseOrderNumber
            dr("3") = mi.CreationTime
            dr("4") = isread
            dr("5") = mi.Body

            Dim emailBody As String
            emailBody = mi.Body

            ' MsgBox("body legth = " & emailBody.Length)
            Dim emailBodyLines As String() = emailBody.Split(separators, StringSplitOptions.RemoveEmptyEntries) ' an array to hold each 'line' of the email

            ' MsgBox("# in array = " & emailBodyLines.Length)
            Dim a As Integer = -1
            For Each s As String In emailBodyLines
                a += 1
                ListBox1.Items.Add(a & ") " & (s))
            Next
            'ParseEmailLine(emailBodyLines)



            dt.Rows.Add(dr)
            'Me.Label2.Text = "Reading:" & i.ToString()
            'Me.Label2.Refresh()
            'MsgBox("Pausing")
            ' System.Threading.Thread.Sleep(1000)
            mi.Move(olDestinationFolder)
        Next

        Me.DataGridView1.DataSource = dt
        olApp = Nothing
        ns = Nothing
        olSourceFolder = Nothing
        mi = Nothing
        olDestinationFolder = Nothing

        MsgBox("Ready to close")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' MsgBox("ListBox Items = " & ListBox1.Items.Count) 'there are 353 items in listbox1 - only 52 are printing
        Dim PrintDialog1 As New PrintDialog
        Dim result As DialogResult = PrintDialog1.ShowDialog()
        If result = DialogResult.OK Then docToPrint.Print()
    End Sub
End Class
