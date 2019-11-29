Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO
Public Class Form1
    Public MyConnectionString As String = ""
    Dim SQLDB As String = "eCom_Email"
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Call EnterClick(sender, e)
        Dim FindValue As Integer
        RichTextBox1.Clear()
        txtboxDate.Text = FormatDateTime(Now, DateFormat.ShortDate)
        If txtboxName.Text <> "" Then
            FindValue = 1
        ElseIf txtboxEmail.Text <> "" Then
            FindValue = 2
        ElseIf txtboxOrderNo.Text <> "" Then
            FindValue = 3
        End If
        If FindValue <> 0 Then
            Open_SQL()
        End If
        If FindValue = 0 Then
            MsgBox("One of the three find fileds Must be filled in - all are empty")
        End If
    End Sub
    Public Sub Open_SQL()

        ' MsgBox(SQLDB & " - " & PayTransID  & " - " & CreationTime & " - " & SentOn & " - " & ReceivedTime & " - " & SenderName & " - " & SenderEmailAddress)
        Dim MyConnection As SqlConnection
        'Dim MyConnectionString As String
        Dim cmd As New SqlCommand
        Dim MyConnectionString As String = ""
        MyConnectionString = ConfigurationManager.ConnectionStrings("eCom_Email").ConnectionString   ' from App.config

        'MyConnectionString = "Data Source = .\SQLEXPRESS; Database = " & SQLDB & "; Integrated Security=SSPI"
        'MyConnectionString = "Data Source = DESKTOP-11TM25R\SQLEXPRESS; DataBase =" & SQLDB & "; Integrated Security=SSPI"   '; User Instance=True"
        'MsgBox("MyConnection String = " + MyConnectionString)
        MyConnection = New SqlConnection(MyConnectionString)
        Try
            MyConnection.Open()
            '  MsgBox("Con String = " + MyConnection.State.ToString)
        Catch Err As Exception
            MsgBox("Database Cannot Be Opened (line 40) - Call for Tech Support")
            MsgBox("Error = " + Err.ToString)
            MyConnection.Close()
            MyConnection.Dispose()
            Return
            ' System.Windows.Forms.Application.Exit()
        End Try
        Try
            If txtboxName.Text <> "" Then
                Dim queryString = "SELECT PayTransID, eMailBody FROM dbo.eComEmailBody WHERE eMailBody LIKE '%" + txtboxName.Text + "%'"
                SQLQueryExe(queryString, MyConnectionString)
            ElseIf txtboxEmail.Text <> "" Then
                Dim queryString = "SELECT PayTransID, eMailBody FROM dbo.eComEmailBody WHERE SenderEmailAddress LIKE '%" + txtboxEmail.Text + "%'"
                SQLQueryExe(queryString, MyConnectionString)
            ElseIf txtboxOrderNo.Text <> "" Then
                Dim queryString = "SELECT PayTransID, eMailBody FROM dbo.eComEmailBody WHERE PayTransID LIKE '%" + txtboxOrderNo.Text + "%'"
                SQLQueryExe(queryString, MyConnectionString)
            End If
        Catch ex As Exception
            MsgBox("Unable to execute the query, error to follow.")
            MsgBox("Error---" + ex.ToString)
        End Try
    End Sub
    Public Sub SQLQueryExe(queryString As String, MyConnectionString As String)
        queryString = queryString + " Order By CreationTime Desc"
        Using connection As New SqlConnection(MyConnectionString)
            Dim command = New SqlCommand(queryString, connection)
            Try
                connection.Open()
            Catch ex As Exception
                MsgBox("Database Cannot Be Opened while trying to run Query  - Call for Tech Support")
                connection.Close()
                connection.Dispose()
                Return
            End Try

            Using reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then
                    While reader.Read()
                        RichTextBox1.Text = RichTextBox1.Text + (String.Format("{0}, {1}", reader(0), reader(1))) + vbCrLf
                        RichTextBox1.Text = RichTextBox1.Text + "==================================================" + vbCrLf
                    End While
                Else
                    MsgBox("The data base has no record matching your search input.  Try something less specific.  ie: If you entered a full phone number in ANY INFO then try entering just the last four diget of that number.")
                End If

                Return
            End Using
        End Using


    End Sub
    Private Sub txtboxName_GotFocus(sender As Object, e As EventArgs) Handles txtboxName.GotFocus
        ' Dim Name As String = txtboxName.Text
        txtboxEmail.Clear()
        'Dim eMail As String = "'%'"
        txtboxOrderNo.Clear()
    End Sub
    Private Sub txtboxName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtboxName.KeyPress
        If e.KeyChar = Chr(13) Then 'Chr(13) is the Enter Key
            'Runs the Button1_Click Event
            Button1_Click(Me, EventArgs.Empty)
        End If
    End Sub
    Private Sub txtboxEmail_GotFocus(sender As Object, e As EventArgs) Handles txtboxEmail.GotFocus
        txtboxName.Clear()
        txtboxOrderNo.Clear()
    End Sub
    Private Sub txtboxEmail_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtboxEmail.KeyPress
        If e.KeyChar = Chr(13) Then 'Chr(13) is the Enter Key
            'Runs the Button1_Click Event
            Button1_Click(Me, EventArgs.Empty)
        End If
    End Sub

    Private Sub txtboxOrderNo_GotFocus(sender As Object, e As EventArgs) Handles txtboxOrderNo.GotFocus
        txtboxName.Clear()
        txtboxEmail.Clear()
    End Sub
    Public Sub ConnectionStringSettings()
        Dim MyConnectionString As String = ConfigurationManager.AppSettings("eCom_Email")
    End Sub
    Sub EnterPressed(e)
        If e.KeyChar = Chr(13) Then 'Chr(13) is the Enter Key
            'Runs the Button1_Click Event
            Button1_Click(Me, EventArgs.Empty)
        End If

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtboxDate.Text = FormatDateTime(Now, DateFormat.ShortDate)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
        System.Windows.Forms.Application.Exit()
        End ' End the program
    End Sub


End Class
