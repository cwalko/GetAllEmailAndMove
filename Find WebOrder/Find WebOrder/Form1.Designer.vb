<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtboxName = New System.Windows.Forms.TextBox()
        Me.txtboxEmail = New System.Windows.Forms.TextBox()
        Me.txtboxOrderNo = New System.Windows.Forms.TextBox()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtboxDate = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(5, 121)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 29)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Any info"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 161)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 29)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "eMail"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(4, 208)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(95, 29)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Order #"
        '
        'txtboxName
        '
        Me.txtboxName.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtboxName.Location = New System.Drawing.Point(105, 118)
        Me.txtboxName.Name = "txtboxName"
        Me.txtboxName.Size = New System.Drawing.Size(256, 34)
        Me.txtboxName.TabIndex = 1
        '
        'txtboxEmail
        '
        Me.txtboxEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtboxEmail.Location = New System.Drawing.Point(105, 158)
        Me.txtboxEmail.Name = "txtboxEmail"
        Me.txtboxEmail.Size = New System.Drawing.Size(256, 34)
        Me.txtboxEmail.TabIndex = 2
        '
        'txtboxOrderNo
        '
        Me.txtboxOrderNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtboxOrderNo.Location = New System.Drawing.Point(103, 207)
        Me.txtboxOrderNo.Name = "txtboxOrderNo"
        Me.txtboxOrderNo.Size = New System.Drawing.Size(256, 34)
        Me.txtboxOrderNo.TabIndex = 3
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(372, 9)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(722, 513)
        Me.RichTextBox1.TabIndex = 99
        Me.RichTextBox1.TabStop = False
        Me.RichTextBox1.Text = ""
        '
        'Label4
        '
        Me.Label4.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Label4.Location = New System.Drawing.Point(18, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(343, 86)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = resources.GetString("Label4.Text")
        '
        'txtboxDate
        '
        Me.txtboxDate.Location = New System.Drawing.Point(143, 446)
        Me.txtboxDate.Name = "txtboxDate"
        Me.txtboxDate.Size = New System.Drawing.Size(100, 22)
        Me.txtboxDate.TabIndex = 9
        Me.txtboxDate.TabStop = False
        Me.txtboxDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(56, 446)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(38, 17)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Date"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(31, 329)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(315, 56)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Please note:  The search results will bring back multiple orders if the customer " &
    "has ordered with us in previous weeks of this year's dinners."
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(154, 388)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 12
        Me.Button2.Text = "Exit"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(145, 247)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(98, 39)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Search"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1105, 529)
        Me.Controls.Add(Me.txtboxName)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtboxEmail)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.txtboxOrderNo)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtboxDate)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "Find Web Order"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents txtboxName As TextBox
    Friend WithEvents txtboxEmail As TextBox
    Friend WithEvents txtboxOrderNo As TextBox
    Friend WithEvents RichTextBox1 As RichTextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents txtboxDate As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
End Class
