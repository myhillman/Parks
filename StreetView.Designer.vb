<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class StreetView
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.tbEncodedURL = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbDecodedURL = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnEncode = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnPaste = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'tbEncodedURL
        '
        Me.tbEncodedURL.Location = New System.Drawing.Point(100, 43)
        Me.tbEncodedURL.Name = "tbEncodedURL"
        Me.tbEncodedURL.Size = New System.Drawing.Size(889, 20)
        Me.tbEncodedURL.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 46)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Street View URL"
        '
        'tbDecodedURL
        '
        Me.tbDecodedURL.Location = New System.Drawing.Point(100, 81)
        Me.tbDecodedURL.Name = "tbDecodedURL"
        Me.tbDecodedURL.ReadOnly = True
        Me.tbDecodedURL.Size = New System.Drawing.Size(889, 20)
        Me.tbDecodedURL.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 84)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(75, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Encoded URL"
        '
        'btnEncode
        '
        Me.btnEncode.Location = New System.Drawing.Point(461, 128)
        Me.btnEncode.Name = "btnEncode"
        Me.btnEncode.Size = New System.Drawing.Size(72, 30)
        Me.btnEncode.TabIndex = 4
        Me.btnEncode.Text = "Encode"
        Me.btnEncode.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(97, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(332, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Will recode a dynamic Street View URL into a static Street View URL"
        '
        'btnPaste
        '
        Me.btnPaste.Location = New System.Drawing.Point(996, 39)
        Me.btnPaste.Name = "btnPaste"
        Me.btnPaste.Size = New System.Drawing.Size(62, 27)
        Me.btnPaste.TabIndex = 6
        Me.btnPaste.Text = "Paste"
        Me.btnPaste.UseVisualStyleBackColor = True
        '
        'StreetView
        '
        Me.AcceptButton = Me.btnEncode
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1070, 174)
        Me.Controls.Add(Me.btnPaste)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnEncode)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbDecodedURL)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbEncodedURL)
        Me.Name = "StreetView"
        Me.Text = "StreetView"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tbEncodedURL As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents tbDecodedURL As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents btnEncode As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents btnPaste As Button
End Class
