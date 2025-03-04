<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm3rdPartyExpl
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
        Me.cmdOK = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabInput = New System.Windows.Forms.TabPage
        Me.txtInput = New System.Windows.Forms.TextBox
        Me.TabOutPut = New System.Windows.Forms.TabPage
        Me.txtParts = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TabHelp = New System.Windows.Forms.TabPage
        Me.txtHelp = New System.Windows.Forms.TextBox
        Me.TabControl1.SuspendLayout()
        Me.TabInput.SuspendLayout()
        Me.TabOutPut.SuspendLayout()
        Me.TabHelp.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(37, 326)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(100, 29)
        Me.cmdOK.TabIndex = 0
        Me.cmdOK.Text = "Place Parts"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(143, 326)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(94, 29)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabInput)
        Me.TabControl1.Controls.Add(Me.TabOutPut)
        Me.TabControl1.Controls.Add(Me.TabHelp)
        Me.TabControl1.Location = New System.Drawing.Point(33, 17)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(546, 303)
        Me.TabControl1.TabIndex = 2
        '
        'TabInput
        '
        Me.TabInput.BackColor = System.Drawing.Color.Transparent
        Me.TabInput.Controls.Add(Me.txtInput)
        Me.TabInput.Location = New System.Drawing.Point(4, 22)
        Me.TabInput.Name = "TabInput"
        Me.TabInput.Padding = New System.Windows.Forms.Padding(3)
        Me.TabInput.Size = New System.Drawing.Size(538, 277)
        Me.TabInput.TabIndex = 0
        Me.TabInput.Text = "Input"
        Me.TabInput.UseVisualStyleBackColor = True
        '
        'txtInput
        '
        Me.txtInput.Location = New System.Drawing.Point(8, 9)
        Me.txtInput.Multiline = True
        Me.txtInput.Name = "txtInput"
        Me.txtInput.ReadOnly = True
        Me.txtInput.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtInput.Size = New System.Drawing.Size(524, 256)
        Me.txtInput.TabIndex = 0
        '
        'TabOutPut
        '
        Me.TabOutPut.BackColor = System.Drawing.Color.Transparent
        Me.TabOutPut.Controls.Add(Me.txtParts)
        Me.TabOutPut.Controls.Add(Me.Label2)
        Me.TabOutPut.Controls.Add(Me.Label1)
        Me.TabOutPut.Location = New System.Drawing.Point(4, 22)
        Me.TabOutPut.Name = "TabOutPut"
        Me.TabOutPut.Padding = New System.Windows.Forms.Padding(3)
        Me.TabOutPut.Size = New System.Drawing.Size(538, 277)
        Me.TabOutPut.TabIndex = 1
        Me.TabOutPut.Text = "Output"
        Me.TabOutPut.UseVisualStyleBackColor = True
        '
        'txtParts
        '
        Me.txtParts.Location = New System.Drawing.Point(23, 52)
        Me.txtParts.MaxLength = 1000000
        Me.txtParts.Multiline = True
        Me.txtParts.Name = "txtParts"
        Me.txtParts.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtParts.Size = New System.Drawing.Size(496, 219)
        Me.txtParts.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(19, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(191, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "HangerAttribute , Attributename , Value"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(19, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(303, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Name of Part , E , N , EL , RotE , RotN , RotEL , Len , Attr=Val"
        '
        'TabHelp
        '
        Me.TabHelp.Controls.Add(Me.txtHelp)
        Me.TabHelp.Location = New System.Drawing.Point(4, 22)
        Me.TabHelp.Name = "TabHelp"
        Me.TabHelp.Size = New System.Drawing.Size(538, 277)
        Me.TabHelp.TabIndex = 2
        Me.TabHelp.Text = "Help"
        Me.TabHelp.UseVisualStyleBackColor = True
        '
        'txtHelp
        '
        Me.txtHelp.Location = New System.Drawing.Point(8, 11)
        Me.txtHelp.Multiline = True
        Me.txtHelp.Name = "txtHelp"
        Me.txtHelp.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtHelp.Size = New System.Drawing.Size(513, 254)
        Me.txtHelp.TabIndex = 0
        '
        'frm3rdPartyExpl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(627, 376)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Name = "frm3rdPartyExpl"
        Me.Text = "Example 3rd Party Application"
        Me.TabControl1.ResumeLayout(False)
        Me.TabInput.ResumeLayout(False)
        Me.TabInput.PerformLayout()
        Me.TabOutPut.ResumeLayout(False)
        Me.TabOutPut.PerformLayout()
        Me.TabHelp.ResumeLayout(False)
        Me.TabHelp.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabInput As System.Windows.Forms.TabPage
    Friend WithEvents TabOutPut As System.Windows.Forms.TabPage
    Friend WithEvents TabHelp As System.Windows.Forms.TabPage
    Friend WithEvents txtParts As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtHelp As System.Windows.Forms.TextBox
    Friend WithEvents txtInput As System.Windows.Forms.TextBox
End Class
