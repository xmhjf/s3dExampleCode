<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm3rdPartyConfig
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
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtExchPath = New System.Windows.Forms.TextBox
        Me.txtCfgDisplay = New System.Windows.Forms.TextBox
        Me.cmdSPMSelFile = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdApply = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtCfgPath = New System.Windows.Forms.TextBox
        Me.txtLogPath = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkShowLog = New System.Windows.Forms.CheckBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtLogLevel = New System.Windows.Forms.ComboBox
        Me.CmdLogPath = New System.Windows.Forms.Button
        Me.cmdConfig = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 176)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "ExchangeFile"
        '
        'txtExchPath
        '
        Me.txtExchPath.Location = New System.Drawing.Point(103, 169)
        Me.txtExchPath.Name = "txtExchPath"
        Me.txtExchPath.Size = New System.Drawing.Size(436, 20)
        Me.txtExchPath.TabIndex = 1
        '
        'txtCfgDisplay
        '
        Me.txtCfgDisplay.Location = New System.Drawing.Point(103, 195)
        Me.txtCfgDisplay.Multiline = True
        Me.txtCfgDisplay.Name = "txtCfgDisplay"
        Me.txtCfgDisplay.Size = New System.Drawing.Size(436, 142)
        Me.txtCfgDisplay.TabIndex = 2
        '
        'cmdSPMSelFile
        '
        Me.cmdSPMSelFile.Location = New System.Drawing.Point(552, 169)
        Me.cmdSPMSelFile.Name = "cmdSPMSelFile"
        Me.cmdSPMSelFile.Size = New System.Drawing.Size(24, 20)
        Me.cmdSPMSelFile.TabIndex = 3
        Me.cmdSPMSelFile.Text = "..."
        Me.cmdSPMSelFile.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(103, 343)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(85, 24)
        Me.cmdOK.TabIndex = 4
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(194, 343)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(83, 24)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdApply
        '
        Me.cmdApply.Location = New System.Drawing.Point(283, 343)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(73, 24)
        Me.cmdApply.TabIndex = 6
        Me.cmdApply.Text = "Apply"
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "General Parameters"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(20, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Configuration File"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(39, 45)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Log File Path"
        '
        'txtCfgPath
        '
        Me.txtCfgPath.Location = New System.Drawing.Point(112, 16)
        Me.txtCfgPath.Name = "txtCfgPath"
        Me.txtCfgPath.Size = New System.Drawing.Size(418, 20)
        Me.txtCfgPath.TabIndex = 10
        '
        'txtLogPath
        '
        Me.txtLogPath.Location = New System.Drawing.Point(112, 45)
        Me.txtLogPath.Name = "txtLogPath"
        Me.txtLogPath.Size = New System.Drawing.Size(418, 20)
        Me.txtLogPath.TabIndex = 11
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkShowLog)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtLogLevel)
        Me.GroupBox1.Controls.Add(Me.CmdLogPath)
        Me.GroupBox1.Controls.Add(Me.cmdConfig)
        Me.GroupBox1.Controls.Add(Me.txtLogPath)
        Me.GroupBox1.Controls.Add(Me.txtCfgPath)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 9)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(573, 141)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "l"
        '
        'chkShowLog
        '
        Me.chkShowLog.AutoSize = True
        Me.chkShowLog.Location = New System.Drawing.Point(112, 118)
        Me.chkShowLog.Name = "chkShowLog"
        Me.chkShowLog.Size = New System.Drawing.Size(237, 17)
        Me.chkShowLog.TabIndex = 16
        Me.chkShowLog.Text = "Show Content of Logfile when program stops"
        Me.chkShowLog.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(20, 79)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(174, 13)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Log Level (0-no log,, 100-maximum)"
        '
        'txtLogLevel
        '
        Me.txtLogLevel.FormattingEnabled = True
        Me.txtLogLevel.Items.AddRange(New Object() {"0", "50", "100"})
        Me.txtLogLevel.Location = New System.Drawing.Point(216, 71)
        Me.txtLogLevel.Name = "txtLogLevel"
        Me.txtLogLevel.Size = New System.Drawing.Size(60, 21)
        Me.txtLogLevel.TabIndex = 14
        '
        'CmdLogPath
        '
        Me.CmdLogPath.Location = New System.Drawing.Point(536, 45)
        Me.CmdLogPath.Name = "CmdLogPath"
        Me.CmdLogPath.Size = New System.Drawing.Size(26, 22)
        Me.CmdLogPath.TabIndex = 13
        Me.CmdLogPath.Text = "..."
        Me.CmdLogPath.UseVisualStyleBackColor = True
        '
        'cmdConfig
        '
        Me.cmdConfig.Location = New System.Drawing.Point(536, 16)
        Me.cmdConfig.Name = "cmdConfig"
        Me.cmdConfig.Size = New System.Drawing.Size(26, 22)
        Me.cmdConfig.TabIndex = 12
        Me.cmdConfig.Text = "..."
        Me.cmdConfig.UseVisualStyleBackColor = True
        '
        'frm3rdPartyConfig
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(599, 379)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdApply)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdSPMSelFile)
        Me.Controls.Add(Me.txtCfgDisplay)
        Me.Controls.Add(Me.txtExchPath)
        Me.Controls.Add(Me.Label3)
        Me.Name = "frm3rdPartyConfig"
        Me.Text = "Configuration for Example 3rd Party Application"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtExchPath As System.Windows.Forms.TextBox
    Friend WithEvents txtCfgDisplay As System.Windows.Forms.TextBox
    Friend WithEvents cmdSPMSelFile As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdApply As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtCfgPath As System.Windows.Forms.TextBox
    Friend WithEvents txtLogPath As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CmdLogPath As System.Windows.Forms.Button
    Friend WithEvents cmdConfig As System.Windows.Forms.Button
    Friend WithEvents chkShowLog As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtLogLevel As System.Windows.Forms.ComboBox
End Class
