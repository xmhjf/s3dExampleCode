'*******************************************************************
'Copyright (C) 2009, Intergraph Corporation. All rights reserved.
'
'Abstract:
'    frm3dPartyConfig.vb - example for a 3d party plugin
'
'Description:
'
'   This form will be shown, Configuraion of the 3party Program is called
'
'Notes:
'
'
'History
'GWE                Sep/01/2007      created
'Chandra Sekhar R   APR/01/2009      Converted to .net   
'******************************************************************
Option Explicit On

Public Class frm3rdPartyConfig
    Private lResult As Long
    Public Function ShowModal()
        lResult = -1
        Me.ShowDialog()
        ShowModal = lResult
    End Function

    Private Sub cmdApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdApply.Click
        ' If values are change, store in Session file

        ' If values are change, store in Session file

        If (Me.txtCfgPath.Text <> gsIniPath Or _
           Me.txtExchPath.Text <> gsExchangeFilePath Or _
           Me.txtCfgDisplay.Text <> gs3dPartyParts Or _
           Me.txtLogPath.Text <> gsLogFilePath Or _
           Me.txtLogLevel.Text <> CStr(gLogLevel) Or _
           (Me.chkShowLog.Checked = True) <> gbShowLogfile) Then
            gsIniPath = Me.txtCfgPath.Text
            gsExchangeFilePath = Me.txtExchPath.Text
            gs3dPartyParts = Me.txtCfgDisplay.Text
            gsLogFilePath = Me.txtLogPath.Text
            gLogLevel = CLng(Me.txtLogLevel.Text)
            Me.txtLogLevel.Text = gLogLevel
            gbShowLogfile = Me.chkShowLog.Checked = 1
            WriteToPreference()
        End If
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        lResult = 0
        cmdApply_Click(sender, e)
        Me.Dispose()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        lResult = -1
        Me.Dispose()
    End Sub

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub cmdSPMSelFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSPMSelFile.Click
        On Error Resume Next

        OpenFileDialog1.CheckPathExists = True
        OpenFileDialog1.CheckFileExists = True
        OpenFileDialog1.AddExtension = True
        OpenFileDialog1.Title = "Select 3d Party Data Exchange File"
        OpenFileDialog1.Filter = "Text Files|*.txt;*.cfg;*.ini|All Files|*.*"
        OpenFileDialog1.InitialDirectory = "."
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.DefaultExt = ""
        Dim oResult As System.Windows.Forms.DialogResult
        oResult = OpenFileDialog1.ShowDialog()

        If (oResult = Windows.Forms.DialogResult.Cancel) Then
            ' cancel selected
            Exit Sub
        End If

        If (OpenFileDialog1.FileName = "") Then
            Exit Sub
        End If

        Me.txtExchPath.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub frm3rdPartyConfig_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.txtCfgPath.Text = gsIniPath
        Me.txtExchPath.Text = gsExchangeFilePath
        Me.txtCfgDisplay.Text = gs3dPartyParts
        Me.txtLogPath.Text = gsLogFilePath
        Me.txtLogLevel.Text = CStr(gLogLevel)
        Me.chkShowLog.Checked = IIf(gbShowLogfile, 1, 0)

    End Sub

    Private Sub cmdPartsSelFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        OpenFileDialog1.CheckPathExists = True
        OpenFileDialog1.CheckFileExists = True
        OpenFileDialog1.AddExtension = True
        OpenFileDialog1.Title = "Select 3d Party Project File"
        OpenFileDialog1.Filter = "Project File|*.prj|All Files|*.*"
        OpenFileDialog1.InitialDirectory = "."
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.DefaultExt = ""
        Dim oResult As System.Windows.Forms.DialogResult
        oResult = OpenFileDialog1.ShowDialog()

        If (oResult = Windows.Forms.DialogResult.Cancel) Then
            ' cancel selected
            Exit Sub
        End If
        If (OpenFileDialog1.FileName = "") Then
            Exit Sub
        End If

        Me.txtCfgDisplay.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub txtCfgPath_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtExchPath.TextChanged
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdConfig.Click

        OpenFileDialog1.CheckPathExists = True
        OpenFileDialog1.CheckFileExists = True
        OpenFileDialog1.AddExtension = True

        OpenFileDialog1.Title = "Select HangerProg-Configuration File"
        OpenFileDialog1.Filter = "Configuration Files|*.ini|All Files|*.*"
        OpenFileDialog1.InitialDirectory = "."
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.DefaultExt = ""
        Dim oResult As System.Windows.Forms.DialogResult
        oResult = OpenFileDialog1.ShowDialog()

        If (oResult = Windows.Forms.DialogResult.Cancel) Then
            ' cancel selected
            Exit Sub
        End If

        If (OpenFileDialog1.FileName = "") Then
            Exit Sub
        End If

        Me.txtCfgPath.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdLogPath.Click

        OpenFileDialog1.CheckFileExists = False
        OpenFileDialog1.AddExtension = True
        OpenFileDialog1.Title = "Select 3d Party LogFile"
        OpenFileDialog1.Filter = "Log Files|*.log|Text Files|*.txt|All Files|*.*"
        OpenFileDialog1.InitialDirectory = "."
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.DefaultExt = ""
        Dim oResult As System.Windows.Forms.DialogResult
        oResult = OpenFileDialog1.ShowDialog()

        If (oResult = Windows.Forms.DialogResult.Cancel) Then
            ' cancel selected
            Exit Sub
        End If

        If (OpenFileDialog1.FileName = "") Then
            Exit Sub
        End If

        Me.txtLogPath.Text = OpenFileDialog1.FileName
    End Sub

End Class