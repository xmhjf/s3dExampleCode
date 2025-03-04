'*******************************************************************
'Copyright (C) 2009, Intergraph Corporation. All rights reserved.
'
'Abstract:
'    frm3dPartyExpl.vb - example for a 3d party plugin
'
'Description:
'
'   This form will be shown, when "ComputePositions" is called.
'   The user can see, what Sp3d has transferred to the
'   3d party plugin (INput),
'   and the user may enter some lines (Output) to say
'   Sp3d, what support components to place.
'
'Notes:
'
'
'History
'GWE                Sep/01/2007      created
'Chandra Sekhar R   APR/01/2009      Converted to .net   
'******************************************************************
Option Explicit On
Public Class frm3rdPartyExpl

    '
    '   The return status
    '       0 - Place
    '      -1 - Cancel
    '
    Public lStatus As Long
    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        lStatus = -1
        Me.Hide()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        lStatus = 0
        Me.Hide()
    End Sub

    Private Sub frm3rdPartyExpl_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Call cmdCancel_Click(sender, e)
    End Sub

    Private Sub frm3dPartyExpl_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        txtHelp.Text = _
            "The Input tab displays the information prepared by S3D and passed to the 3rd party application." & vbCrLf & _
            vbCrLf & _
            "In the Output tab, the user can enter information defining which parts can be placed and where." & vbCrLf & _
            "You enter information about the parts, one per line.  The meaning of the parameters on the line are as follows:" & vbCrLf & _
            vbCrLf & _
            "Name of Part:  This one should be shown as PartNumber to match the column in the xls." & vbCrLf & _
            "East-Delta: This is the East-Direction distance (in meters) of the Component from the Support Position." & vbCrLf & _
            "North-Delta: This is the North-Direction distance (in meters) of the Component from the Support Position." & vbCrLf & _
            "Z-Height: This is the vertical distance (in meters) of the Component from the Support Position." & vbCrLf & _
            "East/West-Rotation: This is the Rotation of the Component about X-Axis in degrees" & vbCrLf & _
            "North/South-Rotation: This is the Rotation of the Component about Y-Axis in degrees" & vbCrLf & _
            "Z-Rotation: This is the Rotation of the Component about Z-Axis in degrees" & vbCrLf & _
            "Length: This is an Optional Parameter value which is length of the Component (in meters)." & vbCrLf & _
            "Parameter = Value: This is an Optional Parameter value which can be given to set on the component.ex..TOTAL_TRAVEL=0.02 " & vbCrLf & _
            vbCrLf & _
            "! at the beginning of any line: This is a comment line." & vbCrLf

        If (txtParts.Text.Length <> 0) Then
            cmdOK.Enabled = True
        Else
            cmdOK.Enabled = False
        End If

    End Sub

    'select all(CTRL+A) for textbox
    Private Sub txtParts_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtParts.KeyDown
        If ((e.Control) And (e.KeyCode = System.Windows.Forms.Keys.A)) Then
            txtParts.SelectAll()
            e.SuppressKeyPress = True
            e.Handled = True
        End If
    End Sub

    Private Sub txtParts_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtParts.TextChanged
        If (txtParts.Text.Length <> 0) Then
            cmdOK.Enabled = True
        Else
            cmdOK.Enabled = False
        End If
    End Sub
End Class