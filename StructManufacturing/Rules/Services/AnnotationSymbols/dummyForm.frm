VERSION 5.00
Begin VB.Form dummyForm 
   Caption         =   "f"
   ClientHeight    =   11205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   11205
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "dummyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, _
                            ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, _
                                                                            lpMetrics As TEXTMETRIC) As Long
                                                                            
Private Type Size
    cx As Long
    cy As Long
End Type

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Public Function GetFontMetricData(ByVal sFont As String, ByVal dFontSize As Double, ByVal oOrigPos As DPosition, _
                                     ByVal oVector As DVector, ByVal sJust As String, ByVal sText As String) As DPosition
    ' A graphical representation of the various
    ' Font details (by Mike Williams).
    Dim mymetrics As TEXTMETRIC, s1 As String, savey As Single
    Dim InternalLeading As Single
    Dim Ascent As Single
    Dim Descent As Single
    Dim ExternalLeading As Single
    Dim txtHeight As Single
    Dim TotalWidth As Double
    
    Me.BackColor = vbWhite
    Me.Width = Me.ScaleX(7, vbInches, vbTwips)
    Me.Height = Me.ScaleY(5, vbInches, vbTwips)
    Me.AutoRedraw = False 'TR-197970: Need to reset to prevent an error.
    
    Me.Line (0, 0)-(Me.ScaleWidth, Me.ScaleHeight), RGB(200, 200, 255), BF
    
    Me.ScaleMode = vbPixels
    Me.Font.Name = sFont
    If dFontSize = 0 Then
        dFontSize = 10 'Cannot have zero font size, will give an error!
    End If
    Me.Font.Size = dFontSize ' 72 points = 1 inch
    txtHeight = Me.TextHeight(sTEXT)
    TotalWidth = Me.TextWidth(sTEXT)
    'MsgBox "TotalHeight:" & txtHeight & "   & TotalWidth:" & TotalWidth
    
    GetTextMetrics Me.hdc, mymetrics
    
    InternalLeading = Me.ScaleY(mymetrics.tmInternalLeading, vbPixels, Me.ScaleMode)
    Ascent = Me.ScaleY(mymetrics.tmAscent, vbPixels, Me.ScaleMode)
    Descent = Me.ScaleY(mymetrics.tmDescent, vbPixels, Me.ScaleMode)
    
    ExternalLeading = Me.ScaleY(mymetrics.tmExternalLeading, vbPixels, Me.ScaleMode)
    
    Me.CurrentX = 0: Me.CurrentY = 0
    Me.FontTransparent = True
    'Me.Print "MjtplqdgfM"
    'Me.Print "<Njplk>-..BK.902-01-L.2-01"
    Me.FontTransparent = False
    'Me.Print "<Njplk>-..BK.902-01-L.2-01"
    savey = Me.CurrentY
    Me.Line (0, InternalLeading)-(Me.ScaleWidth, InternalLeading), vbYellow
    Me.Line (0, Ascent)-(Me.ScaleWidth, Ascent), vbGreen
    Me.Line (0, Ascent + Descent)-(Me.ScaleWidth, Ascent + Descent), vbBlue
    Me.Line (0, Ascent + Descent + ExternalLeading)-(Me.ScaleWidth, Ascent + Descent + ExternalLeading), vbRed
    
    Me.CurrentX = 0
    Me.CurrentY = savey
    'Get the height and width of our text
    Dim dTextSize As Size
    GetTextExtentPoint32 Me.hdc, "<NoBlk>-..BK.902-01-L.2-01", Len("<NoBlk>-..BK.902-01-L.2-01"), dTextSize
    
    Dim oNewPos As New DPosition
    oVector.length = 1
    
    '********************* CALCULATION OF POSITIONS ***********************'
    
'                   ul__________um___________ur                                  .
'                    |                       |                                  /|\
'                    |                       |                                   | Add Descent
'                    |                       |                                   |                      Subtract width or width/2
'    ---------------ml----------mm-----------mr----------------            <-----.                 .----->
'                    |                       |                       Add width or width/2          |
'                    |                       |                                                     |
'             ......bl..........bm...........br...                                                \|/
'                    |_______________________|                      (. means orig pos)             . Subtract Descent
'                   ll          lm           lr
    
    '**********************************************************************'
    oNewPos.Z = 0
    
    Select Case UCase(sJust)
        Case "LL"
            oNewPos.x = oOrigPos.x + Descent * -oVector.y
            oNewPos.y = oOrigPos.y + Descent * oVector.x
        Case "LM"
            oNewPos.x = oOrigPos.x + TotalWidth / 2 * -oVector.x
            oNewPos.y = oOrigPos.y + TotalWidth / 2 * -oVector.y
            
            oNewPos.x = oNewPos.x + Descent * -oVector.y
            oNewPos.y = oNewPos.y + Descent * oVector.x
        Case "LR"
            oNewPos.x = oOrigPos.x + TotalWidth * -oVector.x
            oNewPos.y = oOrigPos.y + TotalWidth * -oVector.y
            
            oNewPos.x = oNewPos.x + Descent * -oVector.y
            oNewPos.y = oNewPos.y + Descent * oVector.x
        Case "UL"
            oNewPos.x = oOrigPos.x - (txtHeight - Descent) * -oVector.y
            oNewPos.y = oOrigPos.y - (txtHeight - Descent) * oVector.x
        Case "UM"
            oNewPos.x = oOrigPos.x + TotalWidth / 2 * -oVector.x
            oNewPos.y = oOrigPos.y + TotalWidth / 2 * -oVector.y
            
            oNewPos.x = oNewPos.x - (txtHeight - Descent) * -oVector.y
            oNewPos.y = oNewPos.y - (txtHeight - Descent) * oVector.x
        Case "UR"
            oNewPos.x = oOrigPos.x + TotalWidth * -oVector.x
            oNewPos.y = oOrigPos.y + TotalWidth * -oVector.y
            
            oNewPos.x = oNewPos.x - (txtHeight - Descent) * -oVector.y
            oNewPos.y = oNewPos.y - (txtHeight - Descent) * oVector.x
        Case "ML"
            oNewPos.x = oOrigPos.x - (txtHeight / 2 - Descent) * -oVector.y
            oNewPos.y = oOrigPos.y - (txtHeight / 2 - Descent) * oVector.x
        Case "MM"
            oNewPos.x = oOrigPos.x + TotalWidth / 2 * -oVector.x
            oNewPos.y = oOrigPos.y + TotalWidth / 2 * -oVector.y
            
            oNewPos.x = oNewPos.x - (txtHeight / 2 - Descent) * -oVector.y
            oNewPos.y = oNewPos.y - (txtHeight / 2 - Descent) * oVector.x
        Case "MR"
            oNewPos.x = oOrigPos.x + TotalWidth * -oVector.x
            oNewPos.y = oOrigPos.y + TotalWidth * -oVector.y
            
            oNewPos.x = oNewPos.x - (txtHeight / 2 - Descent) * -oVector.y
            oNewPos.y = oNewPos.y - (txtHeight / 2 - Descent) * oVector.x
        Case "BL"
            Set oNewPos = oOrigPos
        Case "BM"
            oNewPos.x = oOrigPos.x + TotalWidth / 2 * -oVector.x
            oNewPos.y = oOrigPos.y + TotalWidth / 2 * -oVector.y
        Case "BR"
            oNewPos.x = oOrigPos.x + TotalWidth * -oVector.x
            oNewPos.y = oOrigPos.y + TotalWidth * -oVector.y
        Case Else
            Set oNewPos = oOrigPos
    End Select
    
    Set GetFontMetricData = oNewPos
    
    'MsgBox "dtextSize.cx:" & dTextsize.cx & "  dtextSize.cy:" & dTextsize.cy

End Function



