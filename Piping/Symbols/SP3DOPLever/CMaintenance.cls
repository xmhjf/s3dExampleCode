VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   ProgID:         SP3DOPLever.OPLever
'   Author:         RUK
'   Creation Date:  Wednesday, Sugust 27, 2008
'   Description: This symbol is graphical implementation of Lever operator
'   Source: E-141 section of the Design document)
'   This symbol implements following partdatabasis
'   Default
'   Lever, Type 1   (60)
'   Lever, inclined, Type 1 (61)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   27.08.2008      RUK     CR-148069  Provide a more realistic lever operator symbol (Source - E-141 section of the Design document)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:"    'Used for error messages
Private Const Lever_Type1 = 60
Private Const Lever_Inclined_Type1 = 61

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Long
    
    Dim parOffsetFrmValCen As Double
    Dim parLeverLength As Double
    Dim parLeverLength1 As Double
    Dim parLeverHeight As Double
    Dim parLeverHeight1 As Double
    Dim parLeverHeight2 As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parOffsetFrmValCen = arrayOfInputs(2)
    'parLeverLength = arrayOfInputs(3)
    'parLeverLength1 = arrayOfInputs(4)
    'parLeverHeight = arrayOfInputs(5)
    'parLeverHeight1 = arrayOfInputs(6)
    'parLeverHeight2 = arrayOfInputs(7)

    iOutput = 0

    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition

    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition

    Dim objMaintBody As Object
    
    Dim dInclAngle As Double
    Dim dHandleAngle As Double
    Dim dLeverWidth As Double
    Dim dLeverLength1 As Double
    Dim dLeverHeight1 As Double
    Dim dLeverHeight2 As Double
    Dim dLeverLength As Double
    Dim dPoints() As Double
    Dim dHandleLength As Double
    
    'Checking for the PartDataBasis
    Dim oOperator As IJValveOperator
    Dim lPartDataBasis As Long
    Set oOperator = oPartFclt
    lPartDataBasis = oOperator.DimensionalBasis
    Set oOperator = Nothing
    
    Select Case lPartDataBasis
    Case Is <= 1, Lever_Type1
        parOffsetFrmValCen = arrayOfInputs(2)
        parLeverLength = arrayOfInputs(3)
        parLeverLength1 = arrayOfInputs(4)
        parLeverHeight1 = arrayOfInputs(6)
        parLeverHeight2 = arrayOfInputs(7)
        
        dLeverWidth = 0.15 * parLeverLength
        If CmpDblLessThan(dLeverWidth, 0.05) Then dLeverWidth = 0.05
        If CmpDblGreaterthan(dLeverWidth, 0.1) Then dLeverWidth = 0.1
        
        dLeverLength = parLeverLength
        If CmpDblGreaterthan(parLeverLength1, LINEAR_TOLERANCE) Then
            dLeverLength1 = parLeverLength1
        Else
            dLeverLength1 = 0.4 * parLeverLength
        End If
        
        dLeverHeight1 = parLeverHeight1
        If CmpDblGreaterthan(parLeverHeight2, LINEAR_TOLERANCE) Then
            dLeverHeight2 = parLeverHeight2
        Else
            dLeverHeight2 = 1.2 * dLeverHeight1
        End If
        dHandleAngle = Atn((dLeverHeight2 + 0.1 * dLeverHeight1) / _
                        (0.8 * (parLeverLength - dLeverLength1)))
        
        'Create the Maintanance Body
        oStPoint.Set 0, parOffsetFrmValCen - dLeverHeight2, 0
        oEnPoint.Set 0, parOffsetFrmValCen + dLeverHeight1, 0
        Set objMaintBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 2 * parLeverLength, True)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMaintBody
        Set objMaintBody = Nothing
        
    Case Lever_Inclined_Type1
        parOffsetFrmValCen = arrayOfInputs(2)
        parLeverLength = arrayOfInputs(3)
        parLeverHeight = arrayOfInputs(5)
        
        dLeverWidth = 0.15 * parLeverLength
        If CmpDblLessThan(dLeverWidth, 0.05) Then dLeverWidth = 0.05
        If CmpDblGreaterthan(dLeverWidth, 0.1) Then dLeverWidth = 0.1
    
        dLeverLength1 = 0.4 * parLeverLength
        dInclAngle = Atn(0.8 * parLeverHeight / parLeverLength)
        
        'Compute the LeverLength
        dLeverLength = parLeverLength / Cos(dInclAngle)
        dHandleAngle = Atn((0.2 * parLeverHeight + dLeverLength1 * Tan(dInclAngle)) / _
                                    (parLeverLength - dLeverLength1))
        
        'Create the Maintanance Body
        oStPoint.Set 0, parOffsetFrmValCen, 0
        oEnPoint.Set 0, oStPoint.y + parLeverHeight + 0.05 * parLeverHeight, 0
        Set objMaintBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                        2 * (parLeverLength + 0.05 * parLeverHeight * Tan(dInclAngle)), True)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMaintBody
        Set objMaintBody = Nothing
    
    Case Else
        GoTo ErrorLabel
    End Select
        
    'Remove the References
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set objMaintBody = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
