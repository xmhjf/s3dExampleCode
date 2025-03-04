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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   ProgID          :  SP3DMultiCableTransit.MCableTransit
'   File            :  CMaintenance.cls
'   Author          :  PK
'   Creation Date   :  Monday, 19 Nov 2007
'   Description     :  Multiple Cable Transit
'   Change History  :
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   19-NOV-2007      PK        TR-129855:Created the symbol.
'   02.JUN.2008      PK        CR-143345  Update MCT symbol to be able to specify offset for each port independently
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Maintenance:" 'Used for error messages
Private Const NEGLIGIBLE_THICHNESS = 0.0001 'Used when the value of Port Offset becomes zero.

Private Sub Class_Initialize()
''''

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput                 As Double
    Dim parWidth                As Double
    Dim parDepth                As Double
    Dim parLength               As Double
    Dim parPortOffset           As Double
    Dim parThickness            As Double
    Dim parPartitionsAlongWidth As Long
    Dim parPartitionsAlongDepth As Long
    Dim parVerticalPortOffset   As Double
    Dim parHorizontalPortOffset As Double
    Dim parPort2Offset          As Double
    
    Dim dTotalWidth             As Double
    Dim dPort1Offset            As Double
    Dim dPort2Offset            As Double
    
    Dim oStPoint            As AutoMath.DPosition
    Dim oEnPoint            As AutoMath.DPosition
        
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    iOutput = 0
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parLength = arrayOfInputs(4)
    parPortOffset = arrayOfInputs(5)
    parThickness = arrayOfInputs(6)
    parPartitionsAlongWidth = arrayOfInputs(7)
    parPartitionsAlongDepth = arrayOfInputs(8)
    parVerticalPortOffset = arrayOfInputs(9)
    parHorizontalPortOffset = arrayOfInputs(10)
    
    If UBound(arrayOfInputs) >= 11 Then
        parPort2Offset = arrayOfInputs(11)
        dPort1Offset = parPortOffset
        dPort2Offset = parPort2Offset
    Else
        dPort1Offset = parPortOffset
        dPort2Offset = parPortOffset
    End If
    
    If CmpDblEqual(dPort1Offset, 0) Then
        dPort1Offset = NEGLIGIBLE_THICHNESS
    End If
    
    If CmpDblEqual(dPort2Offset, 0) Then
        dPort2Offset = NEGLIGIBLE_THICHNESS
    End If
    
    dTotalWidth = (parPartitionsAlongWidth + 1) * parWidth
    'Create Maintenance Box1
    Dim oMaintBox1  As Object
    oStPoint.Set -parLength / 2, -dTotalWidth / 2, -parDepth / 2
    oEnPoint.Set -parLength / 2 - dPort1Offset, dTotalWidth / 2, parDepth / 2
    Set oMaintBox1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oMaintBox1
    Set oMaintBox1 = Nothing
    
    'Create Maintenance Box2
    Dim oMaintBox2  As Object
    oStPoint.Set parLength / 2, -dTotalWidth / 2, -parDepth / 2
    oEnPoint.Set parLength / 2 + dPort2Offset, dTotalWidth / 2, parDepth / 2
    Set oMaintBox2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oMaintBox2
    Set oMaintBox2 = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
