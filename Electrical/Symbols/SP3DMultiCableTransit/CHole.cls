VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   ProgID          :  SP3DMultiCableTransit.MCableTransit
'   File            :  CHole.cls
'   Author          :  PK
'   Creation Date   :  Monday, 19 Nov 2007
'   Description     :  Multiple Cable Transit
'   Change History  :
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   19-NOV-2007      PK        TR-129855:Created the symbol.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Hole:" 'Used for error messages

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
    Dim oStPoint                As AutoMath.DPosition
    Dim oEnPoint                As AutoMath.DPosition
    Dim dTotalWidth             As Double
    
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
            
    dTotalWidth = (parPartitionsAlongWidth + 1) * parWidth
    'Insert your code for output 1(Cable Transit Hole)
    Dim oBox As Object
    oStPoint.Set -parLength / 2, dTotalWidth / 2, -parDepth / 2
    oEnPoint.Set parLength / 2, -dTotalWidth / 2, parDepth / 2
    Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
    Set oBox = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
  
End Sub
