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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.

'   ProgID          :  SP3DCableRiserTray.RiserTray
'   File            :  CHole.cls
'   Author          :  PK
'   Creation Date   :  Friday, 25 April 2008
'   Description     :  Cable Tray Riser Box
'   Change History  :
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   25-APR-2008      PK        CR-136329:Created the symbol.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Hole:" 'Used for error messages
Private Const NEGLIGIBLETHICKNESS = 0.0001

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
    Dim parHeight               As Double
    Dim parPortOffset           As Double
    Dim parThickness            As Double
 
    Dim oStPoint                As AutoMath.DPosition
    Dim oEnPoint                As AutoMath.DPosition
        
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    iOutput = 0
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parHeight = arrayOfInputs(4)
    parPortOffset = arrayOfInputs(5)
    parThickness = arrayOfInputs(6)
    
    Dim lCableTrayPartDataBasis As Long
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
    
    Dim dThickness As Double
    If CmpDblEqual(parPortOffset, 0) Then
        dThickness = NEGLIGIBLETHICKNESS
    Else
        dThickness = parThickness
    End If
                
    'Insert your code for output 1(Cable Transit Hole)
    Dim oBox As Object
    If lCableTrayPartDataBasis = 138 Then
        oStPoint.Set 0, parWidth / 2, parDepth / 2
        oEnPoint.Set -dThickness, -parWidth / 2, -parDepth / 2
    Else
          If lCableTrayPartDataBasis = 139 Then
             oStPoint.Set -parHeight / 2, parWidth / 2, parDepth / 2
             oEnPoint.Set -dThickness - parHeight / 2, -parWidth / 2, -parDepth / 2
          End If
    End If
    
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

