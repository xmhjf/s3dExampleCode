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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.

'   ProgID          :  SP3DCableRiserTray.RiserTray
'   File            :  CMaintenance.cls
'   Author          :  PK
'   Creation Date   :  Friday, 25 April 2008
'   Description     :  Cable Tray Riser Box
'   Change History  :
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   25-APR-2008      PK        CR-136329:Created the symbol.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Maintenance:" 'Used for error messages
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
      
    Dim oStPoint            As AutoMath.DPosition
    Dim oEnPoint            As AutoMath.DPosition
        
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
    
    Dim dPortOffset As Double
    If CmpDblEqual(parPortOffset, 0) Then
        dPortOffset = NEGLIGIBLETHICKNESS
    Else
        dPortOffset = parPortOffset
    End If
    
    'Create Maintenance Box
    Dim oMaintBox  As Object
    If lCableTrayPartDataBasis = 138 Then
        oStPoint.Set -parThickness, parWidth / 2, parDepth / 2
        oEnPoint.Set -parThickness - dPortOffset, -parWidth / 2, -parDepth / 2
    Else
        If lCableTrayPartDataBasis = 139 Then
              oStPoint.Set -parThickness - parHeight / 2, parWidth / 2, parDepth / 2
              oEnPoint.Set -parThickness - dPortOffset - parHeight / 2, -parWidth / 2, _
                                                                    -parDepth / 2
        End If
    End If
    
    Set oMaintBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oMaintBox
    Set oMaintBox = Nothing
 
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


