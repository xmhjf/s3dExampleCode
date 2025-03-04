VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   CInsulation.cls
'   Author:         RRK
'   Creation Date:  Tuesday, Aug 7 2007
'   Description:
'   This symbol is prepared for Shut off damper that caters to Round and Rectangular shapes as per CR-123916

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parDamperLength As Double
    Dim parActuatorPosition As Double
    Dim parCentertoSlot As Double
    Dim parSlotToLeverEnd As Double
    Dim parActuatorWidth As Double
    Dim parRodExtension As Double
    Dim parAngle As Double
    Dim parHVACShape As Double
    Dim parInsulationThickness As Double
    
    Dim CornerRadius As Double
    CornerRadius = 0

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parDamperLength = arrayOfInputs(4)
    parActuatorPosition = arrayOfInputs(5)
    parCentertoSlot = arrayOfInputs(6)
    parSlotToLeverEnd = arrayOfInputs(7)
    parActuatorWidth = arrayOfInputs(8)
    parRodExtension = arrayOfInputs(9)
    parAngle = arrayOfInputs(10)
    parHVACShape = arrayOfInputs(11)
    parInsulationThickness = arrayOfInputs(12)
    
    
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If

'Creating Insulation for Damper Body (Output 1)
    Dim objInsDamperBody As Object
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    If (parHVACShape = 4) Then 'Round
    
        oStPoint.Set -parDamperLength / 2, 0, 0
        oEnPoint.Set parDamperLength / 2, 0, 0
        
        Set objInsDamperBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            parWidth + (2 * parInsulationThickness), True)
        
    ElseIf (parHVACShape = Rectangular) Then
    
        oStPoint.Set -parDamperLength / 2, _
                    (parDepth / 2) + parInsulationThickness, _
                    (parWidth / 2) + parInsulationThickness
        
        oEnPoint.Set parDamperLength / 2, -(parDepth / 2) - parInsulationThickness, _
                    -(parWidth / 2) - parInsulationThickness
                    
        Set objInsDamperBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
    End If
    
'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsDamperBody
    Set objInsDamperBody = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError2 MODULE, METHOD
End Sub


