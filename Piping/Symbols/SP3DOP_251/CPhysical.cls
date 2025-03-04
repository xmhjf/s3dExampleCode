VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   ProgID:         SP3DOP_251.COP_251
'   Author:         MS
'   Creation Date:  1 October 2003
'   Description:
'   This symbol is prepared based on SP3DOP_3.
'   Symbol Model No. is: OP_251 Page No. D-131 PDS Piping Component Data Reference Guide.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
  
    Dim iOutput     As Double
    
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim StemMinimumDiameter As Double
    Dim ActuatorWheelMinimumThickness As Double
    Dim StemExtension As Double
    
    Dim CenterPos       As New AutoMath.DPosition
    Dim CylBotOrigin    As New AutoMath.DPosition
    Dim CylTopOrigin    As New AutoMath.DPosition
    Dim TorOrigin       As New AutoMath.DPosition
    Dim TorAxis         As New AutoMath.DVector
    Dim TorPriRad       As Double
    Dim TorSecRad       As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)

'   Initialize
    iOutput = 0
    
    StemMinimumDiameter = parOperatorDiameter * 0.1
'   Check that the stem diameter is within 15mm and 50mm
   If StemMinimumDiameter < 0.015 Then StemMinimumDiameter = 0.015
   If StemMinimumDiameter > 0.05 Then StemMinimumDiameter = 0.05
   
   ActuatorWheelMinimumThickness = StemMinimumDiameter
   StemExtension = parOperatorHeight * 0.1

    CenterPos.Set 0, 0, 0
    
    CylBotOrigin.Set CenterPos.x, CenterPos.y, CenterPos.z
    CylTopOrigin.Set CylBotOrigin.x, CylBotOrigin.y + parOperatorHeight, CylBotOrigin.z
    
    TorOrigin.Set CylTopOrigin.x, CylTopOrigin.y - StemExtension - TorSecRad, CylTopOrigin.z

' Insert your code for output 1(Stem)
'   Place Cylinder
    Dim ObjStem As IngrGeom3D.Projection3d
    Set ObjStem = PlaceCylinder(m_OutputColl, CylBotOrigin, CylTopOrigin, StemMinimumDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    Set CenterPos = Nothing
    Set CylBotOrigin = Nothing
    Set CylTopOrigin = Nothing

' Insert your code for output 2(Handwheel)
'   Place Torus
    TorSecRad = ActuatorWheelMinimumThickness / 2
    TorPriRad = parOperatorDiameter / 2 - TorSecRad
    TorAxis.Set 0, 1, 0
    Dim ObjHandwheel    As IngrGeom3D.Torus3d
    Set ObjHandwheel = PlaceTorus(m_OutputColl, TorOrigin, TorAxis, TorPriRad, TorSecRad)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheel
    Set ObjHandwheel = Nothing
    Set TorOrigin = Nothing
    Set TorAxis = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

