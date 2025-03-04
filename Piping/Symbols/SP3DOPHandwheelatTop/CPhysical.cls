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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Friday, August 17 2007
'   Description     :  Operator With Handwheel at Top
'   Reference       :  Hydrant Valve Angle Type.doc
'
'   Change History:
'   dd.mmm.yyyy     who      change description
'   -----------    -----     ------------------
'   24.Aug.2007     PK      CR 125185, Created the symbol.
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
    
    Dim parOperatorHeight     As Double
    Dim parOperatorDiameter   As Double
    Dim parValveFacetoCenter  As Double
    
    Dim StemMinimumDiameter           As Double
    Dim ActuatorWheelMinimumThickness As Double
    Dim TorPriRad                     As Double
    Dim TorSecRad                     As Double
    Dim iOutput                       As Integer
    
    Dim oOriginPos      As New Automath.DPosition
    Dim CylBotOrigin    As New Automath.DPosition
    Dim CylTopOrigin    As New Automath.DPosition
    Dim TorOrigin       As New Automath.DPosition
    Dim TorAxis         As New Automath.DVector

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)        'H1
    parOperatorDiameter = arrayOfInputs(3)      'd
    parValveFacetoCenter = arrayOfInputs(4)     'H2
    
    iOutput = 0

    StemMinimumDiameter = parOperatorDiameter * 0.1
    
    'Check that the stem diameter is within 15mm and 50mm
    If CmpDblLessThan(StemMinimumDiameter, 0.015) Then StemMinimumDiameter = 0.015
    If CmpDblGreaterthan(StemMinimumDiameter, 0.05) Then StemMinimumDiameter = 0.05
   
    ActuatorWheelMinimumThickness = StemMinimumDiameter
   
    'Prepare Torus Radii
    TorSecRad = ActuatorWheelMinimumThickness / 2
    TorPriRad = parOperatorDiameter / 2 - TorSecRad
    
    'Prepare data points
    oOriginPos.Set 0, 0, 0
    
    CylBotOrigin.Set oOriginPos.x, oOriginPos.y, oOriginPos.z
    CylTopOrigin.Set oOriginPos.x, oOriginPos.y + parOperatorHeight - parValveFacetoCenter, _
        oOriginPos.z
    
    TorOrigin.Set CylTopOrigin.x, CylTopOrigin.y - TorSecRad, CylTopOrigin.z
    
    'Prepare axis vector for Torus
    TorAxis.Set 0, 1, 0

    'Insert your code for output 1(Stem)
    'Place Cylinder
    Dim ObjStem As IngrGeom3D.Projection3d
    Set ObjStem = PlaceCylinder(m_OutputColl, CylBotOrigin, CylTopOrigin, StemMinimumDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing

    'Insert your code for output 2(Handwheel)
    'Place Torus
    Dim ObjHandwheel As IngrGeom3D.Torus3d
    Set ObjHandwheel = PlaceTorus(m_OutputColl, TorOrigin, TorAxis, TorPriRad, TorSecRad)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheel
    Set ObjHandwheel = Nothing
   
    Set oOriginPos = Nothing
    Set CylBotOrigin = Nothing
    Set CylTopOrigin = Nothing
    Set TorOrigin = Nothing
    Set TorAxis = Nothing
  
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
