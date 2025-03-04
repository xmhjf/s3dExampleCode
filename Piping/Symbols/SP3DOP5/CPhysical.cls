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
'   Author:          svsmylav
'   Creation Date:  Saturday, Dec 22 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Page No. D-130 of PDS Piping Component Data Reference Guide.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   27.Aug.2003     MS/ACM TR47240:  Operator stem diameters are made proportionate so
'					      that they donot appear for small valve sizes
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    
    Dim parOperatorHeight As Double
    Dim parStemMinimumDiameter As Double
    Dim parStemAngle As Double
    Dim parOperatorDiameter As Double
    Dim parActuatorWheelMinimumThickness As Double
    Dim parStemExtension As Double
    
'   Variables defined by svsmylav
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
    parStemAngle = arrayOfInputs(3)
    parOperatorDiameter = arrayOfInputs(4)
    
'   Initialize
    iOutput = 0
    
'    parStemMinimumDiameter = 0.05
'    parActuatorWheelMinimumThickness = 0.05
'    parStemExtension = parOperatorDiameter / 3

    parStemMinimumDiameter = parOperatorDiameter * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemMinimumDiameter < 0.015 Then parStemMinimumDiameter = 0.015
   If parStemMinimumDiameter > 0.05 Then parStemMinimumDiameter = 0.05
   
   parActuatorWheelMinimumThickness = parStemMinimumDiameter
   parStemExtension = parOperatorHeight * 0.1


'   Prepare Torus Radii
    TorSecRad = parActuatorWheelMinimumThickness / 2
    TorPriRad = parOperatorDiameter / 2 - TorSecRad
    
'   Prepare data points
    CenterPos.Set 0, 0, 0
    
    CylBotOrigin.Set CenterPos.x, CenterPos.y, CenterPos.z
    CylTopOrigin.Set CylBotOrigin.x - parOperatorHeight * Cos(parStemAngle), _
                     CylBotOrigin.y + parOperatorHeight * Sin(parStemAngle), _
                     CylBotOrigin.z
    
    TorOrigin.Set CylTopOrigin.x + (parStemExtension + TorSecRad) * Cos(parStemAngle), _
                  CylTopOrigin.y - (parStemExtension + TorSecRad) * Sin(parStemAngle), _
                  CylTopOrigin.z
    
'   Prepare axis vector for Torus
    TorAxis.Set -Cos(parStemAngle), Sin(parStemAngle), 0

' Insert your code for output 1(Spindle for handwheel)
'   Place Cylinder
    Dim ObjCylinder As IngrGeom3D.Projection3d
    Set ObjCylinder = PlaceCylinder(m_OutputColl, CylBotOrigin, CylTopOrigin, parStemMinimumDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
    Set ObjCylinder = Nothing
    
' Insert your code for output 2(Handwheel)
'   Place Torus
    Dim objTorus    As IngrGeom3D.Torus3d

    Set objTorus = PlaceTorus(m_OutputColl, TorOrigin, TorAxis, TorPriRad, TorSecRad)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTorus
    Set objTorus = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub




