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
'   Creation Date:  Sunday, Dec 23 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Page No. D-134 of PDS Piping Component Data Reference Guide.
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

    Dim parStemMinimumDiameter As Double
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parOperCenFromTop As Double
    Dim parOperatorDepth As Double
    Dim parArmLength1 As Double
    Dim parArmLength As Double
    Dim parArmDiameter As Double
    Dim parActuatorOffset As Double
    Dim parActuatorDiameter As Double
    Dim parStemExtension As Double
    Dim parActuatorStemMinimumDiameter As Double
    Dim parActuatorWheelMinimumThickness As Double
    
'   Variables defined by svsmylav
    Dim CenterPos       As New AutoMath.DPosition
    Dim StemStPt        As New AutoMath.DPosition
    Dim StemEndPt       As New AutoMath.DPosition
    Dim operatorStartPoint   As New AutoMath.DPosition
    Dim operatorEndPoint   As New AutoMath.DPosition
    Dim parArmStartPoint   As New AutoMath.DPosition
    Dim parArmEndPoint   As New AutoMath.DPosition
    Dim parHwAxisStartPoint   As New AutoMath.DPosition
    Dim parHwAxisEndPoint   As New AutoMath.DPosition
    Dim WheelStPt   As New AutoMath.DPosition
    Dim WheelEndPt   As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parOperCenFromTop = arrayOfInputs(4)
    parOperatorDepth = arrayOfInputs(5)
    parArmLength1 = arrayOfInputs(6)
    parArmLength = arrayOfInputs(7)
    parArmDiameter = arrayOfInputs(8)
    parActuatorOffset = arrayOfInputs(9)
    parActuatorDiameter = arrayOfInputs(10)
    parStemExtension = arrayOfInputs(11)
    
'   Initialize
    iOutput = 0
    
'    parStemMinimumDiameter = 0.05
'    parActuatorStemMinimumDiameter = 0.05
'    parActuatorWheelMinimumThickness = 0.05

    parStemMinimumDiameter = parOperatorDiameter * 0.1
    '   Check that the stem diameter is within 15mm and 50mm
    If parStemMinimumDiameter < 0.015 Then parStemMinimumDiameter = 0.015
    If parStemMinimumDiameter > 0.05 Then parStemMinimumDiameter = 0.05
    
    parActuatorStemMinimumDiameter = parStemMinimumDiameter
    If parActuatorStemMinimumDiameter >= parArmDiameter / 2 Then
        parActuatorStemMinimumDiameter = parArmDiameter / 2
    End If
    
    parActuatorWheelMinimumThickness = parStemMinimumDiameter
    
'   Prepare data points
    CenterPos.Set 0, 0, 0
    
    StemStPt.Set CenterPos.x, CenterPos.y, CenterPos.z
    StemEndPt.Set CenterPos.x, CenterPos.y + parOperatorHeight, CenterPos.z
    
    operatorStartPoint.Set CenterPos.x, parOperatorHeight - parStemExtension + parOperCenFromTop, CenterPos.z
    operatorEndPoint.Set CenterPos.x, operatorStartPoint.y - parOperatorDepth, CenterPos.z
    
    parArmStartPoint.Set CenterPos.x + parOperatorDiameter / 3, _
                         CenterPos.y + parOperatorHeight - parStemExtension, _
                         CenterPos.z
    parArmEndPoint.Set CenterPos.x + parArmLength, parArmStartPoint.y, CenterPos.z
                       
    parHwAxisStartPoint.Set CenterPos.x + parArmLength1, parArmStartPoint.y, CenterPos.z
    parHwAxisEndPoint.Set parHwAxisStartPoint.x, parArmStartPoint.y, CenterPos.z + parActuatorOffset
    
    WheelEndPt.Set parHwAxisEndPoint.x, parHwAxisEndPoint.y, parHwAxisEndPoint.z
    WheelStPt.Set WheelEndPt.x, WheelEndPt.y, WheelEndPt.z - parActuatorWheelMinimumThickness
  

 ' Insert your code for output 1(Spindle)
    Dim ObjStem As IngrGeom3D.Projection3d
    Set ObjStem = PlaceCylinder(m_OutputColl, StemStPt, StemEndPt, parStemMinimumDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
 ' Insert your code for output 2(Operator)
    Dim ObjOperator As IngrGeom3D.Projection3d
    Set ObjOperator = PlaceCylinder(m_OutputColl, operatorStartPoint, operatorEndPoint, parOperatorDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOperator
    Set ObjOperator = Nothing
    
 ' Insert your code for output 3(Arm)
    Dim ObjArm As IngrGeom3D.Projection3d
    Set ObjArm = PlaceCylinder(m_OutputColl, parArmStartPoint, parArmEndPoint, parArmDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArm
    Set ObjArm = Nothing
    
 ' Insert your code for output 4(Handwheel Axis)
    Dim ObjHWAxis As IngrGeom3D.Projection3d
    Set ObjHWAxis = PlaceCylinder(m_OutputColl, parHwAxisStartPoint, parHwAxisEndPoint, parActuatorStemMinimumDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHWAxis
    Set ObjHWAxis = Nothing
    
 ' Insert your code for output 5(Handwheel)
    Dim ObjWheel As IngrGeom3D.Projection3d
    Set ObjWheel = PlaceCylinder(m_OutputColl, WheelStPt, WheelEndPt, parActuatorDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWheel
    Set ObjWheel = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

