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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Oct 1 2004
'   Description:
'   This Symbol details were taken from PDS Piping Component Data Page No. D-134.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   01.Oct.2004     svsmylav       CR-57807 Added Maintenace Envelope to the operator
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
  

 ' Insert your code for output 6(Maintenance Spindle)
    Dim ObjMaintStem As IngrGeom3D.Projection3d
    Set ObjMaintStem = PlaceCylinder(m_OutputColl, StemStPt, StemEndPt, parStemMinimumDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintStem
    Set ObjMaintStem = Nothing
    
 ' Insert your code for output 7(Maintenance Operator)
    Dim ObjMaintOperator As IngrGeom3D.Projection3d
    Set ObjMaintOperator = PlaceCylinder(m_OutputColl, operatorStartPoint, operatorEndPoint, parOperatorDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintOperator
    Set ObjMaintOperator = Nothing
    
 ' Insert your code for output 8(Maintenance Arm)
    Dim ObjMaintArm As IngrGeom3D.Projection3d
    Set ObjMaintArm = PlaceCylinder(m_OutputColl, parArmStartPoint, parArmEndPoint, parArmDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintArm
    Set ObjMaintArm = Nothing
    
 ' Insert your code for output 9(Maintenance Handwheel Axis)
    Dim ObjMaintHWAxis As IngrGeom3D.Projection3d
    Set ObjMaintHWAxis = PlaceCylinder(m_OutputColl, parHwAxisStartPoint, parHwAxisEndPoint, parActuatorStemMinimumDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintHWAxis
    Set ObjMaintHWAxis = Nothing
    
 ' Insert your code for output 10(Maintenance Handwheel)
    Dim ObjMaintWheel As IngrGeom3D.Projection3d
    Set ObjMaintWheel = PlaceCylinder(m_OutputColl, WheelStPt, WheelEndPt, parActuatorDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintWheel
    Set ObjMaintWheel = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

