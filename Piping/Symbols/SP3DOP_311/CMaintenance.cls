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
'   This Symbol details were taken from PDS Piping Component Data SN=OP_311 Page # D-133.
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
    Dim parStemAngle As Double
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parActuatorLength As Double
    Dim parOperatorDepth As Double
    Dim parActuatorOffset As Double
    Dim parActuatorDiameter As Double
    Dim parStemExtension As Double
    Dim parActuatorStemMinimumDiameter As Double
    Dim parActuatorWheelMinimumThickness As Double

    Dim CenterPos           As New AutoMath.DPosition
    Dim StemStPt            As New AutoMath.DPosition
    Dim StemEndPt           As New AutoMath.DPosition
    Dim operatorStartPoint  As New AutoMath.DPosition
    Dim operatorEndPoint    As New AutoMath.DPosition
    Dim parHwAxisStartPoint As New AutoMath.DPosition
    Dim parHwAxisEndPoint   As New AutoMath.DPosition
    Dim OpBottom            As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parStemAngle = arrayOfInputs(2) 'Y4
    parOperatorHeight = arrayOfInputs(3) 'Y5
    parOperatorDiameter = arrayOfInputs(4) 'Y6
    parActuatorLength = arrayOfInputs(5) 'Y7
    parOperatorDepth = arrayOfInputs(6)  'Y8
    parActuatorOffset = arrayOfInputs(7) 'Y9
    parActuatorDiameter = arrayOfInputs(8) 'Y10
    parStemExtension = arrayOfInputs(9)  'Y11
    
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
    If parActuatorStemMinimumDiameter >= parActuatorDiameter / 2 Then
        parActuatorStemMinimumDiameter = parActuatorDiameter / 2
    End If
    
    parActuatorWheelMinimumThickness = parStemMinimumDiameter

'   Prepare distances
    OpBottom = parOperatorHeight - (parActuatorLength + parOperatorDepth)   'Y5 - Y7 - Y8
    
'   Prepare data points
    CenterPos.Set 0, 0, 0
       
 ' Insert your code for output 5(Maintenance Spindle)
    Dim ObjMaintStem As Object

    StemStPt.Set CenterPos.x, CenterPos.y, CenterPos.z
    StemEndPt.Set CenterPos.x - OpBottom * Cos(parStemAngle), _
                  CenterPos.y + OpBottom * Sin(parStemAngle), _
                  CenterPos.z
'   A value of 0.0000001 is used in Z direction to avoid symbol placement failure for specific angles(gives assertion errors).
'   This should be removed  once it is resolved.
    StemStPt.z = StemStPt.z + 0.0000001
    Set ObjMaintStem = PlaceCone(m_OutputColl, StemEndPt, StemStPt, parOperatorDiameter / 2, 0, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintStem
    Set ObjMaintStem = Nothing
    
 ' Insert your code for output 6(Maintenance Operator)
    operatorStartPoint.Set -OpBottom * Cos(parStemAngle), _
                            OpBottom * Sin(parStemAngle), _
                            CenterPos.z
    operatorEndPoint.Set operatorStartPoint.x - parOperatorDepth * Cos(parStemAngle), _
                         operatorStartPoint.y + parOperatorDepth * Sin(parStemAngle), _
                         CenterPos.z
                
    Dim ObjMaintOperator As IngrGeom3D.Projection3d
    Set ObjMaintOperator = PlaceCylinder(m_OutputColl, operatorStartPoint, operatorEndPoint, parOperatorDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintOperator
    Set ObjMaintOperator = Nothing
    
 ' Insert your code for output 7(Handwheel Axis )
    parHwAxisStartPoint.Set operatorStartPoint.x + parActuatorOffset * Sin(parStemAngle) - parOperatorDepth * Cos(parStemAngle), _
                            operatorStartPoint.y + parActuatorOffset * Cos(parStemAngle) + parOperatorDepth * Sin(parStemAngle), _
                            CenterPos.z
    parHwAxisEndPoint.Set parHwAxisStartPoint.x - parActuatorLength * Cos(parStemAngle), _
                          parHwAxisStartPoint.y + parActuatorLength * Sin(parStemAngle), _
                          CenterPos.z
    Dim ObjMaintHandwheelAxis As IngrGeom3D.Projection3d
    Set ObjMaintHandwheelAxis = PlaceCylinder(m_OutputColl, parHwAxisStartPoint, parHwAxisEndPoint, parActuatorStemMinimumDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintHandwheelAxis
    Set ObjMaintHandwheelAxis = Nothing
    Set CenterPos = Nothing
    Set StemStPt = Nothing
    Set StemEndPt = Nothing
    Set operatorStartPoint = Nothing
    Set operatorEndPoint = Nothing
    Set parHwAxisStartPoint = Nothing
    Set parHwAxisEndPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
