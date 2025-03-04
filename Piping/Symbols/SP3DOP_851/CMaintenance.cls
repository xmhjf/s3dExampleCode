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
'   Author:          MS
'   Creation Date:  Wednesday, September 22 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   08.22.2004      MS  Added Maintenace Envelope to the operator
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages

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
    Dim parStemExtension As Double
    Dim parOperatorDepth As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorStemMinimumDiameter As Double
    Dim parActuatorWheelMinimumThickness As Double
    
'   Variables defined by svsmylav
    Dim CenterPos           As New AutoMath.DPosition
    Dim StemStPt            As New AutoMath.DPosition
    Dim StemEndPt           As New AutoMath.DPosition
    Dim operatorStartPoint  As New AutoMath.DPosition
    Dim operatorEndPoint    As New AutoMath.DPosition
    Dim OpToHWShaftStPt     As New AutoMath.DPosition
    Dim OpToHWShaftEndPt    As New AutoMath.DPosition
    Dim WheelStPt           As New AutoMath.DPosition
    Dim WheelEndPt          As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parStemExtension = arrayOfInputs(4)
    parOperatorDepth = arrayOfInputs(5)
    parActuatorDiameter = arrayOfInputs(6)


     parStemMinimumDiameter = parOperatorDiameter * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemMinimumDiameter < 0.015 Then parStemMinimumDiameter = 0.015
   If parStemMinimumDiameter > 0.05 Then parStemMinimumDiameter = 0.05
   
   parActuatorStemMinimumDiameter = parStemMinimumDiameter
   If parActuatorStemMinimumDiameter >= parActuatorDiameter Then
        parActuatorStemMinimumDiameter = parActuatorDiameter / 2
  End If
   parActuatorWheelMinimumThickness = parActuatorStemMinimumDiameter

'   Prepare data points
    CenterPos.Set 0, 0, 0
    
    StemStPt.Set CenterPos.x, CenterPos.y, CenterPos.z
    StemEndPt.Set CenterPos.x, CenterPos.y + parOperatorHeight - (parStemExtension + parOperatorDepth), CenterPos.z
    
    operatorStartPoint.Set StemEndPt.x, StemEndPt.y, StemEndPt.z
    operatorEndPoint.Set StemEndPt.x, StemEndPt.y + parOperatorDepth, StemEndPt.z
    
    OpToHWShaftStPt.Set operatorEndPoint.x, operatorEndPoint.y, operatorEndPoint.z
    OpToHWShaftEndPt.Set operatorEndPoint.x, _
    operatorEndPoint.y + parStemExtension - parActuatorWheelMinimumThickness, operatorEndPoint.z
    
    WheelStPt.Set OpToHWShaftEndPt.x, OpToHWShaftEndPt.y, OpToHWShaftEndPt.z
    WheelEndPt.Set OpToHWShaftEndPt.x, OpToHWShaftEndPt.y + parActuatorWheelMinimumThickness, OpToHWShaftEndPt.z
    
'   Initialize
    iOutput = 0

 ' Insert your code for output 5(Maintenance Cylinder 1)
    Dim ObjMaintCyl1 As IngrGeom3D.Projection3d
    Set ObjMaintCyl1 = PlaceCylinder(m_OutputColl, StemStPt, StemEndPt, parActuatorStemMinimumDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl1
    Set ObjMaintCyl1 = Nothing
    
 ' Insert your code for output 6(Maintenance Cylinder 2)
    Dim ObjMaintCyl2 As IngrGeom3D.Projection3d
    Set ObjMaintCyl2 = PlaceCylinder(m_OutputColl, operatorStartPoint, operatorEndPoint, parOperatorDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl2
    Set ObjMaintCyl2 = Nothing
    
 ' Insert your code for output 7(Maintenance Cylinder 3)
    Dim ObjMaintCyl3 As IngrGeom3D.Projection3d
    Set ObjMaintCyl3 = PlaceCylinder(m_OutputColl, OpToHWShaftStPt, OpToHWShaftEndPt, parActuatorStemMinimumDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl3
    Set ObjMaintCyl3 = Nothing
    
 ' Insert your code for output 8(Maintenance Cylinder 4)
    Dim ObjMaintCyl4 As IngrGeom3D.Projection3d
    Set ObjMaintCyl4 = PlaceCylinder(m_OutputColl, WheelStPt, WheelEndPt, parActuatorDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl4
    Set ObjMaintCyl4 = Nothing
        
    Set CenterPos = Nothing
    Set StemStPt = Nothing
    Set StemEndPt = Nothing
    Set operatorStartPoint = Nothing
    Set operatorEndPoint = Nothing
    Set OpToHWShaftStPt = Nothing
    Set OpToHWShaftEndPt = Nothing
    Set WheelStPt = Nothing
    Set WheelEndPt = Nothing
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
