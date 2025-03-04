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

    Dim ObjMaintBox1 As Object
    
    Dim parStemMinimumDiameter As Double
    Dim parOperatorHeight As Double
    Dim parOperatorHeight1 As Double
    Dim parOperatorLength As Double
    Dim parOperatorLength1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorCenter As Double
    Dim parOperatorWidth1 As Double
    Dim parOperatorWidth As Double
    Dim parOperatorDepth As Double

'   Variables defined by svsmylav
    Dim CenterPos       As New AutoMath.DPosition
    Dim StemStPt        As New AutoMath.DPosition
    Dim StemEndPt       As New AutoMath.DPosition
    Dim BoxStCorner     As New AutoMath.DPosition
    Dim BoxOppCorner    As New AutoMath.DPosition
    Dim RHSCylStPt       As New AutoMath.DPosition
    Dim RHSCylEndPt      As New AutoMath.DPosition
    Dim LHSCylStPt       As New AutoMath.DPosition
    Dim LHSCylEndPt      As New AutoMath.DPosition
    Dim StemHt          As Double
    Dim BoxLen          As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorHeight1 = arrayOfInputs(3)
    parOperatorLength = arrayOfInputs(4)
    parOperatorLength1 = arrayOfInputs(5)
    parActuatorLength = arrayOfInputs(6)
    parActuatorLength1 = arrayOfInputs(7)
    parActuatorDiameter = arrayOfInputs(8)
    parActuatorCenter = arrayOfInputs(9)
    parOperatorWidth1 = arrayOfInputs(10)
    parOperatorWidth = arrayOfInputs(11)
    parOperatorDepth = arrayOfInputs(12)
    
'   Initialize
    iOutput = 0

'    parStemMinimumDiameter = 0.05
    
    parStemMinimumDiameter = (parOperatorLength + parOperatorLength1) * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemMinimumDiameter < 0.015 Then parStemMinimumDiameter = 0.015
   If parStemMinimumDiameter > 0.05 Then parStemMinimumDiameter = 0.05
   
'   Prepare Distances
    StemHt = parOperatorHeight - parOperatorDepth
    
    BoxLen = parOperatorLength + parOperatorLength1
    
'   Prepare data points
    CenterPos.Set 0, 0, 0

    StemStPt.Set CenterPos.x, CenterPos.y, CenterPos.z
    StemEndPt.Set StemStPt.x, StemStPt.y + StemHt, StemStPt.z
    
    BoxStCorner.Set StemStPt.x + parOperatorLength1, StemEndPt.y, _
                    StemStPt.z + parOperatorWidth - (parActuatorCenter + parOperatorWidth1)
                      
    BoxOppCorner.Set BoxStCorner.x - BoxLen, BoxStCorner.y + parOperatorDepth, _
                      BoxStCorner.z - parOperatorWidth
                      
    RHSCylStPt.Set BoxStCorner.x, parOperatorHeight - parOperatorHeight1, _
                    BoxOppCorner.z + parActuatorCenter
                    
    RHSCylEndPt.Set RHSCylStPt.x + parActuatorLength, RHSCylStPt.y, RHSCylStPt.z
                    
    LHSCylStPt.Set BoxOppCorner.x, RHSCylStPt.y, RHSCylStPt.z
    LHSCylEndPt.Set LHSCylStPt.x - parActuatorLength1, LHSCylStPt.y, LHSCylStPt.z
    
 ' Insert your code for output 5 ( Maintenance Cylinder 1)
    Dim ObjMaintCyl1 As IngrGeom3D.Projection3d
    Set ObjMaintCyl1 = PlaceCylinder(m_OutputColl, StemStPt, StemEndPt, parStemMinimumDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl1
    Set ObjMaintCyl1 = Nothing
    
 ' Insert your code for output 6(Maintenance Box 1)
    Set ObjMaintBox1 = PlaceBox(m_OutputColl, BoxStCorner, BoxOppCorner)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintBox1
    Set ObjMaintBox1 = Nothing
    
 ' Insert your code for output 7(Maintenance Cylinder 2)
    Dim ObjMaintCyl2 As IngrGeom3D.Projection3d
    Set ObjMaintCyl2 = PlaceCylinder(m_OutputColl, RHSCylStPt, RHSCylEndPt, parActuatorDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl2
    Set ObjMaintCyl2 = Nothing
    
 ' Insert your code for output 8(Maintenance Cylinder 3)
    Dim ObjMaintCyl3 As IngrGeom3D.Projection3d
    Set ObjMaintCyl3 = PlaceCylinder(m_OutputColl, LHSCylStPt, LHSCylEndPt, parActuatorDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl3
    Set ObjMaintCyl3 = Nothing
    
        
    Set CenterPos = Nothing
    Set StemStPt = Nothing
    Set StemEndPt = Nothing
    Set BoxStCorner = Nothing
    Set BoxOppCorner = Nothing
    Set RHSCylStPt = Nothing
    Set RHSCylEndPt = Nothing
    Set LHSCylStPt = Nothing
    Set LHSCylEndPt = Nothing
        
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
