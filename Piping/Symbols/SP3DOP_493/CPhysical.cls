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
' Page No. D-141 of PDS Piping Component Data Reference Guide.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   27.Aug.2003     MS/ACM TR47240:  Operator stem diameters are made proportionate so
'                                             that they donot appear for small valve sizes
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
    Dim ObjBox As Object
    
    Dim parStemMinimumDiameter As Double
    Dim parOperatorHeight As Double
    Dim parOperatorHeight1 As Double
    Dim parOperatorWidth As Double
    Dim parOperatorWidth1 As Double
    Dim parOperatorLength As Double
    Dim parOperatorLength1 As Double
    Dim parOperatorDepth As Double
    Dim parActuatorOffset As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength2 As Double
    Dim parActuatorDiameter1 As Double
    
'   Variables defined by svsmylav
    Dim CenterPos       As New AutoMath.DPosition
    Dim StemStPt        As New AutoMath.DPosition
    Dim StemEndPt       As New AutoMath.DPosition
    Dim BoxStCorner   As New AutoMath.DPosition
    Dim BoxOppCorner   As New AutoMath.DPosition
    Dim MainCylStPt     As New AutoMath.DPosition
    Dim MainCylEndPt    As New AutoMath.DPosition
    Dim SecCylStPt      As New AutoMath.DPosition
    Dim SecCylEndPt     As New AutoMath.DPosition
    Dim StemHt          As Double
    Dim AxisOffset      As Double
    Dim BoxLen         As Double
    Dim BoxWid         As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorHeight1 = arrayOfInputs(3)
    parOperatorWidth = arrayOfInputs(4)
    parOperatorWidth1 = arrayOfInputs(5)
    parOperatorLength = arrayOfInputs(6)
    parOperatorLength1 = arrayOfInputs(7)
    parOperatorDepth = arrayOfInputs(8)
    parActuatorOffset = arrayOfInputs(9)
    parActuatorDiameter = arrayOfInputs(10)
    parActuatorLength1 = arrayOfInputs(11)
    parActuatorLength = arrayOfInputs(12)
    parActuatorLength2 = arrayOfInputs(13)
    parActuatorDiameter1 = arrayOfInputs(14)
    
    iOutput = 0
    
'    parStemMinimumDiameter = 0.05

    parStemMinimumDiameter = (parOperatorLength + parOperatorLength1) * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemMinimumDiameter < 0.015 Then parStemMinimumDiameter = 0.015
   If parStemMinimumDiameter > 0.05 Then parStemMinimumDiameter = 0.05
   
'   Prepare Distances
    StemHt = parOperatorHeight - parOperatorDepth
    
    BoxLen = parOperatorLength + parOperatorLength1
    BoxWid = parOperatorWidth + parOperatorWidth1
    
    AxisOffset = parOperatorLength1 + parActuatorDiameter / 2
    
'   Prepare data points
    CenterPos.Set 0, 0, 0

    StemStPt.Set CenterPos.x, CenterPos.y, CenterPos.z
    StemEndPt.Set StemStPt.x, StemStPt.y + StemHt, StemStPt.z
    
    BoxStCorner.Set StemEndPt.x + parOperatorLength1, _
                      StemEndPt.y, StemEndPt.z + parOperatorWidth
                      
    BoxOppCorner.Set BoxStCorner.x - BoxLen, BoxStCorner.y + parOperatorDepth, _
                      BoxStCorner.z - BoxWid
    
    MainCylStPt.Set CenterPos.x + AxisOffset, CenterPos.y + parOperatorHeight - parOperatorHeight1, _
                    CenterPos.z + parActuatorLength1
    MainCylEndPt.Set MainCylStPt.x, MainCylStPt.y, MainCylStPt.z - parActuatorLength
                    
    SecCylStPt.Set MainCylStPt.x, MainCylStPt.y, MainCylStPt.z
    SecCylEndPt.Set SecCylStPt.x, SecCylStPt.y, SecCylStPt.z + parActuatorLength2

 ' Insert your code for output 1(Spindle)
    Dim ObjStem As IngrGeom3D.Projection3d
    Set ObjStem = PlaceCylinder(m_OutputColl, StemStPt, StemEndPt, parStemMinimumDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
 ' Insert your code for output 2(Rectangular Box)
    Set ObjBox = PlaceBox(m_OutputColl, BoxStCorner, BoxOppCorner)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox
    Set ObjBox = Nothing
    
 ' Insert your code for output 3(Main Cylinder)
    Dim ObjMainCyl As IngrGeom3D.Projection3d
    Set ObjMainCyl = PlaceCylinder(m_OutputColl, MainCylStPt, MainCylEndPt, parActuatorDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainCyl
    Set ObjMainCyl = Nothing
    
 ' Insert your code for output 4(Second Cylinder)
    Dim ObjSecCyl As IngrGeom3D.Projection3d
    Set ObjSecCyl = PlaceCylinder(m_OutputColl, SecCylStPt, SecCylEndPt, parActuatorDiameter1, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSecCyl
    Set ObjSecCyl = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
