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
' Page No. D-140 of PDS Piping Component Data Reference Guide.
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
    Dim parStemAngle As Double
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parOperatorDepth As Double
    Dim parActuatorOffset As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorDepth As Double

'   Variables defined by svsmylav
    Dim CenterPos       As New AutoMath.DPosition
    Dim StemStPt        As New AutoMath.DPosition
    Dim StemEndPt       As New AutoMath.DPosition
    Dim MainCylStPt     As New AutoMath.DPosition
    Dim MainCylEndPt    As New AutoMath.DPosition
    Dim SecCylStPt      As New AutoMath.DPosition
    Dim SecCylEndPt     As New AutoMath.DPosition
    Dim StemHt          As Double
    Dim AxisOffset      As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parStemAngle = arrayOfInputs(2)
    parOperatorHeight = arrayOfInputs(3)
    parOperatorDiameter = arrayOfInputs(4)
    parOperatorDepth = arrayOfInputs(5)
    parActuatorOffset = arrayOfInputs(6)
    parActuatorDiameter = arrayOfInputs(7)
    parActuatorDepth = arrayOfInputs(8)

    
'   Initialize
    iOutput = 0
    
'    parStemMinimumDiameter = 0.05

    parStemMinimumDiameter = parOperatorDiameter * 0.1
    '   Check that the stem diameter is within 15mm and 50mm
    If parStemMinimumDiameter < 0.015 Then parStemMinimumDiameter = 0.015
    If parStemMinimumDiameter > 0.05 Then parStemMinimumDiameter = 0.05
    
'   Prepare Distances
    StemHt = parOperatorHeight - parOperatorDepth
    
    AxisOffset = (parOperatorDiameter + parActuatorDiameter) / 2
    
'   Prepare data points
    CenterPos.Set 0, 0, 0

    StemStPt.Set CenterPos.x, CenterPos.y, CenterPos.z
    StemEndPt.Set StemStPt.x - StemHt * Cos(parStemAngle), _
                  StemStPt.y + StemHt * Sin(parStemAngle), StemStPt.z
    
    MainCylStPt.Set StemEndPt.x, StemEndPt.y, StemEndPt.z
    MainCylEndPt.Set MainCylStPt.x - parOperatorDepth * Cos(parStemAngle), _
                    MainCylStPt.y + parOperatorDepth * Sin(parStemAngle), MainCylStPt.z
                    
    SecCylStPt.Set MainCylStPt.x + AxisOffset * Sin(parStemAngle), _
                   MainCylStPt.y + AxisOffset * Cos(parStemAngle), MainCylStPt.z
                   
    SecCylEndPt.Set SecCylStPt.x - parActuatorDepth * Cos(parStemAngle), _
                    SecCylStPt.y + parActuatorDepth * Sin(parStemAngle), SecCylStPt.z

' Insert your code for output 1(Spindle)
    Dim ObjStem As IngrGeom3D.Projection3d
    Set ObjStem = PlaceCylinder(m_OutputColl, StemStPt, StemEndPt, parStemMinimumDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
' Insert your code for output 2(Main Cylinder)
    Dim ObjMainCyl As IngrGeom3D.Projection3d
    Set ObjMainCyl = PlaceCylinder(m_OutputColl, MainCylStPt, MainCylEndPt, parOperatorDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainCyl
    Set ObjMainCyl = Nothing
    
' Insert your code for output 3(Second Cylinder)
    Dim ObjSecCyl As IngrGeom3D.Projection3d
    Set ObjSecCyl = PlaceCylinder(m_OutputColl, SecCylStPt, SecCylEndPt, parActuatorDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSecCyl
    Set ObjSecCyl = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


