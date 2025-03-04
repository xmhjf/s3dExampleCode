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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   ProgID:         SP3DOPHandwheelD.HandwheelD
'   Author:         RUK
'   Creation Date:  Monday, 11 August 2008
'
'   Description:
'   This Symbol details were taken from Appendix E-138 of the Design Document
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   11.Aug.2008     RUK     CR-146837  Provide needle valve symbol
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:"    'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    
    Dim parOperatorHeight As Double
    Dim parOffsetFrmValCen As Double
    Dim parHandwheelDiameter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOffsetFrmValCen = arrayOfInputs(3)
    parHandwheelDiameter = arrayOfInputs(4)
    
    Dim iOutput As Long
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition
    Dim oVector As AutoMath.DVector
    
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oCenter = New DPosition
    Set oVector = New DVector
    
    Dim dStemDia As Double
    Dim dSpokeDia As Double
    Dim dHwTorDia As Double
    
    dStemDia = 0.05 * parHandwheelDiameter
    dSpokeDia = 0.05 * parHandwheelDiameter
   
    If dStemDia < 0.015 Then dStemDia = 0.015
    If dStemDia > 0.05 Then dStemDia = 0.05
    
    dSpokeDia = dStemDia
    dHwTorDia = dStemDia
    
    'Handwheel Stem
    Dim objHwStem As Object
    oStPoint.Set 0, parOffsetFrmValCen, 0
    oEnPoint.Set 0, parOperatorHeight - dHwTorDia / 2, 0
    
    Set objHwStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dStemDia, True)

    'Add to the Ouput
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHwStem
    Set objHwStem = Nothing
    
    'Handwheel
    Dim objHandwheel As Object
    Dim oLine As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim oCollection As Collection
    Dim oCompStr As IngrGeom3D.ComplexString3d
    
    Set oLine = New Line3d
    Set oArc = New Arc3d
    Set oCollection = New Collection
    Set oCompStr = New ComplexString3d
    
    'Line 1
    oStPoint.Set 0, parOperatorHeight - 0.9 * dHwTorDia, 0
    oEnPoint.Set parHandwheelDiameter / 2 - 0.4 * dHwTorDia, oStPoint.y, 0
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Line 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, parOperatorHeight, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Line 3
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set parHandwheelDiameter / 2, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
        
    'Line 4
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, parOperatorHeight - dHwTorDia, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Line 5
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set parHandwheelDiameter / 2 - 0.1 * dHwTorDia, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Line 6
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, parOperatorHeight - 0.1 * dHwTorDia, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Line 7
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set parHandwheelDiameter / 2 - 0.3 * dHwTorDia, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Line 8
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, parOperatorHeight - dHwTorDia, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Line 9
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set 0, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    oStPoint.Set 0, parOperatorHeight - 0.9 * dHwTorDia, 0
    Set oCompStr = PlaceTrCString(oStPoint, oCollection)
    
    oCenter.Set 0, parOperatorHeight, 0
    oVector.Set 0, 1, 0
    Set objHandwheel = PlaceRevolution(m_OutputColl, oCompStr, oVector, oCenter, 2 * PI, True)

    'Add to the Ouput
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandwheel
    Set objHandwheel = Nothing
    
    Dim iCount As Long
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    
    Set oLine = Nothing
    Set oArc = Nothing
    Set oCollection = Nothing
    Set oCompStr = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenter = Nothing
    Set oVector = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
