VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHatchCoverHole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RUK
'   Creation Date:  Monday, July 28 2008
'
'   Description:
'               Hatch Cover Assembly.
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   28.07.2008      RUK     CR-147109  Provide marine manhole and hatch cover equipment symbols with corner radius
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CHatchCoverHole:" 'Used for error messages
Private Const dInch = 2.54 / 100
Private Const NEGLIGIBLE_THICKNESS = 0.0001

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parCoverLength As Double
    Dim parCoverWidth As Double
    Dim parCoverThickness As Double
    Dim parCoverHeight As Double
    Dim parCornerRadius As Double
    Dim parDistBetBoltHoles As Double
    Dim parHoleLength As Double
    Dim parHoleWidth As Double
    Dim parHoleThickness As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCoverLength = arrayOfInputs(2)
    parCoverWidth = arrayOfInputs(3)
    parCoverThickness = arrayOfInputs(4)
    parCoverHeight = arrayOfInputs(5)
    parCornerRadius = arrayOfInputs(6)
    parDistBetBoltHoles = arrayOfInputs(7)
    parHoleLength = arrayOfInputs(8)
    parHoleWidth = arrayOfInputs(9)
    parHoleThickness = arrayOfInputs(10)
   
    iOutput = 0
    
    Dim oCenPoint As AutoMath.DPosition
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oVector As AutoMath.DVector
    Dim iCount As Integer
    
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oCenPoint = New DPosition
    Set oVector = New DVector
    
    'Calculate the Corner radius such that shymbol shouldn't fail
    Dim dTempDim As Double
    If CmpDblLessThan(parCoverLength, parCoverWidth) Then
        dTempDim = parCoverLength
    Else
        dTempDim = parCoverWidth
    End If
    
    If CmpDblLessThanOrEqualTo((dTempDim - 12 * dInch) / 2, parCornerRadius) Then
        parCornerRadius = (dTempDim - 12 * dInch) / 2 - NEGLIGIBLE_THICKNESS
    End If
    
    'Create the Hole
    Dim objHole As Object
    Dim oCompStr As IngrGeom3D.ComplexString3d
    Dim oCollection As Collection
    Dim oLine As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    
    Set oCompStr = New ComplexString3d
    Set oCollection = New Collection
    Set oLine = New Line3d
    Set oArc = New Arc3d
    
    'Create the Outer Complex String
    Dim dArcRad As Double
    dArcRad = parCornerRadius + 6 * dInch + (parHoleLength - parCoverLength) / 2
    
    'Line 1
    oStPoint.Set parHoleLength / 2 - dArcRad, parHoleWidth / 2, 0
    oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 1
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set -parHoleLength / 2, (parHoleWidth / 2 - dArcRad), oStPoint.z
    oCenPoint.Set oStPoint.x, oEnPoint.y, oEnPoint.z
    Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    'Line 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, -(parHoleWidth / 2 - dArcRad), oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set -(parHoleLength / 2 - dArcRad), -parHoleWidth / 2, oStPoint.z
    oCenPoint.Set oEnPoint.x, oStPoint.y, oStPoint.z
    Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    'Line 3
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set parHoleLength / 2 - dArcRad, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 3
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set parHoleLength / 2, -(parHoleWidth / 2 - dArcRad), oStPoint.z
    oCenPoint.Set oStPoint.x, oEnPoint.y, oEnPoint.z
    Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    'Line 4
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, (parHoleWidth / 2 - dArcRad), oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 4
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set parHoleLength / 2 - dArcRad, parHoleWidth / 2, oStPoint.z
    oCenPoint.Set oEnPoint.x, oStPoint.y, oStPoint.z
    Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    oStPoint.Set parHoleLength / 2 - dArcRad, parHoleWidth / 2, 0
    Set oCompStr = PlaceTrCString(oStPoint, oCollection)
    
    oVector.Set 0, 0, -1
    Set objHole = PlaceProjection(m_OutputColl, oCompStr, oVector, parHoleThickness, True)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHole
    Set objHole = Nothing
    
    'Remove the References
    Set oCenPoint = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oVector = Nothing
    Set oCompStr = Nothing
    Set oLine = Nothing
    Set oArc = Nothing
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    Set oCollection = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub


