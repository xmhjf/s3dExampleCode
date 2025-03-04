VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMHCoverHole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CMHCoverHole.cls
'   Author:         RUK
'   Creation Date:  Thursday, December 6 2007
'
'   Description:
'       CR-131706  Provide manhole and hatch cover equipment symbols.
'               Manhole Cover Assembly.
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   06.12.2007      RUK     Created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CMHCoverHole:" 'Used for error messages
Private Const dInch = 2.54 / 100

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parCoverLength As Double
    Dim parCoverWidth As Double
    Dim parCoverThickness As Double
    Dim parCoverHeight As Double
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
    parDistBetBoltHoles = arrayOfInputs(6)
    parHoleLength = arrayOfInputs(7)
    parHoleWidth = arrayOfInputs(8)
    parHoleThickness = arrayOfInputs(9)
    
    'Origin is at DP1/PP1
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
    
    'Create the Complex String
    Dim dOuterArcRad As Double
    dOuterArcRad = parHoleWidth / 2
    
    'Line 1
    oStPoint.Set parHoleLength / 2 - dOuterArcRad, parHoleWidth / 2, 0
    oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 1
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, -oStPoint.y, oStPoint.z
    oCenPoint.Set -parHoleLength / 2, 0, oEnPoint.z
    Set oArc = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    'Line 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, -oStPoint.y, oStPoint.z
    oCenPoint.Set parHoleLength / 2, 0, oStPoint.z
    Set oArc = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    oStPoint.Set parHoleLength / 2 - dOuterArcRad, parHoleWidth / 2, 0
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


