VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Jul 15 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Swivel Joint Clamp symbol details are provided in biopharm.pdf, MC=CSJ, page no. 84 (Model A12MPS of
'    Tri-clamp is represented by this symbol)
'    The Symbol is created with five Outputs - one insualtion Output and other are physical outputs.
'    In the PDS Eden proportions of inputs are used to generate dimensions requred for preparing the symbol.
'    This symbol uses the same proportions.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI       As Double
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    PI = 4 * Atn(1)
     Set m_oSymGeomHelper = New SymbolServices
    
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
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

    Dim objWingNutEar As Object
    Dim objWingNut As Object

    Dim parClampDiameter As Double
    Dim parClampWidth As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parClampDiameter = arrayOfInputs(2)
    parClampWidth = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)

    m_oSymGeomHelper.OutputCollection = m_OutputColl

    iOutput = 0

'   Center of the origin is at the center of clamp body

'   Insert your code for output 2(Clamp Body)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    stPoint.Set -parClampWidth / 2, 0, 0
    enPoint.Set stPoint.x + parClampWidth, stPoint.y, stPoint.z
    
'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parClampDiameter

'   Insert your code for output 3(Wing Nut Ear)
    Dim dEarLength As Double

    dEarLength = parClampDiameter * 0.625
    stPoint.Set -parClampWidth / 2, -parClampWidth / 2, 0
    enPoint.Set stPoint.x + parClampWidth, stPoint.y + parClampWidth, stPoint.z + dEarLength
    Set objWingNutEar = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objWingNutEar
    Set objWingNutEar = Nothing

'   Insert your code for output 4(Wing Nut Base)
    Dim dCylinderDimension As Double 'Used for height and diameter
    
    dCylinderDimension = 0.75 * parClampWidth

    stPoint.Set 0, parClampWidth / 2, dEarLength - parClampWidth / 2
    enPoint.Set stPoint.x, stPoint.y + dCylinderDimension, stPoint.z
    
'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDimension
    Set stPoint = Nothing
    Set enPoint = Nothing

'   Insert your code for output 5(Wing Nut)
    Dim dDim6 As Double
    Dim dDim8 As Double
    
    dDim6 = dCylinderDimension / 2
    dDim8 = parClampWidth * 2

    Dim LineStrPoints(0 To 11) As Double

    Dim oLineString As IngrGeom3D.LineString3d

'   Point 1
    LineStrPoints(0) = 0
    LineStrPoints(1) = parClampWidth / 2 + dCylinderDimension * 0.75
    LineStrPoints(2) = dEarLength - parClampWidth / 2 - dDim6 / 2
    
'   Point 2
    LineStrPoints(3) = LineStrPoints(0) + dDim8 * Cos(PI / 3)
    LineStrPoints(4) = LineStrPoints(1) + dDim8 * Sin(PI / 3)
    LineStrPoints(5) = LineStrPoints(2)
    
'   Point 3
    LineStrPoints(6) = LineStrPoints(0) - dDim8 * Cos(PI / 3)
    LineStrPoints(7) = LineStrPoints(1) + dDim8 * Sin(PI / 3)
    LineStrPoints(8) = LineStrPoints(2)
    
'   Point 4
    LineStrPoints(9) = LineStrPoints(0)
    LineStrPoints(10) = LineStrPoints(1)
    LineStrPoints(11) = LineStrPoints(2)

    Dim oProjVec As AutoMath.DVector
    Set oProjVec = New AutoMath.DVector
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New GeometryFactory

    oProjVec.Set 0, 0, 1
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    Set objWingNut = PlaceProjection(m_OutputColl, oLineString, oProjVec, dDim6, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objWingNut
    Set objWingNut = Nothing
    Set oLineString = Nothing
    Set oGeomFactory = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
