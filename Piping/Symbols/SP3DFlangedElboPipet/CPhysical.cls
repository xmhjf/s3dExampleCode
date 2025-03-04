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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         dkl
'   Creation Date:  Thursday, Mar 8 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is based on the SP3DElbolet.
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   27.Mar.2007     dkl     CR 116207. Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI       As Double
Private Const MODULE = "SimplePhysical:" 'Used for error messages

Private Sub Class_Initialize()
   PI = 4 * Atn(1)
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
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2     As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0
    
' Place Elbolet Body
    Dim objElboletBody As Object
    Dim LineStrPoints(0 To 20) As Double
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oAxisVect As AutoMath.DVector
    Set oAxisVect = New AutoMath.DVector
    Dim oCenPoint As AutoMath.DPosition
    Set oCenPoint = New AutoMath.DPosition
    
'   Assumptions:
'      - Port1 and the center are coincident.
'      - Length of the conical part of the Elbo-pipet is .5 x Dia of the header pipe.
'      - Radius of the Conical end is .65 x Dia of the branch pipe.
'      - Length of the cylinderical part is .7 x Face to Face.
'      - Length of the beveled slope is .2 x Face to Face.
'      - Length of the Tapered part before the flange is .1 x Face to Face.
'      - Radius of the Tapered part before the flange is flange Diameter / 2 corresponding
'                                                                  to the branch pipe.

    Dim dLengthOfConicalEnd As Double
    Dim dLengthOfBevelSlope As Double
    Dim dLengthOfCylindericalPart As Double
    Dim dLengthOfTaperBeforeFlange As Double
    Dim dRadiusOfConicalEnd As Double

    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                  sptOffset, depth
                                                                        
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                  sptOffset2, depth2
    
    ' The Face to Face dimension of the Elbo-Pipet includes Flange Thickness. As this symbol
    ' is meant only for Flanged End preparation,the length of the Elbo-pipet Body computed
    ' in the symbol excludes the Flange Thickness.
    
    Dim dLengthOfElboPipetBody As Double
    dLengthOfElboPipetBody = parFacetoFace - flangeThick2
    
    dLengthOfConicalEnd = 0.5 * pipeDiam
    dLengthOfBevelSlope = 0.2 * dLengthOfElboPipetBody
    dLengthOfCylindericalPart = 0.7 * dLengthOfElboPipetBody
    dLengthOfTaperBeforeFlange = 0.1 * dLengthOfElboPipetBody
    dRadiusOfConicalEnd = 0.65 * pipeDiam2
    
    ' Point 1
    LineStrPoints(0) = -dLengthOfConicalEnd
    LineStrPoints(1) = 0
    LineStrPoints(2) = 0
    ' Point 2
    LineStrPoints(3) = 0
    LineStrPoints(4) = LineStrPoints(1) + dRadiusOfConicalEnd
    LineStrPoints(5) = LineStrPoints(2)
    ' Point 3
    LineStrPoints(6) = LineStrPoints(3) + dLengthOfBevelSlope
    LineStrPoints(7) = LineStrPoints(1) + 0.5 * pipeDiam2
    LineStrPoints(8) = LineStrPoints(2)
    ' Point 4
    LineStrPoints(9) = LineStrPoints(6) + dLengthOfCylindericalPart
    LineStrPoints(10) = LineStrPoints(7)
    LineStrPoints(11) = LineStrPoints(2)
    ' Point 5
    LineStrPoints(12) = LineStrPoints(9) + dLengthOfTaperBeforeFlange
    LineStrPoints(13) = flangeDiam2 / 2
    LineStrPoints(14) = LineStrPoints(2)
    ' Point 6
    LineStrPoints(15) = LineStrPoints(12)
    LineStrPoints(16) = LineStrPoints(1)
    LineStrPoints(17) = LineStrPoints(2)
    ' Point 7
    LineStrPoints(18) = LineStrPoints(0)
    LineStrPoints(19) = LineStrPoints(1)
    LineStrPoints(20) = LineStrPoints(2)
    
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    oAxisVect.Set 1, 0, 0
    oCenPoint.Set 0, -0.00001, 0
    Set objElboletBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenPoint, _
                                2 * PI, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objElboletBody
    Set objElboletBody = Nothing
    Set oAxisVect = Nothing
    Set oCenPoint = Nothing
    Set geomFactory = Nothing
    
' Place Nozzle 1

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace  As Double
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set 0, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    
    oPlacePoint.Set parFacetoFace + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

