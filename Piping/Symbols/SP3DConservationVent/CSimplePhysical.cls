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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         dkl
'   Creation Date:  Wednesday, Jul 13 2006
'   Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     Conservation Vent (Source: Shand&Jurs, 94020 Conservation Vent, www.ljtechnologies.com/sj/pdf/94020-2B.pdf)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper

Private PI       As Double
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    PI = 4 * Atn(1)
    Set m_oGeomHelper = New SymbolServices
  
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_terminate()
    Set m_oGeomHelper = Nothing
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

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
        
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) ' K
    parFace2toCenter = arrayOfInputs(3)  ' L
    parInstrumentHeight1 = arrayOfInputs(4)  ' J
    parInstrumentLength = arrayOfInputs(5)  ' B

    m_oGeomHelper.OutputCollection = m_OutputColl
    
    iOutput = 0
' Insert your code for output 1(Top Cylinder)
 ' dCylHeight is height of this cylinder above the base trapezoid(that is placed below).It is not the height of this cyl.
    Dim dCylHeight As Double
    Dim dRadTopCyl As Double
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    dCylHeight = parFace1toCenter
    dRadTopCyl = 0.52 * parFace1toCenter  ' Assumption: radius of small cyl is 0.52 parFace1toCenter
    stPoint.Set dCylHeight / 2, 0, 0
    enPoint.Set -parFace1toCenter, 0, 0
                                        
'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 2 * dRadTopCyl

'Top Conical Lid
    Dim objTopConicalLid As Object
    Dim LineStrPoints(0 To 17) As Double
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim oLineString As IngrGeom3D.LineString3d
    Dim axisVect As AutoMath.DVector
    Set axisVect = New AutoMath.DVector
    Dim cenPoint As AutoMath.DPosition
    Set cenPoint = New AutoMath.DPosition
    
' Assumptions :
'  1. The height of the base cone (dBaseConeHeight) is dConeheight / 8 (dConeheight is total height of the two cones)
'  2. The top radius of the base cone (dTopRBaseCone) is dConeheight / 8
'  3. The top radius of the Upper cone (dTopRUpperCone) is dConeheight / 16
    
    Dim dConeheight As Double
    Dim dBaseConeHeight As Double
    Dim dTopRBaseCone As Double
    Dim dTopRUpperCone As Double
    
    dConeheight = parInstrumentHeight1 - parFace1toCenter - dCylHeight / 2
    dBaseConeHeight = dConeheight / 8
    dTopRBaseCone = dConeheight / 8
    dTopRUpperCone = dConeheight / 16
    
    LineStrPoints(0) = dCylHeight / 2
    LineStrPoints(1) = dRadTopCyl
    LineStrPoints(2) = 0
    
    LineStrPoints(3) = dCylHeight / 2 + dBaseConeHeight
    LineStrPoints(4) = dTopRBaseCone
    LineStrPoints(5) = LineStrPoints(2)
    
    LineStrPoints(6) = dCylHeight / 2 + dConeheight
    LineStrPoints(7) = dTopRUpperCone
    LineStrPoints(8) = LineStrPoints(2)
    
    LineStrPoints(9) = dCylHeight / 2 + dConeheight
    LineStrPoints(10) = 0
    LineStrPoints(11) = LineStrPoints(2)
    
    LineStrPoints(12) = -dCylHeight / 2
    LineStrPoints(13) = 0
    LineStrPoints(14) = LineStrPoints(2)
        
    LineStrPoints(15) = LineStrPoints(0)
    LineStrPoints(16) = LineStrPoints(1)
    LineStrPoints(17) = LineStrPoints(2)
    
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, LineStrPoints)
    axisVect.Set 1, 0, 0
    cenPoint.Set 0, -0.00001, 0
    Set objTopConicalLid = PlaceRevolution(m_OutputColl, oLineString, axisVect, cenPoint, _
                                2 * PI, True)

'   Set the outputoLineString
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTopConicalLid
    Set objTopConicalLid = Nothing

' Insert your code for output 3(Lower Cylinder)
    Dim dRadLowerCyl As Double
    dRadLowerCyl = parFace1toCenter * 0.7 ' Assumption: radius of Lower cyl is 0.7*parFace1toCenter
    stPoint.Set -dCylHeight / 2, parInstrumentLength - dRadLowerCyl, 0
    enPoint.Set -parFace1toCenter, parInstrumentLength - dRadLowerCyl, 0
                                        
'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 2 * dRadLowerCyl
    Set stPoint = Nothing
    Set enPoint = Nothing

' Insert your code for output 4(Projection of Trapezoid - Box between large and small cylinders)
    Dim objProjTrapezoid As Object
    Dim LineStrPoints1(0 To 14) As Double
            
    LineStrPoints1(0) = -dCylHeight / 2
    LineStrPoints1(1) = 0
    LineStrPoints1(2) = dRadTopCyl
    
    LineStrPoints1(3) = LineStrPoints1(0)
    LineStrPoints1(4) = parInstrumentLength - dRadLowerCyl
    LineStrPoints1(5) = dRadLowerCyl
    
    LineStrPoints1(6) = LineStrPoints1(0)
    LineStrPoints1(7) = LineStrPoints1(4)
    LineStrPoints1(8) = -dRadLowerCyl
    
    LineStrPoints1(9) = LineStrPoints1(0)
    LineStrPoints1(10) = 0
    LineStrPoints1(11) = -dRadTopCyl
    
    LineStrPoints1(12) = LineStrPoints1(0)
    LineStrPoints1(13) = LineStrPoints1(1)
    LineStrPoints1(14) = LineStrPoints1(2)
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPoints1)
    axisVect.Set -1, 0, 0
   
    Set objProjTrapezoid = PlaceProjection(m_OutputColl, oLineString, axisVect, parFace1toCenter / 2, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objProjTrapezoid
    Set objProjTrapezoid = Nothing
 
'Lower Conical Lid
    Dim objLowerConicalLid As Object
    Dim LineStrPoints2(0 To 17) As Double
    Dim Factor As Double
    Factor = dRadLowerCyl * 0.25 ' Assumption - Base cone radius is .25 * dRadLowerCyl
    LineStrPoints2(0) = -dCylHeight / 2
    LineStrPoints2(1) = parInstrumentLength - Factor 'Subtracting Factor this way prevents overlap with top cylinder
    LineStrPoints2(2) = 0
    
    LineStrPoints2(3) = -dCylHeight / 2 + dBaseConeHeight
    LineStrPoints2(4) = parInstrumentLength - dRadLowerCyl + dTopRBaseCone
    LineStrPoints2(5) = LineStrPoints2(2)
    
    LineStrPoints2(6) = -dCylHeight / 2 + dConeheight
    LineStrPoints2(7) = parInstrumentLength - dRadLowerCyl + dTopRUpperCone
    LineStrPoints2(8) = LineStrPoints2(2)
    
    LineStrPoints2(9) = LineStrPoints2(6)
    LineStrPoints2(10) = parInstrumentLength - dRadLowerCyl
    LineStrPoints2(11) = LineStrPoints2(2)
    
    LineStrPoints2(12) = LineStrPoints2(0)
    LineStrPoints2(13) = LineStrPoints2(10)
    LineStrPoints2(14) = LineStrPoints2(2)
        
    LineStrPoints2(15) = LineStrPoints2(0)
    LineStrPoints2(16) = LineStrPoints2(1)
    LineStrPoints2(17) = LineStrPoints2(2)
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, LineStrPoints2)
    
    axisVect.Set 1, 0, 0
    cenPoint.Set 0, parInstrumentLength - dRadLowerCyl - 0.00001, 0
    Set objLowerConicalLid = PlaceRevolution(m_OutputColl, oLineString, axisVect, cenPoint, _
                                2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLowerConicalLid
    Set objLowerConicalLid = Nothing
    Set geomFactory = Nothing
    Set cenPoint = Nothing
    Set axisVect = Nothing
    Set oLineString = Nothing

' Place Nozzle 1 (Inlet Port)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim dNozzleLength As Double
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle1   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
     
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
    Set objNozzle1 = Nothing

' Place Nozzle 2 (Outlet Port)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim objNozzle2   As GSCADNozzleEntities.IJDNozzle
    oPlacePoint.Set 0, -parFace2toCenter - sptOffset + depth, 0
    oDir.Set 0, -1, 0
    dNozzleLength = parFace2toCenter
    If CmpDblEqual(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    
    Set objNozzle2 = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
    Set objNozzle2 = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub


