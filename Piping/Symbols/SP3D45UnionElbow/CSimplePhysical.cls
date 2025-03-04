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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         SSP
'   Creation Date:  Wednesday, Aug 27 2003
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The Details were taken from PDS Piping Component Data - Reference Guide at Page no D-55
'    SN=F63 and MC=E45U. Symbol is created with 12 Outputs - Seven are Insulation aspect
'    outputs. Other are physical aspect outputs and corresponding function are as follows:
'      ObjUnion1 - 'PlaceProjection'
'      ObjElbow - 'PlaceRevolution'
'      ObjUnion2 - 'PlaceProjection' and
'      two objNozzle - 'CreateNozzleWithLength'.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim parUnionDiameter As Double
    Dim parTangentLength As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim axisVect    As New AutoMath.DVector
    Dim Angle As Double
    Dim ArcStPointToCenter  As Double
    Dim iCount As Integer
    
    Dim iOutput     As Double

    Dim ObjUnion1 As Object
    Dim ObjElbow As Object
    Dim ObjUnion2 As Object
    

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parUnionDiameter = arrayOfInputs(4)
    parTangentLength = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    iOutput = 0
    
' Insert your code for output 8(Union At Port1)
    Dim LineStrPoints(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d

    Angle = PI / 3
    stPoint.Set -parFace1toCenter + parTangentLength, 0, 0
'   First point is considered on positive Z-axis at an angle (PI / 6 + Angle) from positive Y-axis.
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = stPoint.x
        LineStrPoints(3 * iCount - 2) = stPoint.y + (parUnionDiameter / 2) * Cos(PI / 6 + iCount * Angle)
        LineStrPoints(3 * iCount - 1) = stPoint.z + (parUnionDiameter / 2) * Sin(PI / 6 + iCount * Angle)
    Next iCount

'   Refere to Note 19 - assumption: projection Length is 0.2 times parFace1toCenter
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    axisVect.Set -1, 0, 0
    Set ObjUnion1 = PlaceProjection(m_OutputColl, oLineString, axisVect, parFace1toCenter / 5, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjUnion1
    Set ObjUnion1 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Insert your code for output 9(Elbow)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim ElbowRadius As Double
    Angle = PI / 4
    ArcStPointToCenter = parFace1toCenter - parTangentLength
    ElbowRadius = ArcStPointToCenter / Tan(Angle / 2)
'   Construct a circle that will be used to create the revolution
    Dim objCircle   As IngrGeom3D.Circle3d
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -ArcStPointToCenter, 0, 0, _
                        1, 0, 0, _
                        pipeDiam / 2)
'   Revolve the circle about the Z axis
    Dim CP As New AutoMath.DPosition 'arc center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set -ArcStPointToCenter, ElbowRadius, 0
    CV.Set 0, 0, 1
    Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle, _
                    CV, CP, Angle, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
    Set ObjElbow = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCircle
    Set objCircle = Nothing
    sp3dElem1.Remove
    
' Insert your code for output 10(Union At Port2)
    Dim LineStrPoints2(0 To 20) As Double
'   From co-ordinate transformation get union2 hexagon points from LineStrPoints of union1
    Angle = 5 * PI / 4
    For iCount = 1 To 7
        LineStrPoints2(3 * iCount - 3) = LineStrPoints(3 * iCount - 3) * Cos(Angle) - _
                                    LineStrPoints(3 * iCount - 2) * Sin(Angle)
        LineStrPoints2(3 * iCount - 2) = LineStrPoints(3 * iCount - 3) * Sin(Angle) + _
                                    LineStrPoints(3 * iCount - 2) * Cos(Angle)
        LineStrPoints2(3 * iCount - 1) = LineStrPoints(3 * iCount - 1)
    Next iCount
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints2)
    axisVect.Set Cos(PI / 4), Sin(PI / 4), 0
    Set ObjUnion2 = PlaceProjection(m_OutputColl, oLineString, axisVect, parFace1toCenter / 5, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjUnion2
    Set ObjUnion2 = Nothing
    Set geomFactory = Nothing
    Set axisVect = Nothing
    Set sp3dElem1 = oLineString
    Set oLineString = Nothing
    sp3dElem1.Remove
    Set sp3dElem1 = Nothing

' Place Nozzle 1
    Dim dNozzleLength As Double
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim PorttoOrigin As Double
    PorttoOrigin = -parFace1toCenter - sptOffset + depth
    oPlacePoint.Set PorttoOrigin, 0, 0
    oDir.Set -1, 0, 0

    dNozzleLength = parTangentLength
    If dNozzleLength < flangeThick Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim TangentLength2 As Double
    TangentLength2 = parFace2toCenter - ElbowRadius * Tan(PI / 8)
    PorttoOrigin = parFace2toCenter + sptOffset - depth
    Angle = PI / 4
    oPlacePoint.Set PorttoOrigin * Cos(Angle), PorttoOrigin * Sin(Angle), 0
    oDir.Set Cos(Angle), Sin(Angle), 0

    dNozzleLength = TangentLength2
    If dNozzleLength < flangeThick Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
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


