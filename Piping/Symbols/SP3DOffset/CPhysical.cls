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
'   Copyright (c) 2008 Intergraph Corporation. All rights reserved.
'
'   File:           CPhysical.cls
'   Author:         VRK
'   Creation Date:  Wednesday, Jan 02 2008
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   02.Jan.2008     VRK     CR-131510:Create offset symbol to support the following options:
'                                      i.Offset, 45 degree
'                                      ii.Mechanical joint offset
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private Const MODULE = "Physical:"    'Used for error messages

Private Sub Class_Initialize()

    PI = 4 * Atn(1)

End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam1 As Double
    Dim flangeThick As Double
    Dim cptOffset As Double
    Dim flangeDiam1 As Double
    Dim depth As Double
    Dim pipeDiam2 As Double
    Dim flangeThick2 As Double
    Dim cptOffset2 As Double
    Dim flangeDiam2 As Double
    Dim depth2 As Double

    Dim iOutput As Integer

    Dim parSeattoSeat As Double
    Dim parOffset As Double
    Dim parTangentLength2 As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parSeattoSeat = arrayOfInputs(2)
    'parOffset = arrayOfInputs(3)
    'parTangentLength2 = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    'parAngle= arrayOfInputs(6)

    iOutput = 0
    parAngle = PI / 4    '45 Degrees

    Dim dElbowRadius As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dTangentLength1 As Double
    Dim dTangentLength2 As Double
    Dim dFacetoFace As Double

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, _
                       cptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                       cptOffset2, depth2

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    ' 45o offset, specified by seat-to-seat dimension, vertical offset, and tangent length 2
    Select Case lPartDataBasis
    Case Is <= 1, 586
        parSeattoSeat = arrayOfInputs(2)
        parOffset = arrayOfInputs(3)
        parTangentLength2 = arrayOfInputs(4)
        dFacetoFace = parSeattoSeat + (depth - cptOffset) + (depth2 - cptOffset2)
        dFace1toCenter = 0.5 * dFacetoFace
        dFace2toCenter = 0.5 * dFacetoFace

    Case Else
        GoTo ErrorLabel    'Invalid specification.

    End Select
    
    'Assuming tangent length 1 is 0.15 times Face to Face.
    dTangentLength1 = 0.15 * dFacetoFace
    If CmpDblLessThanOrEqualTo(parTangentLength2, 0) Then
        dTangentLength2 = dTangentLength1
    Else
        dTangentLength2 = parTangentLength2
    End If

    'Checking for Zero Tangent Length
    If CmpDblLessThan(dTangentLength1, flangeThick) Then dTangentLength1 = flangeThick
    If CmpDblLessThan(dTangentLength2, flangeThick2) Then dTangentLength2 = flangeThick2
    
    dElbowRadius = ((dFacetoFace - dTangentLength1 - dTangentLength2) - parOffset) / (2 * Sqr(2) - 2)

    'Create Body
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim Surfset As IngrGeom3D.IJElements
    Dim oTraceStr As IngrGeom3D.ComplexString3d
    Dim oCircle As IngrGeom3D.Circle3d
    Dim oCollection As Collection
    Dim iCount As Integer
    Dim oLine As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition
    Dim oNormal As AutoMath.DVector
    Dim stnorm() As Double
    Dim ednorm() As Double
    Dim ObjSurface As Object
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set oCenter = New AutoMath.DPosition
    Set oTraceStr = New ComplexString3d
    Set oCollection = New Collection
    Set oLine = New Line3d
    Set oArc = New Arc3d
    Set oNormal = New DVector

    oCenter.Set -dFace1toCenter + dTangentLength1, 0.5 * parOffset, 0
    oNormal.Set -1, 0, 0
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                    oCenter.x, oCenter.y, oCenter.z, _
                                                                    oNormal.x, oNormal.y, oNormal.z, pipeDiam2 / 2)

    oStPoint.Set -dFace1toCenter + dTangentLength1, 0.5 * parOffset, 0
    oCenter.Set -dFace1toCenter + dTangentLength1, -(dElbowRadius - 0.5 * parOffset), 0
    oEnPoint.Set -dFace1toCenter + dTangentLength1 + dElbowRadius * Cos(parAngle), 0.5 * parOffset - (dElbowRadius - dElbowRadius * Sin(parAngle)), 0
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                          oCenter.x, oCenter.y, oCenter.z, _
                                                          oStPoint.x, oStPoint.y, oStPoint.z, _
                                                          oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc

    oStPoint.Set -dFace1toCenter + dTangentLength1 + dElbowRadius * Cos(parAngle), 0.5 * parOffset - (dElbowRadius - dElbowRadius * Sin(parAngle)), 0
    oEnPoint.Set dFace2toCenter - dTangentLength2 - dElbowRadius * Cos(parAngle), -0.5 * parOffset + (dElbowRadius - dElbowRadius * Sin(parAngle)), 0
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine

    oStPoint.Set dFace2toCenter - dTangentLength2 - dElbowRadius * Cos(parAngle), -0.5 * parOffset + (dElbowRadius - dElbowRadius * Sin(parAngle)), 0
    oCenter.Set dFace2toCenter - dTangentLength2, dElbowRadius - 0.5 * parOffset, 0
    oEnPoint.Set dFace2toCenter - dTangentLength2, -0.5 * parOffset, 0
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                          oCenter.x, oCenter.y, oCenter.z, _
                                                          oStPoint.x, oStPoint.y, oStPoint.z, _
                                                          oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc
   
    oStPoint.Set -dFace1toCenter + dTangentLength1, 0.5 * parOffset, 0
    Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
    Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                  m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                  CircularCorner, 0, stnorm, ednorm, False)

    For Each ObjSurface In Surfset
        m_OutputColl.AddOutput "Object", ObjSurface
    Next ObjSurface

    'Remove References
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    For iCount = 1 To Surfset.Count
        Surfset.Remove 1
    Next iCount
    Set oGeomFactory = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenter = Nothing
    Set oTraceStr = Nothing
    Set oCollection = Nothing
    Set oLine = Nothing
    Set oArc = Nothing
    Set oNormal = Nothing
    Set oCircle = Nothing
    Set ObjSurface = Nothing
    
    'Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim dPorttoOrigin As Double
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    dPorttoOrigin = dFace1toCenter + cptOffset - depth
    oPlacePoint.Set -dPorttoOrigin, 0.5 * parOffset, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dTangentLength1)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    ' Place Nozzle 2
    dPorttoOrigin = dFace2toCenter + cptOffset2 - depth2
    oPlacePoint.Set dPorttoOrigin, -0.5 * parOffset, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dTangentLength2)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oGeomFactory = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
