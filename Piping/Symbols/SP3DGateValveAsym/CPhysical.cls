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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'
'  Project:
'
'  File:
'
'  Description:
'
'  Author:
'
'  History:
'       SS Mar/08/2000
'           used JObjectCollection instead of UntransactedMiddleElems
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  1.NOV.2007      RRK  CR-123952 Updated the symbol to use retrive parameters function in basgeom3d.bas
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Run:" 'Used for error messages

Private gscadElem           As IJDObject

Private geomFactory         As IngrGeom3D.GeometryFactory
Private objBspline          As IngrGeom3D.BSplineCurve3d
Private objComplexString    As IngrGeom3D.ComplexString3d

Private bonnetDiam          As Double
Private discDiam            As Double
Private bodyVDiam1          As Double
Private bodyVDiam2          As Double
Private bonnetWidth         As Double
Private actWidth            As Double
Private actDiam             As Double
Private stemDiam            As Double
Private currentHeight       As Double

Private PI                  As Double

Private Sub Class_Initialize()

    PI = 4 * Atn(1)
    
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPipePort       As GSCADNozzleEntities.IJDPipePort
    Dim nozzleIndex     As Integer
    Dim circleCenter    As AutoMath.DPosition
    Dim circleNormal    As AutoMath.DVector
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim projVec         As AutoMath.DVector
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim ValveL          As Double
    Dim NozzleFactory   As GSCADNozzleEntities.NozzleFactory
    Dim oNozzle         As GSCADNozzleEntities.IJDNozzle
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim cptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim bodyDiam        As Double
    Dim valveHeight     As Double
    Dim flangeThick     As Double
    Dim cptOffset            As Double
    Dim flangeDiam      As Double
'    Dim discWidth       As Double
    Dim stemFlWidth     As Double
    Dim stemFlDiam      As Double
    Dim iOutput         As Double
    Dim depth         As Double
    
    Dim face1c          As Double
    Dim face2c          As Double
    Dim faceToFace      As Double
    
    faceToFace = face1c + face2c
    
    Dim ValveL1c        As Double
    Dim ValveL2c        As Double

    
    Set oPartFclt = arrayOfInputs(1)
    Set geomFactory = New IngrGeom3D.GeometryFactory

' Use nozzle 1 as pipe port for flange and pipe parameters
    RetrieveParameters 1, oPartFclt, m_OutputColl, bodyDiam, flangeThick, flangeDiam, cptOffset, depth
    
    face1c = arrayOfInputs(2)
    face2c = arrayOfInputs(3)
    valveHeight = arrayOfInputs(4)
    actDiam = arrayOfInputs(5)
    
'    ValveL = discWidth - (2 * flangeThick)
    ValveL1c = face1c - flangeThick
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, cptOffset2, depth2
    ValveL2c = face2c - flangeThick2
    ValveL = ValveL1c + ValveL2c
    
'    Set oPipePort = Nothing
'   keep oNozzle for nozzle1 cretation later

    If CmpDblEqual(flangeDiam, 0) Then
        discDiam = bodyDiam
        stemFlWidth = bodyDiam * 0.17
        stemFlDiam = discDiam * 1.5
        bonnetDiam = discDiam * 1.3
    Else
        discDiam = (bodyDiam + flangeDiam) / 2.2
        stemFlWidth = flangeThick
        stemFlDiam = flangeDiam
        bonnetDiam = discDiam
    End If
    
    bonnetWidth = valveHeight * 0.15
    stemDiam = bodyDiam * 0.4
    actWidth = bodyDiam * 0.25
    
'----------------------- Horisontal Body
' Construction of Horisontal Body Cylinder
    
    iOutput = 1
    Set circleCenter = New AutoMath.DPosition
'    circleCenter.Set -(ValveL / 2), 0, 0
    circleCenter.Set -ValveL1c, 0, 0
    Set circleNormal = New AutoMath.DVector
    circleNormal.Set 1, 0, 0
    
' Construct a circle that will be used to project the disc
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        discDiam * 0.5)
    
    Set projVec = New AutoMath.DVector
    projVec.Set 1, 0, 0

' Project the disc of body
    Set objProjection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                        objCircle, _
                                                        projVec.x, projVec.y, projVec.z, _
                                                        ValveL, False)
    
' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objProjection
    Set objProjection = Nothing
    iOutput = iOutput + 1
    
'Delete the circle that was created for the projection
    Set gscadElem = objCircle
    Set objCircle = Nothing
    gscadElem.Remove
    
'----------------------- Body Vertical Bottom part (Revolution surface)
    currentHeight = 0
    CreateBodyVB m_OutputColl, arrayOfOutputs(iOutput), currentHeight
    iOutput = iOutput + 1

'----------------------- Body Vertical part (Projection)
    CreateBodyV m_OutputColl, arrayOfOutputs(iOutput), currentHeight
    iOutput = iOutput + 1
    currentHeight = currentHeight + 1.5 * bonnetDiam
    'currentHeight = currentHeight + 1.5 * discDiam

'----------------------- Body Vertical Flange
' Construction of Vertical Flange

    Dim transUp         As New AutoMath.DT4x4
    Dim transVec        As New AutoMath.DVector

' Use existing ComplexString object, just move it up (on Y axis)
    transUp.[Scale] 1.2
    objComplexString.Transform transUp
    transVec.Set 0, currentHeight, 0
    Set transUp = New AutoMath.DT4x4
    transUp.Translate transVec
    objComplexString.Transform transUp
' Project the Flange
    projVec.Set 0, 1, 0
    Set objProjection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                            objComplexString, _
                                                            projVec.x, projVec.y, projVec.z, _
                                                            stemFlWidth, True)

' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objProjection
    iOutput = iOutput + 1
    Set objProjection = Nothing
    currentHeight = currentHeight + stemFlWidth



'----------------------- bonnet flange
' Construction of bonnet flange Cylinder, starting where the vertical body left off

' Use existing ComplexString object, just move it up (on Y axis)
    Set transUp = New AutoMath.DT4x4
    transVec.Set 0, stemFlWidth + 0.002, 0
    transUp.Translate transVec
    objComplexString.Transform transUp

    Set transUp = Nothing
    Set transVec = Nothing

' Project the bonnet flange
    Set objProjection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                    objComplexString, _
                                                    projVec.x, projVec.y, projVec.z, _
                                                    stemFlWidth, True)

' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objProjection
    Set objProjection = Nothing
    iOutput = iOutput + 1

    currentHeight = currentHeight + stemFlWidth + 0.002

'Delete the objComplexString that was created for the projection
    Set gscadElem = objComplexString
    Set objComplexString = Nothing
    gscadElem.Remove
    
'----------------------- bonnet
    CreateBonnet m_OutputColl, arrayOfOutputs(iOutput), currentHeight
    currentHeight = currentHeight + bonnetDiam / 2
    'currentHeight = currentHeight + discDiam / 2
    iOutput = iOutput + 1
    
'----------------------- stem body
' Construction of stem Cylinder, starting where the top of the disc
    circleCenter.Set 0, currentHeight, 0
    circleNormal.Set 0, 1, 0
    
    ' Construct a circle that will be used to project the stem
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        stemDiam * 0.5)

    ' Project the Stem body
    projVec.Set 0, 1, 0
    Set objProjection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                        objCircle, _
                        projVec.x, projVec.y, projVec.z, _
                        valveHeight - currentHeight, True)
    
    ' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objProjection
    Set objProjection = Nothing
    currentHeight = valveHeight - 1.5 * actWidth
    iOutput = iOutput + 1
    
'Delete the circle that was created for the projection
    Set gscadElem = objCircle
    Set objCircle = Nothing
    gscadElem.Remove
    
'----------------------- actuator body
' Creation of actuator torus
    CreateActuator m_OutputColl, arrayOfOutputs(iOutput), currentHeight
    iOutput = iOutput + 1

' Construction of actuator spike1 (cylinder)
    circleCenter.Set -actDiam / 2, currentHeight, 0
    circleNormal.Set 1, 0, 0
' Construct a circle that will be used to project the spike
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        actWidth * 0.25)
' Project the Spike body
    projVec.Set 1, 0, 0
    Set objProjection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                        objCircle, _
                        projVec.x, projVec.y, projVec.z, _
                        actDiam, True)
' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objProjection
    iOutput = iOutput + 1

' Construction of actuator spike2
    circleCenter.Set 0, currentHeight, -actDiam / 2
    circleNormal.Set 0, 0, 1
' Construct a circle that will be used to project the spike
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        actWidth * 0.25)
' Project the Spike body
    projVec.Set 0, 0, 1
    Set objProjection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                        objCircle, _
                        projVec.x, projVec.y, projVec.z, _
                        actDiam, True)
    
    ' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objProjection
    Set objProjection = Nothing
    iOutput = iOutput + 1
    
'Delete the circle that was created for the projection
    Set gscadElem = objCircle
    Set objCircle = Nothing
    gscadElem.Remove
    
''===========================
''Construction of nozzle1
''===========================
    Dim pos                 As New AutoMath.DPosition
    Dim dir                 As New AutoMath.DVector
    Dim xctrpt              As Double
    dir.Set -1, 0, 0
    'Position of the nozzle should be the connect point of the nozzle
    xctrpt = face1c - (depth - cptOffset)
    pos.Set -xctrpt, 0, 0

    Set oNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, dir, pos)
' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    iOutput = iOutput + 1
    Set oNozzle = Nothing
    
'======================
' Creation of nozzle2
'======================
    dir.Set 1, 0, 0
    xctrpt = face2c - (depth2 - cptOffset2)
    pos.Set xctrpt, 0, 0
    
    Set oNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, dir, pos)
' Set the output
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    iOutput = iOutput + 1
    
    Set oNozzle = Nothing
    Set geomFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing
    Set oPartFclt = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
'    MsgBox "Error encountered while computing default representation", vbExclamation, "BFYLP definition"
    Resume Next
    
End Sub

Private Sub CreateBodyVB(ByVal m_OutputColl As Object, outname As String, currHeight As Double)
    
' Construction of Bottom Part of vertical body of Gate Valve
' This is revolution of BSpline curve

    Const METHOD = "CreateBodyVB"
    On Error GoTo ErrorHandler
        
    Dim objRevolution   As IngrGeom3D.Revolution3d
    
    Set objBspline = getMyBspline(currHeight, bonnetDiam)
    'Set objBspline = getMyBspline(currHeight, discDiam)
    Dim axisVector      As New AutoMath.DVector
    Dim centerPoint     As New AutoMath.DPosition
    
    axisVector.Set -1, 0, 0
    centerPoint.Set 0, currHeight, 0
    Set objRevolution = geomFactory.Revolutions3d.CreateByCurve( _
                                                    m_OutputColl.ResourceManager, _
                                                    objBspline, _
                                                    axisVector.x, axisVector.y, axisVector.z, _
                                                    centerPoint.x, centerPoint.y, centerPoint.z, _
                                                    PI, False)

' Set the output
    m_OutputColl.AddOutput outname, objRevolution
    Set objRevolution = Nothing

' Release objects
    Set gscadElem = objBspline
    Set objBspline = Nothing
    gscadElem.Remove
    Exit Sub
    
ErrorHandler:
  ReportUnanticipatedError MODULE, METHOD

    
End Sub

Private Sub CreateBodyV(ByVal m_OutputColl As Object, outname As String, currHeight As Double)

' Construction of projection based on complex string from 2 bspline curves
    
    Const METHOD = "CreateBodyV"
    On Error GoTo ErrorHandler
        
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim objBspline2     As IngrGeom3D.BSplineCurve3d
    Dim curvesColl      As IJElements
    
    Set curvesColl = New JObjectCollection
    Set objBspline = getMyBspline(currHeight, bonnetDiam)
    'Set objBspline = getMyBspline(currHeight, discDiam)
    curvesColl.Add objBspline
    
    Dim transUp     As New AutoMath.DT4x4
    Dim rotateVec   As New AutoMath.DVector
    
' Use existing BSpline object, just move it up (on Y axis)
    rotateVec.Set 0, 1, 0
    transUp.Rotate PI, rotateVec
    Set objBspline2 = getMyBspline(currHeight, bonnetDiam)
    'Set objBspline2 = getMyBspline(currHeight, discDiam)
    objBspline2.Transform transUp
    
    curvesColl.Add objBspline2
'    Set objBspline2 = Nothing
    
    Set objComplexString = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, _
                                            curvesColl)
                                            
    Dim projVec         As New AutoMath.DVector
    projVec.Set 0, 1, 0
    Set objProjection = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                        objComplexString, _
                        projVec.x, projVec.y, projVec.z, _
                        1.5 * bonnetDiam, False)
                        '1.5 * discDiam, False)
    
' Set the output
    m_OutputColl.AddOutput outname, objProjection
    Set objProjection = Nothing
    Set projVec = Nothing

' Release objects

    Set gscadElem = objBspline
    Set objBspline = Nothing
    gscadElem.Remove
    Set gscadElem = objBspline2
    Set objBspline2 = Nothing
    gscadElem.Remove
    Dim oCurve As Object
    For Each oCurve In curvesColl
        Set gscadElem = oCurve
        Set oCurve = Nothing
        gscadElem.Remove
    Next oCurve
    curvesColl.Clear
'    Set gscadElem = objComplexString
'    Set objComplexString = Nothing
'    gscadElem.Remove
    
    Exit Sub
    
ErrorHandler:
  ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub CreateBonnet(ByVal m_OutputColl As Object, outname As String, currHeight As Double)
    
' Construction of bonnet based on revolution of bspline curve
    Const METHOD = "CreateBonnet"
    On Error GoTo ErrorHandler
        
    Dim objRevolution   As IngrGeom3D.Revolution3d

' Use existing BSpline object, just move it up (on Y axis)
    Set objBspline = getMyBspline(currHeight, bonnetDiam)
    'Set objBspline = getMyBspline(currHeight, discDiam)
    
    Dim axisVector      As New AutoMath.DVector
    Dim centerPoint     As New AutoMath.DPosition
    
    axisVector.Set 1, 0, 0
    centerPoint.Set 0, currHeight, 0
    Set objRevolution = geomFactory.Revolutions3d.CreateByCurve( _
                                                    m_OutputColl.ResourceManager, _
                                                    objBspline, _
                                                    axisVector.x, axisVector.y, axisVector.z, _
                                                    centerPoint.x, centerPoint.y, centerPoint.z, _
                                                    PI, False)

' Set the output
    m_OutputColl.AddOutput outname, objRevolution
    Set objRevolution = Nothing

' Release objects
    Set gscadElem = objBspline
    Set objBspline = Nothing
    gscadElem.Remove
    Exit Sub
    
ErrorHandler:
  ReportUnanticipatedError MODULE, METHOD

    
End Sub

Private Sub CreateActuator(ByVal m_OutputColl As Object, outname As String, currentWidth As Double)

' Construction of actuator torus
    
    Dim centerPoint     As New AutoMath.DPosition
    Dim vNormal         As New AutoMath.DVector
    Dim originDir       As New AutoMath.DVector
    Dim minorRadius     As Double
    
    Dim objTorus     As IngrGeom3D.Torus3d
    
    centerPoint.Set 0, currentWidth, 0
    vNormal.Set 0, 1, 0
    originDir.Set 1, 0, 0
    minorRadius = actWidth / 2
    Set objTorus = geomFactory.Tori3d.CreateByAxisMajorMinorRadius(m_OutputColl.ResourceManager, _
                                        centerPoint.x, centerPoint.y, centerPoint.z, _
                                        vNormal.x, vNormal.y, vNormal.z, _
                                        originDir.x, originDir.y, originDir.z, _
                                        actDiam / 2, minorRadius, False)
    
    ' Set the output
    m_OutputColl.AddOutput outname, objTorus

' Release objects
    Set objTorus = Nothing

End Sub

Private Function getMyBspline(lHeight As Double, lDiam As Double) As IngrGeom3D.BSplineCurve3d
    
    Const METHOD = "getMyBspline:"
    On Error GoTo ErrorHandler
        
    Dim basePoints(6)   As AutoMath.DPosition       ' nPoints
    Dim dblPnts(18)     As Double                   ' nPoints * 3
    Dim iCount          As Integer
    Dim curIndex        As Integer
    Dim nPoints         As Long
    Dim nOrder          As Long
    Dim nKnots          As Double
    Dim discRadius      As Double
    Dim oBspline        As IngrGeom3D.BSplineCurve3d
    
    nPoints = 6
    For iCount = 0 To nPoints
        Set basePoints(iCount) = New AutoMath.DPosition
    Next iCount
    discRadius = lDiam / 2
    basePoints(0).Set -0.75 * discRadius, lHeight, 0
    basePoints(1).Set -0.75 * discRadius, lHeight, -discRadius
    basePoints(2).Set -0.5 * discRadius, lHeight, -1.25 * discRadius
    basePoints(3).Set 0.5 * discRadius, lHeight, -1.25 * discRadius
    basePoints(4).Set 0.75 * discRadius, lHeight, -discRadius
    basePoints(5).Set 0.75 * discRadius, lHeight, 0
    
    nOrder = 4
    nKnots = nPoints + nOrder
    Dim fKnots(10) As Double     ' nKnots
    
    For iCount = 0 To nPoints - 1
        curIndex = iCount * 3
        dblPnts(curIndex) = basePoints(iCount).x
        dblPnts(curIndex + 1) = basePoints(iCount).y
        dblPnts(curIndex + 2) = basePoints(iCount).z
    Next iCount
    
    Dim fWeights() As Double       ' nPoints
    ReDim fWeights(nPoints) As Double
    
    For iCount = 0 To nPoints - 1
        fWeights(iCount) = 1
    Next
        
    For iCount = 0 To nOrder - 1
        fKnots(iCount) = 0#
        fKnots(nKnots - iCount - 1) = 1#
    Next
    For iCount = nOrder To nPoints - 1
        fKnots(iCount) = (iCount - nOrder + 1) / (nPoints + 1 - nOrder)
    Next
    
' Create and keep bspline object for reutilizing in bodyV and Bonnet
    Set oBspline = geomFactory.BSplineCurves3d.CreateByParameters( _
                                                    Nothing, _
                                                    nOrder, nPoints, _
                                                    dblPnts, fWeights, _
                                                    fKnots, False)
    Set getMyBspline = oBspline

    Exit Function
    
ErrorHandler:
  ReportUnanticipatedError MODULE, METHOD

End Function

