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
'   Copyright (c) 2003-2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Tuesday, Feb 26 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'  This symbol is created using three outputs:  1. ObjReducer - Generated from 'CreateByCurves' function
'  by using two curves. Outputs 2 & 3 are CableTray Nozzles.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   30.May.2006     svsmylav               CR-91799 Added two tangent plate outputs.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
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
    Dim ObjReducer As Object
    
    Dim parFacetoFace As Double
    
'   Variables defined by svsmylav
    Dim oPort1 As New AutoMath.DPosition 'Port 1 center point
    Dim oPort2 As New AutoMath.DPosition 'Port 2 center point
    Dim parActualWidth     As Double
    Dim parActualWidth2    As Double
    Dim parActualDepth     As Double
    Dim Port1S(0 To 11)  As Double
    Dim Port2S(0 To 11)  As Double
    Dim HD                 As Double
    Dim HW                 As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)

    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    parFacetoFace = oTrayPart.Length
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
    
    iOutput = 0

'   Assumption: Length of the each Tangent Plate is taken to be parFacetoFace/3
    Dim dTangentPlateLength As Double
    
    dTangentPlateLength = parFacetoFace / 3
    
' Insert your code for output 1(TangentPlate1)
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    HD = parActualDepth / 2
    HW = parActualWidth / 2
    
    oPort1.Set -parFacetoFace / 2, 0, 0
    Port1S(0) = oPort1.x + dTangentPlateLength
    Port1S(1) = oPort1.y + HW
    Port1S(2) = oPort1.z + HD
    
    Port1S(3) = oPort1.x + dTangentPlateLength
    Port1S(4) = oPort1.y + HW
    Port1S(5) = oPort1.z - HD
    
    Port1S(6) = oPort1.x + dTangentPlateLength
    Port1S(7) = oPort1.y - HW
    Port1S(8) = oPort1.z - HD
    
    Port1S(9) = oPort1.x + dTangentPlateLength
    Port1S(10) = oPort1.y - HW
    Port1S(11) = oPort1.z + HD
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim oLineString1 As IngrGeom3D.LineString3d
    Set oLineString1 = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, Port1S)
    
    Dim ObjTangentPlate As Object
    Dim oProjVec As AutoMath.DVector
    Set oProjVec = New AutoMath.DVector
    
    oProjVec.Set -1, 0, 0
    Set ObjTangentPlate = PlaceProjection(m_OutputColl, oLineString1, oProjVec, dTangentPlateLength, False)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTangentPlate
    Set ObjTangentPlate = Nothing

' Insert your code for output 2(Reducer)
'   Prepare the Right Hand side line-string
    Call RetrieveCableTrayPortProperties(2, oPartFclt, parActualWidth2, parActualDepth)
    HD = parActualDepth / 2
    HW = parActualWidth2 / 2
    
    oPort2.Set parFacetoFace / 2, 0, 0
    Port2S(0) = oPort2.x - dTangentPlateLength
    Port2S(1) = oPort2.y + HW
    Port2S(2) = oPort2.z + HD
    
    Port2S(3) = oPort2.x - dTangentPlateLength
    Port2S(4) = oPort2.y + HW
    Port2S(5) = oPort2.z - HD
    
    Port2S(6) = oPort2.x - dTangentPlateLength
    Port2S(7) = oPort2.y - HW
    Port2S(8) = oPort2.z - HD
    
    Port2S(9) = oPort2.x - dTangentPlateLength
    Port2S(10) = oPort2.y - HW
    Port2S(11) = oPort2.z + HD
    Dim oLineString2 As IngrGeom3D.LineString3d
    Set oLineString2 = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, Port2S)

'   Create the Reducer (Ruled surface)
    Set ObjReducer = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                            oLineString1, oLineString2, False)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjReducer
    Set ObjReducer = Nothing
    
' Insert your code for output 3(TangentPlate2)
    oProjVec.Set 1, 0, 0
    Set ObjTangentPlate = PlaceProjection(m_OutputColl, oLineString2, oProjVec, dTangentPlateLength, False)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTangentPlate
    Set ObjTangentPlate = Nothing
    
' Place Nozzle 1
'   Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc
'   Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector

'   oPlacePoint.Set -parFacetoFace / 2, 0, 0
    oDir.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort1.x - dInsertionDepth * oDir.x, oPort1.y - dInsertionDepth * oDir.y, oPort1.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLocation, oDir, oRadialOrient, m_OutputColl)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
'   Set oPlacePoint = Nothing
    Set oPort1 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
' Place Nozzle 2
'   Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
'   oPlacePoint.Set parFacetoFace / 2, 0, 0
    oDir.Set 1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
'   Set oPlacePoint = Nothing
    Set oPort2 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing

' Remove curves
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = oLineString1
    Set oLineString1 = Nothing
    sp3dElem1.Remove
    
'   Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = oLineString2
    Set oLineString2 = Nothing
    sp3dElem1.Remove
    Set oPortLocation = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' This is a sample of placing cylinder and 2 nozzles
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Use nozzle 1 as pipe port for flange and pipe parameters
'''    Dim pipeDiam        As Double
'''    Dim flangeThick     As Double
'''    Dim sptOffset       As Double
'''    Dim flangeDiam      As Double
'''    Dim depth           As Double
'''    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
'''    Dim parOperatorWidth    As Double
'''    parOperatorWidth = arrayOfInputs(2)
'''    Dim stPoint   As New AutoMath.DPosition
'''    Dim enPoint   As New AutoMath.DPosition
'''    stPoint.Set -parOperatorWidth / 2 + flangeThick, 0, 0
'''    enPoint.Set parOperatorWidth / 2 - flangeThick, 0, 0
'''    Dim objCylinder As Object
'''
'''    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
''''   Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
'''    Set objCylinder = Nothing
'''
'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -parOperatorWidth / 2 - sptOffset + depth, 0, 0
'''    oDir.Set -1, 0, 0
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
''''   Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
'''    oPlacePoint.Set parOperatorWidth / 2 + sptOffset - depth, 0, 0
'''    oDir.Set 1, 0, 0
'''
'''    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
''''   Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
