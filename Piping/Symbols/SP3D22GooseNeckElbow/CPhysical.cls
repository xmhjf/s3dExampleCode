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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          RRK
'   Creation Date:  Thursday, July 12 2007
'   Description:
'   This symbol is prepared for Victaulic 22 1/2 Degree Goose Neck Elbow as per CR-123576

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   7.Nov.2007      RRK          TR-128456: Provided checks for flanged ends
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit
Private PI           As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub



Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    
    Dim dPort1X As Double
          
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim csObj As IngrGeom3D.IJElements
    Dim trObj As IngrGeom3D.IJElements
    Dim Surfset   As IngrGeom3D.IJElements
    Dim WCaps As Integer
    
    Set csObj = New JObjectCollection
    Set trObj = New JObjectCollection

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0
    
    Dim dAngle As Double
    dAngle = PI / 8
    
 ' Insert your code for output 1(Body)
  'Origin is assumed to be at the intersection point of two axes(which is port 2) passing through centers of Port 1 and Port 2
  'Distance of port 1 from origin along X axis is specified as follows
    dPort1X = parFacetoFace
    
   'Thickness of grooved portions are assumed to be 0.15 times Face to Face
    Dim dGrooveThickness As Double
    dGrooveThickness = 0.15 * parFacetoFace

    'Setting thickness for flanged ends as flange thickness
    If CmpDblGreaterthan(flangeThick, LINEAR_TOLERANCE) Then
        dGrooveThickness = flangeThick
    End If
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth

    Dim objCircle1   As IngrGeom3D.Circle3d
    Dim objCircle2   As IngrGeom3D.Circle3d

    'Creating Circle near Port 1
    'X Co-ordinate of Center of Circle 1
    Dim dCircle1CentX As Double
    dCircle1CentX = -dPort1X + dGrooveThickness
    
    Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         dCircle1CentX, 0, 0, _
                        1, 0, 0, _
                        pipeDiam / 2)
                            
    'Creating Circle near Port 2
    'Coordinates of center of circle 2
    Dim dCircle2CentX As Double
    Dim dCircle2CentY As Double

    dCircle2CentX = (dGrooveThickness * Cos(dAngle))
    dCircle2CentY = (dGrooveThickness * Sin(dAngle))
    
    Set objCircle2 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -dCircle2CentX, -dCircle2CentY, 0, _
                        Cos(dAngle), Sin(dAngle), 0, _
                        pipeDiam / 2)
    
    'Adding two circles to Collection of cross sections
     csObj.Add objCircle1
     csObj.Add objCircle2
     
     Set objCircle1 = Nothing
     Set objCircle2 = Nothing
     
    'Creating Traces for skinning between two cross sections
     'Creating Trace 1
     Dim points1(1 To 5)   As New AutoMath.DPosition
     Dim objBspline1  As IngrGeom3D.BSplineCurve3d
     Dim objBspline2  As IngrGeom3D.BSplineCurve3d
     
     'Assumption for the distance (along X axis) between end points of trace 1
     Dim dX1 As Double
     dX1 = (-dCircle2CentX + (pipeDiam * Sin(dAngle) / 2)) - dCircle1CentX
     
    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    Dim dNPD As Double
    Set oCollection = oPartFclt.GetNozzles()
    
    'Retrieval Nominal pipe diameter
    Set oPipePort = oCollection.Item(1)
    dNPD = oPipePort.NPD
    If oPipePort.NPDUnitType = "mm" Then
        dNPD = dNPD / 1000
    ElseIf oPipePort.NPDUnitType = "in" Then
        dNPD = dNPD * 0.0254
    End If
    
    Set oPipePort = Nothing
    Set oCollection = Nothing
    
    'Based on the ratio between Face to Face and NPD, the dY(position of outermost point
    'of trace curve 1) over pipediam/2 along Y direction is assumed as follows
    Dim dY1 As Double
    If CmpDblGreaterthanOrEqualTo(parFacetoFace / dNPD, 1.5) Then
       dY1 = 0.3 * pipeDiam
    Else
       dY1 = 0.1 * pipeDiam
    End If
    
    'Defining Points for trace curve 1
    points1(1).Set dCircle1CentX, _
               -pipeDiam / 2, 0
    points1(2).Set dCircle1CentX + (0.25 * dX1), -(pipeDiam / 2) - 0.5 * dY1, 0
    points1(3).Set dCircle1CentX + (0.5 * dX1), -(pipeDiam / 2) - dY1, 0
    points1(4).Set dCircle1CentX + (0.75 * dX1), -(pipeDiam / 2) - 0.5 * dY1, 0
    points1(5).Set -dCircle2CentX + (pipeDiam * Sin(dAngle) / 2), _
                                        -(pipeDiam * Cos(dAngle) / 2) - dCircle2CentY, 0
    Set objBspline1 = PlaceTrBspline(4, points1)
    
    trObj.Add objBspline1
    Set objBspline1 = Nothing
    
    'Creating Trace 2
    Dim dX2 As Double
    Dim dY2 As Double
    'Assumption for the distance (along X axis) between end points of trace 2
    dX2 = (-dCircle2CentX - (pipeDiam * Sin(dAngle) / 2)) - dCircle1CentX
    
    'dY(position of outermost point of trace curve 2) over pipediam/2
    'along Y direction is assumed as follows
    dY2 = 0.05 * pipeDiam
    
    'Defining Points for trace curve 2
    Dim points2(1 To 4)   As New AutoMath.DPosition
    points2(1).Set dCircle1CentX, _
                   pipeDiam / 2, 0
    points2(2).Set dCircle1CentX + (0.25 * dX2), (pipeDiam / 2) + 0.25 * dY2, 0
    
    points2(3).Set dCircle1CentX + (0.5 * dX2), _
                    (pipeDiam / 2) - 0.15 * dY2, 0
    points2(4).Set -dCircle2CentX - (pipeDiam * Sin(dAngle) / 2), _
                   (pipeDiam * Cos(dAngle) / 2) - dCircle2CentY, 0
    Set objBspline2 = PlaceTrBspline(4, points2)
    
    trObj.Add objBspline2
    Set objBspline2 = Nothing
    
    Dim oLineString As IngrGeom3D.LineString3d
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Set Surfset = geomFactory.GeometryServices.CreateBySkinning(m_OutputColl.ResourceManager, _
        trObj, csObj, WCaps)
        
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), Surfset(1)
    Set Surfset = Nothing
    Set trObj = Nothing
    Set csObj = Nothing
    
' Place Nozzle 1
        
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -dPort1X - cptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dGrooveThickness)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth

    oPlacePoint.Set (cptOffset - depth) * Cos(dAngle), (cptOffset - depth) * Sin(dAngle), 0
    oDir.Set Cos(dAngle), Sin(dAngle), 0

    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dGrooveThickness)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub


