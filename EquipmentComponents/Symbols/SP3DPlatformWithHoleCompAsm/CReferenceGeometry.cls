VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRefGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:14:56 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:          MS
'   Creation Date:  Friday, 2 May 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol consists of 12 Physical outputs. Of Which 10 plates, 2 saddles
'  The symbol consists of 5 Reference Geometry output which are 1 Reference plane,2 lines and 2 points.
'
'   Change History:
'   dd.mmm.yyyy  who                change description
'   -----------  -----              ------------------
'   09.Jul.2003  SymbolTeam(India)  Copyright Information, Header  is added.
'   08.Aug.2003  SSP             TR-45806  Error in PlatformWithHoleForPipe symbol:Added small clearance
'                                   between the platform inner /outer arc and circle boundary.
'   23.Aug.2003  SymbolTeam(India)  Updated Error Handling Code
'   23.Jun.2004   svsmylav(svsmylav)  TR-60704: Replaced 'm_OutputColl.ResourceManager' with "Nothing' for
'                                   creating transient Bottom Arc1(oTArc1) output to avoid extra
'                                   curves/lines in the symbol graphics.
'   29.Nov.2004  V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   16.May.2005  svsmylav           CR-76070: Input parameter parRadialPosition is replaced with parPlatformHoletoCenter.

'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private PI As Double
Private Const MODULE = "ReferenceGeometry:" 'Used for error messages
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
       PI = Atn(1) * 4
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double

    Dim parVesselDiameter As Double
    Dim parPlatformWidth As Double
    Dim parPlatformHeight As Double
    Dim parPlatformAngle As Double
    Dim parClearancefromVessel As Double
    Dim parInsulationThickness As Double
    Dim parOpeningDiameter As Double
    Dim parRotation As Double
    Dim parPlatformHoletoCenter As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parPlatformWidth = arrayOfInputs(3)
    parPlatformHeight = arrayOfInputs(4)
    parPlatformAngle = arrayOfInputs(5)
    parClearancefromVessel = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    parOpeningDiameter = arrayOfInputs(8)
    parRotation = arrayOfInputs(9)
    parPlatformHoletoCenter = arrayOfInputs(10)
    
    iOutput = 0

    'Case of 360 degrees.
    If (parPlatformAngle = 2 * PI) Then
        parPlatformAngle = 6.2831
    End If
    'Protect the Hole from going out of the Platform. Aleast it would touch the inner or outer edge of the platform
    Dim VessCentoHoleCen As Double
    VessCentoHoleCen = parVesselDiameter / 2 + parClearancefromVessel + parInsulationThickness + parOpeningDiameter / 2 + 0.001
    If (parPlatformHoletoCenter < VessCentoHoleCen) Then
        parPlatformHoletoCenter = VessCentoHoleCen
    End If
    Dim VessCentoPlatformOutedge As Double
    VessCentoPlatformOutedge = parVesselDiameter / 2 + parClearancefromVessel + parInsulationThickness + parPlatformWidth - parOpeningDiameter / 2 - 0.001
    If (parPlatformHoletoCenter > VessCentoPlatformOutedge) Then
        parPlatformHoletoCenter = VessCentoPlatformOutedge
    End If
    
    'If the parOpeningDiameter is given as zero place a very small hole which is almost not visible
    If (parOpeningDiameter = 0) Then
        parOpeningDiameter = 0.001
    End If

    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
' Insert your code for output 13(Reference plane on Tower)
    ' create non-persistent circle to use for creating default surface ---
    Dim objPlane As IngrGeom3D.Plane3d
    Dim objPlaneCircle As IngrGeom3D.Circle3d
    Set objPlaneCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0, 0, 0, _
                            0, 0, -1, _
                            0.1 * parVesselDiameter / 2)

'   create persistent default surface plane - the plane can mate ---
    Set objPlane = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objPlaneCircle)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
    Set objPlane = Nothing
    Set objPlaneCircle = Nothing
    
' Insert your code for output 16(Point on Tower)
    Dim oTowerPoint As IngrGeom3D.Point3d
    Set oTowerPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, 0, 0, 0)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTowerPoint
    Set oTowerPoint = Nothing
    
    'Creation of Bottom Platform
    'Create the lines and arcs of the top surface . use this lines and arcs to create the complexstring.
    'Use the complex string to project the thickness.
    Dim oComplexString     As New IngrGeom3D.ComplexString3d
    Dim oTLine1 As Object
    Dim oTArc2 As Object
    Dim oTLine2 As Object
    Dim oTArc1 As Object
    Dim ObjTopSurface As IngrGeom3D.Plane3d
    
    
    Dim oCenterPoint As New AutoMath.DPosition
    Dim arrPt(1 To 4)       As New AutoMath.DPosition
    Dim oNormalVector   As IJDVector
    Dim oCollection          As Collection
    
    
    Set oNormalVector = New AutoMath.DVector
    Set oCollection = New Collection
    oNormalVector.Set 0, 0, 1
    
    oCenterPoint.Set 0, 0, 0
    Dim dOrigintoPlatform As Double
    dOrigintoPlatform = parVesselDiameter / 2 + parClearancefromVessel + parInsulationThickness
    arrPt(1).Set (dOrigintoPlatform) * Sin(parPlatformAngle / 2), _
                        (dOrigintoPlatform) * Cos(parPlatformAngle / 2), _
                       0
    arrPt(2).Set (dOrigintoPlatform + parPlatformWidth) * Sin(parPlatformAngle / 2), _
                        (dOrigintoPlatform + parPlatformWidth) * Cos(parPlatformAngle / 2), _
                       0
    arrPt(3).Set -(dOrigintoPlatform + parPlatformWidth) * Sin(parPlatformAngle / 2), _
                        (dOrigintoPlatform + parPlatformWidth) * Cos(parPlatformAngle / 2), _
                       0
    arrPt(4).Set -(dOrigintoPlatform) * Sin(parPlatformAngle / 2), _
                        (dOrigintoPlatform) * Cos(parPlatformAngle / 2), _
                       0
    'Top Line 1
    Set oTLine1 = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, arrPt(1).x, arrPt(1).y, arrPt(1).z, arrPt(2).x, arrPt(2).y, arrPt(2).z)
    'Set the output line
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTLine1
    oCollection.Add oTLine1
    'Top Arc 2
    Set oTArc2 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(m_OutputColl.ResourceManager, _
                                                      oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                                     oNormalVector.x, oNormalVector.y, oNormalVector.z, _
                                                     arrPt(2).x, arrPt(2).y, arrPt(2).z, _
                                                    arrPt(3).x, arrPt(3).y, arrPt(3).z)
    'Set the output arc
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTArc2
    oCollection.Add oTArc2
    'Top Line 2
    Set oTLine2 = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, arrPt(3).x, arrPt(3).y, arrPt(3).z, arrPt(4).x, arrPt(4).y, arrPt(4).z)
    'Set the output line
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTLine2
    oCollection.Add oTLine2
    'Top Arc 1
    oNormalVector.Set 0, 0, -1
    Set oTArc1 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(m_OutputColl.ResourceManager, _
                                                  oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                                 oNormalVector.x, oNormalVector.y, oNormalVector.z, _
                                                 arrPt(4).x, arrPt(4).y, arrPt(4).z, _
                                                arrPt(1).x, arrPt(1).y, arrPt(1).z)
    'Set the output arc
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTArc1
    oCollection.Add oTArc1
    
    Set oComplexString = PlaceTrCString(arrPt(1), oCollection)
    
    'Create an infinite plane and appy the complex string as boundary
    Set ObjTopSurface = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, arrPt(1).x, arrPt(1).y, arrPt(1).z, 0, 0, 1)
    ObjTopSurface.AddBoundary oComplexString
    
    'Create the Hole for the Pipe
    Dim circlePosition As New AutoMath.DPosition
    Dim circleNormal As New AutoMath.DVector
    circlePosition.Set parPlatformHoletoCenter * Sin(parRotation), parPlatformHoletoCenter * Cos(parRotation), 0
    circleNormal.Set 0, 0, 1
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim oElements As IJElements
    Dim objCStr         As IngrGeom3D.ComplexString3d
    Set oElements = New JObjectCollection
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                    circlePosition.x, circlePosition.y, circlePosition.z, _
                    circleNormal.x, circleNormal.y, circleNormal.z, _
                    parOpeningDiameter / 2)
    
    oElements.Add objCircle
    
    Set objCStr = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements)
    ObjTopSurface.AddBoundary objCStr
    
    'Set the top plane output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopSurface
    Set ObjTopSurface = Nothing
    
    oElements.Clear
    objCStr.RemoveCurve True
    'Remove the complex string
    Dim Objcurves As IJDObject
    Set Objcurves = oComplexString
    Objcurves.Remove
    Set oComplexString = Nothing

    Dim count As Integer
    For count = 1 To oCollection.count
        oCollection.Remove 1
    Next count
    Set oCollection = Nothing
    
    'Create the bottom surface plane
    Dim oComplexString1 As New IngrGeom3D.ComplexString3d
    Dim obotCollection As New Collection
    Dim ObjBottSurface As IngrGeom3D.Plane3d
    
    For count = 1 To 4
        arrPt(count).z = -parPlatformHeight
    Next count
    oCenterPoint.Set 0, 0, -parPlatformHeight
    'Bottom Line 1
    Set oTLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPt(1).x, arrPt(1).y, arrPt(1).z, arrPt(2).x, arrPt(2).y, arrPt(2).z)
    obotCollection.Add oTLine1
    'Bottom Arc 2
    oNormalVector.Set 0, 0, 1
    Set oTArc2 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                  oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                                 oNormalVector.x, oNormalVector.y, oNormalVector.z, _
                                                 arrPt(2).x, arrPt(2).y, arrPt(2).z, _
                                                arrPt(3).x, arrPt(3).y, arrPt(3).z)
    obotCollection.Add oTArc2
    'Bottom Line 2
    Set oTLine2 = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPt(3).x, arrPt(3).y, arrPt(3).z, arrPt(4).x, arrPt(4).y, arrPt(4).z)
    obotCollection.Add oTLine2
    'Bottom Arc 1
    oNormalVector.Set 0, 0, -1
    Set oTArc1 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                              oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                             oNormalVector.x, oNormalVector.y, oNormalVector.z, _
                                             arrPt(4).x, arrPt(4).y, arrPt(4).z, _
                                            arrPt(1).x, arrPt(1).y, arrPt(1).z)
    obotCollection.Add oTArc1
    
    Set oComplexString1 = PlaceTrCString(arrPt(1), obotCollection)
    
    'Create an infinite plane and appy the complex string as boundary
    Set ObjBottSurface = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, arrPt(1).x, arrPt(1).y, arrPt(1).z, 0, 0, -1)
    ObjBottSurface.AddBoundary oComplexString1
    
    
    'Create the Hole for the Pipe
    circlePosition.Set parPlatformHoletoCenter * Sin(parRotation), parPlatformHoletoCenter * Cos(parRotation), -parPlatformHeight
    circleNormal.Set 0, 0, 1
    Dim oElements2 As IJElements
    Dim objCStr2         As IngrGeom3D.ComplexString3d
    Set oElements2 = New JObjectCollection
 

    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circlePosition.x, circlePosition.y, circlePosition.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        parOpeningDiameter / 2)
    
    oElements2.Add objCircle

    Set objCStr2 = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements2)
    ObjBottSurface.AddBoundary objCStr2
    
    'Set the top plane output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottSurface
    Set ObjBottSurface = Nothing
    oElements2.Clear
    objCStr2.RemoveCurve True
                
    'Remove the complex string
    Dim Objcurves1 As IJDObject
    Set Objcurves1 = oComplexString1
    Objcurves1.Remove
    Set oComplexString1 = Nothing
    Set oTLine1 = Nothing
    Set oTArc2 = Nothing
    Set oTLine2 = Nothing
    Set oTArc1 = Nothing
    For count = 1 To obotCollection.count
        obotCollection.Remove 1
    Next count
    Set obotCollection = Nothing
    Set oCenterPoint = Nothing
    Set oNormalVector = Nothing
    For count = 1 To 4
        Set arrPt(count) = Nothing
    Next count
    Set circlePosition = Nothing
    Set circleNormal = Nothing
    Set objCircle = Nothing
    Set geomFactory = Nothing

    
    Exit Sub
    
ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

