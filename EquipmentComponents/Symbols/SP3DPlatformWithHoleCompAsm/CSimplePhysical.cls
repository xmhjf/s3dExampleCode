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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:14:55 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-2005, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          MS
'   Creation Date:  Friday, 2 May 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy  who                change description
'   -----------  -----              ------------------
'   09.Jul.2003  SymbolTeam(India)  Copyright Information, Header  is added.
'   08.Aug.2003  SSP             TR-45806  Error in PlatformWithHoleForPipe symbol:Added small clearance
'                                   between the platform inner /outer arc and circle boundary.
'   23.Aug.2003  SymbolTeam(India)  Updated Error Handling Code
'   23.Jun.2004  svsmylav(svsmylav)   TR-60704: Replaced 'm_OutputColl.ResourceManager' with "Nothing'
'                                   while creating oTLine1, oTArc1, oTLine2, oTArc2 and oTopSurface transient outputs to
'                                   avoid extra curves/lines in the symbol graphics.
'   29.Nov.2004  V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   02.Mar.2005  svsmylav(svsmylav)   TR 61802: For invalid input of Radial Position, detailed error log is added and error is raised.
'   16.May.2005  svsmylav             CR-76070: Input parameter parRadialPosition is replaced with parPlatformHoletoCenter.

'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private m_oErrors As IJEditErrors
Private PI As Double
Private Const E_FAIL = &H80004005
Private Const MODULE = "Physical:" 'Used for error messages
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
    PI = Atn(1) * 4
    Set m_oErrors = New IMSErrorLog.JServerErrors
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oErrors = Nothing
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
     
    'Protect Case of 360 degrees.
    If (parPlatformAngle = 2 * PI) Then
        parPlatformAngle = 6.2831
    End If
    'Check if the Hole is going out of the Platform. Aleast it would touch the inner or outer edge of the platform
    Dim VessCentoHoleCen As Double
    Dim VessCentoPlatformOutedge As Double

    VessCentoHoleCen = parVesselDiameter / 2 + parClearancefromVessel + parInsulationThickness + _
                        parOpeningDiameter / 2 + 0.001

    VessCentoPlatformOutedge = parVesselDiameter / 2 + parClearancefromVessel + parInsulationThickness + _
                parPlatformWidth - parOpeningDiameter / 2 - 0.001

    Dim strDescritption As String
    Dim strValidRadialPositionRng As String
    strValidRadialPositionRng = "Valid Radial Position value range is from " & CStr(VessCentoHoleCen) & " to " & _
                                    CStr(VessCentoPlatformOutedge) & " Meters."
    If (parPlatformHoletoCenter < VessCentoHoleCen) Then
        strDescritption = "parPlatformHoletoCenter value = " & CStr(parPlatformHoletoCenter) & " Meters" & _
            " which is less than the computed 'Hole center to Vessel center' distance. " & strValidRadialPositionRng & vbCrLf & _
            "{Minimum value of 'Hole center to Vessel center' is sum of Vessel Radius (Vessel Diameter /2) , Clearance from Vessel, Opening Radius (Opening Diameter /2), Insulation Thickness and 0.001 m (small gap)}"
        m_oErrors.Add E_FAIL, MODULE & METHOD, strDescritption, "UserSymbol"
        Err.Raise E_FAIL, MODULE & METHOD, strDescritption, Err.HelpFile, Err.HelpContext
    End If

    If (parPlatformHoletoCenter > VessCentoPlatformOutedge) Then
'        parPlatformHoletoCenter = VessCentoPlatformOutedge
        strDescritption = "parPlatformHoletoCenter value = " & CStr(parPlatformHoletoCenter) & " Meters" & _
            " which makes the hole boundary to cross the Platform outer edge. " & strValidRadialPositionRng & vbCrLf & _
            "{Maximum value of 'Hole center to Vessel center' is sum of Vessel Radius (Vessel Diameter /2) , Clearance from Vessel, Insulation Thickness, Platform Width and to subtract Opening Radius (Opening Diameter /2) and 0.001 m (small gap)}"
        m_oErrors.Add E_FAIL, MODULE & METHOD, strDescritption, "UserSymbol"
        Err.Raise E_FAIL, MODULE & METHOD, strDescritption, Err.HelpFile, Err.HelpContext
    End If
    
    'If the parOpeningDiameter is given as zero place a very small hole which is almost not visible
    If (parOpeningDiameter = 0) Then
        parOpeningDiameter = 0.001
    End If

    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
     
    'Creation of Bottom Platform
    'Create the lines and arcs of the top surface . use this lines and arcs to create the complexstring.
    'Use the complex string to project the thickness.
     Dim oComplexString     As New IngrGeom3D.ComplexString3d
     Dim oTLine1 As Object
     Dim oTArc2 As Object
     Dim oTLine2 As Object
     Dim oTArc1 As Object
     Dim ObjPlatformBody As Object
        
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
    Set oTLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPt(1).x, arrPt(1).y, arrPt(1).z, arrPt(2).x, arrPt(2).y, arrPt(2).z)
    oCollection.Add oTLine1
    'Top Arc 2
    Set oTArc2 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                          oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                                         oNormalVector.x, oNormalVector.y, oNormalVector.z, _
                                                         arrPt(2).x, arrPt(2).y, arrPt(2).z, _
                                                        arrPt(3).x, arrPt(3).y, arrPt(3).z)
    oCollection.Add oTArc2
    'Top Line 2
    Set oTLine2 = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPt(3).x, arrPt(3).y, arrPt(3).z, arrPt(4).x, arrPt(4).y, arrPt(4).z)
    oCollection.Add oTLine2
    'Top Arc 1
    oNormalVector.Set 0, 0, -1
    Set oTArc1 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                      oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                                     oNormalVector.x, oNormalVector.y, oNormalVector.z, _
                                                     arrPt(4).x, arrPt(4).y, arrPt(4).z, _
                                                    arrPt(1).x, arrPt(1).y, arrPt(1).z)
    oCollection.Add oTArc1
    
    Set oComplexString = PlaceTrCString(arrPt(1), oCollection)
        
    'Create an infinite plane and appy the complex string as boundary
    Dim oTopSurface As IngrGeom3D.Plane3d
    Set oTopSurface = geomFactory.Planes3d.CreateByPointNormal(Nothing, arrPt(1).x, arrPt(1).y, arrPt(1).z, 0, 0, 1)
    oTopSurface.AddBoundary oComplexString
    
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
    oTopSurface.AddBoundary objCStr
    
    oElements.Clear
    objCStr.RemoveCurve True
    
    Dim Surfset   As IngrGeom3D.IJElements
    Dim ObjTopSupport As Object
    Dim WCaps  As Long
    WCaps = 1

    circleNormal.Set 0, 0, -1
    Set Surfset = geomFactory.GeometryServices.CreateByProjection(m_OutputColl.ResourceManager, _
                                                                oTopSurface, circleNormal.x, circleNormal.y, circleNormal.z, _
                                                                parPlatformHeight, WCaps)

    For Each ObjTopSupport In Surfset
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopSupport
    Next ObjTopSupport

    Set oTopSurface = Nothing
    Set circlePosition = Nothing
    Set circleNormal = Nothing
    Set objCircle = Nothing
    
     Dim Objcurves As IJDObject
     Set Objcurves = oComplexString
     Objcurves.Remove
     Set oComplexString = Nothing
     Set ObjTopSupport = Nothing

     Dim count As Integer
     For count = 1 To oCollection.count
         oCollection.Remove 1
     Next count
     Set oCollection = Nothing
     Set oCenterPoint = Nothing
     Set oNormalVector = Nothing
    Set oTLine1 = Nothing
    Set oTArc2 = Nothing
    Set oTLine2 = Nothing
    Set oTArc1 = Nothing
    For count = 1 To 4
        Set arrPt(count) = Nothing
    Next count
    Set geomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub


