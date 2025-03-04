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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:13 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Wednesday, Apr 3 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This is Spherical Vessel (E230) symbol.
' Symbol is created using the following Outputs:
'   i)  The outputs consists of following:
'       a) One Insulation aspect outputs,
'       b) Two Physical aspect outputs:
'           A Vessel uses 'PlaceSphere', a ControlPoint and
'   ii) Variable number of Supports (Can be either cylindrical or Cuboid) and
'       default surfaces are computed as per the user input.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   11.May.2005     svsmylav        CR-76070: Removed the Nozzle related parameters and code to create nozzles.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Const CYLINDRICAL_SUPPORTS = 1
Private Const CUBOID_SUPPORTS = 2
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjVessel As Object
    
    Dim parVesselDiameter As Double
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupports As Double
    Dim parVesselCenterHeight As Double
    Dim parSupportLength As Double
    Dim parSupportThickness As Double
    Dim parSupportRadialLocation As Double
    Dim parSupportHeight As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)     'P2
    parSupportAngularLocation = arrayOfInputs(3) 'P3
    parNumberOfSupports = arrayOfInputs(4)   'P4
    parVesselCenterHeight = arrayOfInputs(5) 'P5
    parSupportLength = arrayOfInputs(6)      'P6
    parSupportThickness = arrayOfInputs(7)   'P7
    parSupportRadialLocation = arrayOfInputs(8) 'P8
    parSupportHeight = arrayOfInputs(9)      'P9
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0
    
' Insert your code for output 2(Vessel)
    Dim centPoint As New AutoMath.DPosition
    centPoint.Set CenterPos.x, _
                  CenterPos.y, _
                  CenterPos.z + parVesselCenterHeight
    Set ObjVessel = PlaceSphere(m_OutputColl, centPoint, parVesselDiameter / 2)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVessel
    Set ObjVessel = Nothing
    
   If parNumberOfSupports > 0 And parSupportLength > 0 And parSupportHeight > 0 Then
'      Insert your code for output  (Supports/DefaultSurfaces)
        Dim supportType As Integer
        Dim normalVect As New AutoMath.DVector
        Dim oCircle As IngrGeom3D.Circle3d
        Dim parAngle As Double
        Dim I As Integer
    
'       Set the number of Supports/DefaultSurfaces and their type
        Dim ObjSupport As Object
        Dim ObjDefaultSurface As Object

        If parSupportThickness = 0 Then
            supportType = CYLINDRICAL_SUPPORTS
        Else
            supportType = CUBOID_SUPPORTS
        End If

        normalVect.Set 0, 0, 1
        
        Dim geomFactory     As IngrGeom3D.GeometryFactory
        Set geomFactory = New IngrGeom3D.GeometryFactory
    
'       Place supports
        Select Case supportType
           Case CYLINDRICAL_SUPPORTS
                For I = 1 To parNumberOfSupports
'                   Center point position in horizontal plane is needed
                    parAngle = PI / 2 - parSupportAngularLocation + _
                                    (I - 1) * (2 * PI / parNumberOfSupports)
                    stPoint.Set CenterPos.x + parSupportRadialLocation * Cos(parAngle), _
                                CenterPos.y + parSupportRadialLocation * Sin(parAngle), _
                                CenterPos.z
                    enPoint.Set stPoint.x, _
                                stPoint.y, _
                                stPoint.z + parSupportHeight
                    Set ObjSupport = PlaceCylinder(m_OutputColl, stPoint, enPoint, parSupportLength, True)
'                   Set the output
                    m_OutputColl.AddOutput "Support_", ObjSupport
                    Set ObjSupport = Nothing
                    
'                   Place DefaultSurface
                    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                    stPoint.x, stPoint.y, stPoint.z, _
                                                    normalVect.x, normalVect.y, normalVect.z, _
                                                    parSupportLength / 2)
                    Dim oComplexStr As New IngrGeom3D.ComplexString3d
                    Call oComplexStr.AddCurve(oCircle, False)
                    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
                                                                        stPoint.x, stPoint.y, stPoint.z, _
                                                                        normalVect.x, normalVect.y, normalVect.z)
                    Call ObjDefaultSurface.AddBoundary(oComplexStr)
'                   Set the output
                    m_OutputColl.AddOutput "DefaultSurfaces_", ObjDefaultSurface
                    Set ObjDefaultSurface = Nothing
'                   Remove curve
                    Dim ObjTmpcurves As IJDObject
                    Set ObjTmpcurves = oComplexStr
                    ObjTmpcurves.Remove
                    Set oComplexStr = Nothing
                    Set ObjTmpcurves = oCircle
                    ObjTmpcurves.Remove
                    Set oCircle = Nothing
                Next I
                Set normalVect = Nothing
    
           Case CUBOID_SUPPORTS
'               Consider an initial support cross section to be on the X axis, in horizontal plane
                Dim initSecPoints(0 To 14)  As Double
    
'               Left hand side Bottom point
                initSecPoints(0) = CenterPos.x + parSupportRadialLocation - parSupportThickness / 2
                initSecPoints(1) = CenterPos.y - parSupportLength / 2
                initSecPoints(2) = CenterPos.z
    
'               Right hand side Bottom point
                initSecPoints(3) = CenterPos.x + parSupportRadialLocation + parSupportThickness / 2
                initSecPoints(4) = CenterPos.y - parSupportLength / 2
                initSecPoints(5) = CenterPos.z
                
'               Right hand side Top point
                initSecPoints(6) = CenterPos.x + parSupportRadialLocation + parSupportThickness / 2
                initSecPoints(7) = CenterPos.y + parSupportLength / 2
                initSecPoints(8) = CenterPos.z
            
'               Left hand side Top point
                initSecPoints(9) = CenterPos.x + parSupportRadialLocation - parSupportThickness / 2
                initSecPoints(10) = CenterPos.y + parSupportLength / 2
                initSecPoints(11) = CenterPos.z
                
'               Left hand side Bottom point
                initSecPoints(12) = initSecPoints(0)
                initSecPoints(13) = initSecPoints(1)
                initSecPoints(14) = initSecPoints(2)
            
    '           Prepare profile points
                Dim oLineString As IngrGeom3D.LineString3d
                Dim objLineString As IJDObject
                Dim lineStrPoints(0 To 14)  As Double
                Dim J As Integer
                
                Dim ProjVector As New AutoMath.DVector
                ProjVector.Set 0, 0, 1
    
                For I = 1 To parNumberOfSupports
                    Set oLineString = New IngrGeom3D.LineString3d
'                   Center point position in horizontal plane is needed
                    parAngle = PI / 2 - parSupportAngularLocation + _
                                    (I - 1) * (2 * PI / parNumberOfSupports)
    
'                   Rotate initial support cross section to the required location
                    For J = 0 To 14 Step 3
                        lineStrPoints(J) = initSecPoints(J) * Cos(parAngle) - initSecPoints(J + 1) * Sin(parAngle)
                        lineStrPoints(J + 1) = initSecPoints(J) * Sin(parAngle) + initSecPoints(J + 1) * Cos(parAngle)
                        lineStrPoints(J + 2) = initSecPoints(J + 2)
                    Next J
                    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, lineStrPoints)
                    Set ObjSupport = PlaceProjection(m_OutputColl, oLineString, ProjVector, parSupportHeight, True)
'                   Set the output
                    m_OutputColl.AddOutput "Supports_", ObjSupport
                    Set ObjSupport = Nothing
                    
'                   Place DefaultSurface
                    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)
'                   Set the output
                    m_OutputColl.AddOutput "DefaultSurfaces_", ObjDefaultSurface
                    Set ObjDefaultSurface = Nothing

'                   Remove  linestring
                    Set objLineString = oLineString
                    objLineString.Remove
                    Set oLineString = Nothing
                Next I
                Set ProjVector = Nothing
                Set geomFactory = Nothing
        End Select
    End If
    Set centPoint = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set CenterPos = Nothing
    Set ObjTmpcurves = Nothing
    Set ObjTmpcurves = Nothing
    Set objLineString = Nothing
    
'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.05)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "SphericalVesselControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
        
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
