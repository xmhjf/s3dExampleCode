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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:21 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Apr 1 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is Spherical Vessel Component symbol.
'    Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E230 Symbol in Page no 293.
'    Symbol is created using the following Outputs:
'    i)  Three standard outputs Consisting of the following:
'       a) One Insulation aspect output,
'       b) One Physical aspect output: A Vessel uses 'PlaceSphere',
'       c) Two ReferenceGeometry aspect outputs: Default surface and a control point.
'    ii) Variable number of Supports (Can be either cylindrical or Cuboid) and
'       default surfaces (for cylindrical supportscase) are computed as per the user input.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
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

    Dim iOutput     As Double
    Dim ObjVessel As Object

    Dim parVesselDiameter As Double
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupports As Long
    Dim parVesselCenterHeight As Double
    Dim parSupportLength As Double
    Dim parSupportThickness As Double
    Dim parSupportRadialLocation As Double
    Dim parSupportHeight As Double
    Dim parInsulationThickness As Double

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

'   Origin is taken at Place point1(DP/PP 1)

' Insert your code for output 2(Vessel)
    Dim centPoint As New AutoMath.DPosition

    centPoint.Set 0, 0, parVesselCenterHeight
    Set ObjVessel = PlaceSphere(m_OutputColl, centPoint, parVesselDiameter / 2)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVessel
    Set ObjVessel = Nothing
    Set centPoint = Nothing

   If parNumberOfSupports > 0 And parSupportLength > 0 And parSupportHeight > 0 Then
'      Insert your code for output  (Supports/DefaultSurfaces)
        Dim supportType As Integer
        Dim oNormalVect As New AutoMath.DVector
        Dim oCircle As IngrGeom3D.Circle3d
        Dim dAngle As Double
        Dim iCount As Integer

'       Set the number of Supports/DefaultSurfaces and their type
        Dim ObjSupport As Object
        Dim ObjSurface As Object

        If parSupportThickness = 0 Then
            supportType = CYLINDRICAL_SUPPORTS
        Else
            supportType = CUBOID_SUPPORTS
        End If

        oNormalVect.Set 0, 0, -1

        Dim stPoint   As New AutoMath.DPosition
        Dim enPoint   As New AutoMath.DPosition
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
'       Place supports
        Select Case supportType
           Case CYLINDRICAL_SUPPORTS
                For iCount = 1 To parNumberOfSupports
'                   Center point position in horizontal plane is needed
                    dAngle = parSupportAngularLocation + _
                                    (iCount - 1) * (2 * PI / parNumberOfSupports)
                    stPoint.Set parSupportRadialLocation * Sin(dAngle), _
                                parSupportRadialLocation * Cos(dAngle), 0
                    enPoint.Set stPoint.x, stPoint.y, stPoint.z + parSupportHeight
                    Set ObjSupport = PlaceCylinder(m_OutputColl, stPoint, enPoint, parSupportLength, False)

'                   Set the output
                    m_OutputColl.AddOutput "Support_", ObjSupport
                    Set ObjSupport = Nothing

'                   Place Surface at bottom of support
                    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                    stPoint.x, stPoint.y, stPoint.z, _
                                                    oNormalVect.x, oNormalVect.y, oNormalVect.z, _
                                                    parSupportLength / 2)
                    Dim oComplexStr As New IngrGeom3D.ComplexString3d
                    Call oComplexStr.AddCurve(oCircle, False)
                    Set ObjSurface = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
                                                                        stPoint.x, stPoint.y, stPoint.z, _
                                                                        oNormalVect.x, oNormalVect.y, oNormalVect.z)
                    Call ObjSurface.AddBoundary(oComplexStr)

'                   Set the output
                    m_OutputColl.AddOutput "Surface_", ObjSurface
                    Set ObjSurface = Nothing

'                   Remove curve
                    Dim ObjTmpcurves As IJDObject
                    Set ObjTmpcurves = oComplexStr
                    ObjTmpcurves.Remove
                    Set oComplexStr = Nothing
                    Set ObjTmpcurves = oCircle
                    ObjTmpcurves.Remove
                    Set oCircle = Nothing
                    Set ObjTmpcurves = Nothing
                Next iCount

        Case CUBOID_SUPPORTS
                Dim oTopPts(0 To 3) As IJDPosition
                Dim oBotPts(0 To 3) As IJDPosition

                For iCount = 0 To 3
                    Set oTopPts(iCount) = New DPosition
                    Set oBotPts(iCount) = New DPosition
                Next iCount

                Dim ObjColl As Collection
                For iCount = 1 To parNumberOfSupports
                    dAngle = parSupportAngularLocation + _
                                    (iCount - 1) * (2 * PI / parNumberOfSupports)
                    'Point 1
                    oBotPts(0).x = Sin(dAngle) * (parSupportRadialLocation) + Cos(dAngle) * parSupportLength / 2
                    oBotPts(0).y = Cos(dAngle) * (parSupportRadialLocation) - Sin(dAngle) * parSupportLength / 2
                    oBotPts(0).z = 0

                    'Point 2
                    oBotPts(1).x = oBotPts(0).x + Sin(dAngle) * parSupportThickness
                    oBotPts(1).y = oBotPts(0).y + parSupportThickness * Cos(dAngle)
                    oBotPts(1).z = 0

                    'Point 4
                    oBotPts(3).x = Sin(dAngle) * (parSupportRadialLocation) - Cos(dAngle) * parSupportLength / 2
                    oBotPts(3).y = Cos(dAngle) * (parSupportRadialLocation) + Sin(dAngle) * parSupportLength / 2
                    oBotPts(3).z = 0

                    'Point3
                    oBotPts(2).x = oBotPts(3).x + parSupportThickness * Sin(dAngle)
                    oBotPts(2).y = oBotPts(3).y + parSupportThickness * Cos(dAngle)
                    oBotPts(2).z = 0

                    oTopPts(0).Set oBotPts(0).x, oBotPts(0).y, parSupportHeight
                    oTopPts(1).Set oBotPts(1).x, oBotPts(1).y, parSupportHeight
                    oTopPts(2).Set oBotPts(2).x, oBotPts(2).y, parSupportHeight
                    oTopPts(3).Set oBotPts(3).x, oBotPts(3).y, parSupportHeight

                    Set ObjColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopPts, oBotPts)

'                   Set the output
                    Dim jCount As Integer
                    For jCount = 1 To ObjColl.Count
                        m_OutputColl.AddOutput "Support_", ObjColl(jCount)
                    Next jCount
                    Set ObjColl = Nothing
                Next iCount
        End Select
        For jCount = 0 To 3
            Set oTopPts(jCount) = Nothing
            Set oBotPts(jCount) = Nothing
        Next jCount
        Set oNormalVect = Nothing
        Set oGeomFactory = Nothing
        Set stPoint = Nothing
        Set enPoint = Nothing
    End If
    
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub
