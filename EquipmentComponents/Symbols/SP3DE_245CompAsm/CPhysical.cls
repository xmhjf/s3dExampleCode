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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:16:00 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-06 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:        svsmylav
'   Creation Date:  Thursday, Apr 1 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is Simple Horizontal Cylindrical Vessel Component symbol.
'    Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E245 Symbol in Page no 298.
'   Symbol is created using the following Outputs:
'   i)  Seven standard outputs Consisting of the following:
'       a) Three Insulation aspect outputs,
'       b) Three Physical aspect outputs:
'           Vessel created using 'PlaceCylinder' and two domes using 'PlaceRevolution'
'       c) Two ReferenceGeometry aspect outputs: default surface and a control point.
'
'   ii) Variable number of Supports (Maximum two supports) and surfaces are
'       computed as per the user input.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006     kkc       DI 95670-Replaced names with initials in the revision history.
'   26.Sep.2006     dkl       TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "SimplePhysical:" 'Used for error messages
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
    Dim ObjVesselLEH As Object
    Dim ObjVesselREH As Object

    Dim parVesselLength As Double
    Dim parVesselDiameter As Double
    Dim parVesselCenterHeight As Double
    Dim parFirstSupportLocation As Double
    Dim parSecondSupportLocation As Double
    Dim parSupportThickness As Double
    Dim parSupportLength As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselLength = arrayOfInputs(2)      'P2
    parVesselDiameter = arrayOfInputs(3)    'P3
    parVesselCenterHeight = arrayOfInputs(4)    'P4
    parFirstSupportLocation = arrayOfInputs(5)  'P5
    parSecondSupportLocation = arrayOfInputs(6) 'P6
    parSupportThickness = arrayOfInputs(7)  'P7
    parSupportLength = arrayOfInputs(8)     'P8
    parInsulationThickness = arrayOfInputs(9)

    iOutput = 0
    
'   Origin is taken at left dome center point (i.e. Placement point 1)
'   Vessel is formed by creating a Cylinder and two Elliptical Heads rotating about X-Axis

' Insert your code for output 1(Vessel)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    stPoint.Set 0, 0, 0
    enPoint.Set parVesselLength, 0, 0
    Set ObjVessel = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVessel
    Set ObjVessel = Nothing

' Insert your code for output 2(VesselLEH)
'   Revolve Left Elliptical Dome Curve about X-Axis
    Dim cenX As Double
    Dim cenY As Double
    Dim cenZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim mMRatio As Double
    Dim StartAngle As Double
    Dim SweepAngle As Double
    Dim norX As Double
    Dim norY As Double
    Dim norZ As Double

    Dim oEqpLeftArc As IngrGeom3D.EllipticalArc3d

    cenX = 0
    cenY = 0
    cenZ = 0
    MajorX = 0
    MajorY = 0
    MajorZ = parVesselDiameter / 2
    mMRatio = 0.5
    StartAngle = PI / 2
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0

    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Set oEqpLeftArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                           cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                           StartAngle, SweepAngle)
    Dim axisVect As New AutoMath.DVector
    Dim VesselRevCenPt As New AutoMath.DPosition

    axisVect.Set 1, 0, 0
    VesselRevCenPt.Set parVesselLength / 2, 0, 0

    Set ObjVesselLEH = PlaceRevolution(m_OutputColl, oEqpLeftArc, axisVect, VesselRevCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselLEH
    Set ObjVesselLEH = Nothing
    Set oEqpLeftArc = Nothing

' Insert your code for output 3(VesselREH)
'   Revolve Right Elliptical Dome Curve about X-Axis
    Dim oEqpRightArc As IngrGeom3D.EllipticalArc3d

    cenX = parVesselLength
    cenY = 0
    cenZ = 0
    MajorX = 0
    MajorY = 0
    MajorZ = parVesselDiameter / 2
    mMRatio = 0.5
    StartAngle = PI
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    Set oEqpRightArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)

'   Revolve Right Elliptical Head about X-Axis
    Set ObjVesselREH = PlaceRevolution(m_OutputColl, oEqpRightArc, axisVect, VesselRevCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselREH
    Set ObjVesselREH = Nothing
    Set axisVect = Nothing
    Set VesselRevCenPt = Nothing
    Set oEqpRightArc = Nothing

    If CmpDblGreaterthan(parVesselCenterHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportThickness, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) Then
        Dim ObjSupport As Object
        Dim ObjDefaultSurface As Object
        If CmpDblGreaterthan(parFirstSupportLocation, LINEAR_TOLERANCE) Then
'           Place First Support
            Dim lineStrPoints(0 To 14)  As Double
            lineStrPoints(0) = parFirstSupportLocation - parSupportThickness / 2
            lineStrPoints(1) = -parSupportLength / 2
            lineStrPoints(2) = -parVesselCenterHeight
            
            lineStrPoints(3) = lineStrPoints(0)
            lineStrPoints(4) = lineStrPoints(1) + parSupportLength
            lineStrPoints(5) = lineStrPoints(2)

            lineStrPoints(6) = lineStrPoints(3) + parSupportThickness
            lineStrPoints(7) = lineStrPoints(4)
            lineStrPoints(8) = lineStrPoints(5)

            lineStrPoints(9) = lineStrPoints(6)
            lineStrPoints(10) = lineStrPoints(7) - parSupportLength
            lineStrPoints(11) = lineStrPoints(8)

            lineStrPoints(12) = lineStrPoints(0)
            lineStrPoints(13) = lineStrPoints(1)
            lineStrPoints(14) = lineStrPoints(2)
            
            axisVect.Set 0, 0, 1
            Set ObjSupport = PlaceProjectedBox(m_OutputColl, lineStrPoints, axisVect, parVesselCenterHeight, False)

'           Set the output
            m_OutputColl.AddOutput "Support_", ObjSupport
            Set ObjSupport = Nothing

'           Create First DefaultSurface
            Set ObjDefaultSurface = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)

'           Set the output
            m_OutputColl.AddOutput "DefaultSurface_", ObjDefaultSurface
            Set ObjDefaultSurface = Nothing
        End If

        If CmpDblGreaterthan(parSecondSupportLocation, LINEAR_TOLERANCE) Then
'           Place Second Support
            lineStrPoints(0) = parFirstSupportLocation + parSecondSupportLocation _
                                - parSupportThickness / 2
            lineStrPoints(1) = -parSupportLength / 2
            lineStrPoints(2) = -parVesselCenterHeight
            
            lineStrPoints(3) = lineStrPoints(0)
            lineStrPoints(4) = lineStrPoints(1) + parSupportLength
            lineStrPoints(5) = lineStrPoints(2)

            lineStrPoints(6) = lineStrPoints(3) + parSupportThickness
            lineStrPoints(7) = lineStrPoints(4)
            lineStrPoints(8) = lineStrPoints(5)

            lineStrPoints(9) = lineStrPoints(6)
            lineStrPoints(10) = lineStrPoints(7) - parSupportLength
            lineStrPoints(11) = lineStrPoints(8)
            
            lineStrPoints(12) = lineStrPoints(0)
            lineStrPoints(13) = lineStrPoints(1)
            lineStrPoints(14) = lineStrPoints(2)
            
            Set ObjSupport = PlaceProjectedBox(m_OutputColl, lineStrPoints, axisVect, parVesselCenterHeight, False)

'           Set the output
            m_OutputColl.AddOutput "Support_", ObjSupport
            Set ObjSupport = Nothing

'           Create Second DefaultSurface
            Set ObjDefaultSurface = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)

'           Set the output
            m_OutputColl.AddOutput "DefaultSurface_", ObjDefaultSurface
            Set ObjDefaultSurface = Nothing
        End If
    End If
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oGeomFactory = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub

Private Function PlaceProjectedBox(ByVal objOutputColl As Object, _
                        dPoints() As Double, _
                        ByVal axisVector As AutoMath.DVector, _
                        height As Double, _
                        isCapped As Boolean) As Object

    Const METHOD = "PlaceProjectedBox:"
    On Error GoTo ErrorHandler
      
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oLineStr As IngrGeom3D.LineString3d
        
    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
    Set objProjection = oGeomFactory.Projections3d.CreateByCurve( _
                                                    objOutputColl.ResourceManager, _
                                                    oLineStr, _
                                                    axisVector.x, axisVector.y, axisVector.z, _
                                                    height, isCapped)
    
    Set PlaceProjectedBox = objProjection
    Set objProjection = Nothing
    Set oLineStr = Nothing
    Set oGeomFactory = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
        
End Function

