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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:16:18 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Tuesday, Apr 2 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'The Physical Aspect of Equipment symbol is created with a Main Vessel (Complex String consist of an Elliptical curve,
'a Vertical Line and a bottom Elliptical Curve), Variable number of supports( Box), Four Nozzles (with Length).
'one on Top, the second along X-axis,  the third along Y-axis and fourth nozzle at the bottom of
'Vessel. Default Surfaces for Physical Aspect are placed at the bottom of Supports.
'The Insulation aspect consist of insulation for main vessel (Complex String consist of an Elliptical curve
'Vertical Line and a bottom Elliptical Curve)and Nozzles(Insulation Diameter is equal to Falnge Dia
'and Insulation thickness).
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   25.Jul.2003     SSP                              Changing Plane Projection for Support legs construction to six planes
'                                                                       and added Default surface at the control point.
'                                                                     (TR-44842  Update Equipment symbols, so as to position the equipment by constraints  )
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   16.May.2005     kkk         CR-76070: Removed the Nozzle related parameters and code to create nozzles.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   28.Sep.2006    RRK  TR-84724  If variables of type double are compared, appropriate CmpDblXXX functions are used.
'                                 The Type of the variable "parNumberOfSupports" is changed from Double to Integer
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
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
    Dim parVesselHeight As Double
    Dim parVesselDiameter As Double
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupports  As Integer   'Number of Support
    Dim parSupportHeight As Double
    Dim parSupportLength As Double
    Dim parSupportThickness As Double
    Dim parVesselStartPoint  As Double
    Dim parInsulationThickness As Double
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselHeight = arrayOfInputs(2)  'P2
    parVesselDiameter = arrayOfInputs(3)    'P3
    parSupportAngularLocation = arrayOfInputs(4)  'P4
    parNumberOfSupports = arrayOfInputs(5)  'P5
    parSupportHeight = arrayOfInputs(6)     'P6
    parSupportLength = arrayOfInputs(7)     'P7
    parSupportThickness = arrayOfInputs(8)  'P8
    parVesselStartPoint = arrayOfInputs(9)      'P9
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0

    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

 ' Insert your code for output 1(Vessel (Shell))
    Dim cenX As Double, cenY As Double, cenZ As Double
    Dim MajorX As Double, MajorY As Double
    Dim MajorZ As Double, mMRatio As Double, StartAngle As Double, SweepAngle As Double
    Dim norX As Double, norY As Double, norZ As Double
    Dim oEqpLine As IngrGeom3D.Line3d
    Dim oEqpArcTop As IngrGeom3D.EllipticalArc3d
    Dim oEqpArcBottom As IngrGeom3D.EllipticalArc3d
    Dim oEqpComplexStr As IngrGeom3D.ComplexString3d

' Vessel is formed by creating a Complex string and rotating about Z-Axis
'Top Convex Dome Curve
    cenX = 0
    cenY = 0
    cenZ = parVesselHeight + parVesselStartPoint
    MajorX = parVesselDiameter / 2
    MajorY = 0
    MajorZ = 0
    mMRatio = 0.5
    StartAngle = PI / 2
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    Set oEqpArcTop = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
    'Vertical Line
    Dim Linepts(0 To 5) As Double
    Linepts(0) = parVesselDiameter / 2
    Linepts(1) = 0
    Linepts(2) = parVesselHeight + parVesselStartPoint
    Linepts(3) = parVesselDiameter / 2
    Linepts(4) = 0
    Linepts(5) = parVesselStartPoint
    Set oEqpLine = geomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                        Linepts(3), Linepts(4), Linepts(5))
    'Bottom Dome Curve
    cenX = 0
    cenY = 0
    cenZ = parVesselStartPoint
    MajorX = parVesselDiameter / 2
    MajorY = 0
    MajorZ = 0
    mMRatio = 0.5
    StartAngle = PI
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    Set oEqpArcBottom = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
                                                            
    Dim EleCollection           As Collection
    Dim startPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
    
    Set EleCollection = New Collection
    EleCollection.Add oEqpArcTop
    EleCollection.Add oEqpLine
    EleCollection.Add oEqpArcBottom
    
    startPoint.Set 0, 0, (parVesselHeight + parVesselStartPoint + parVesselDiameter / 4)
    Set oEqpComplexStr = PlaceTrCString(startPoint, EleCollection)
    axisVect.Set 0, 0, 1
    revCenPt.Set 0, 0, (parVesselHeight / 2 + parVesselStartPoint)
    'Revolve it about Z-Axiz
    Set ObjVessel = PlaceRevolution(m_OutputColl, oEqpComplexStr, axisVect, revCenPt, PI * 2, True)

' Set the output
    m_OutputColl.AddOutput ("Vessel"), ObjVessel
    Set ObjVessel = Nothing
    Set axisVect = Nothing
    Set revCenPt = Nothing
    Set startPoint = Nothing
    
    Dim Objcurves As IJDObject
    Set Objcurves = oEqpComplexStr
    Objcurves.Remove
    
    Set Objcurves = Nothing
    Set oEqpArcBottom = Nothing
    Set oEqpLine = Nothing
    Set oEqpArcTop = Nothing
    
    Dim count As Integer
    For count = 1 To EleCollection.count
        EleCollection.Remove 1
    Next count
    Set EleCollection = Nothing
    
' Insert your code for output (Leg Support )

    Dim IncludedAng As Double
    Dim yAxistoSupport As Double
    Dim iCount As Integer
    If (parNumberOfSupports > 0) And CmpDblGreaterthan(parSupportHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) _
                                                    And CmpDblGreaterthan(parSupportThickness, LINEAR_TOLERANCE) Then
        IncludedAng = (PI * 2) / parNumberOfSupports
        yAxistoSupport = parSupportAngularLocation
        Dim topSurPts(3) As IJDPosition
        Dim botSurPts(3) As IJDPosition
        
        For iCount = 0 To 3
            Set topSurPts(iCount) = New DPosition
            Set botSurPts(iCount) = New DPosition
        Next iCount
        
        For iCount = 0 To parNumberOfSupports - 1
            'form Line String and Project it
            'Point 1
            botSurPts(0).x = Sin(yAxistoSupport) * (parVesselDiameter / 2) + Cos(yAxistoSupport) * parSupportLength / 2
            botSurPts(0).y = Cos(yAxistoSupport) * (parVesselDiameter / 2) - Sin(yAxistoSupport) * parSupportLength / 2
            botSurPts(0).z = 0
            'Point 2
            botSurPts(1).x = botSurPts(0).x + Sin(yAxistoSupport) * parSupportThickness
            botSurPts(1).y = botSurPts(0).y + parSupportThickness * Cos(yAxistoSupport)
            botSurPts(1).z = 0
           
            'Point 4
            botSurPts(3).x = Sin(yAxistoSupport) * (parVesselDiameter / 2) - Cos(yAxistoSupport) * parSupportLength / 2
            botSurPts(3).y = Cos(yAxistoSupport) * (parVesselDiameter / 2) + Sin(yAxistoSupport) * parSupportLength / 2
            botSurPts(3).z = 0
            
             'Point3
            botSurPts(2).x = botSurPts(3).x + parSupportThickness * Sin(yAxistoSupport)
            botSurPts(2).y = botSurPts(3).y + parSupportThickness * Cos(yAxistoSupport)
            botSurPts(2).z = botSurPts(3).z
            
            topSurPts(0).Set botSurPts(0).x, botSurPts(0).y, botSurPts(0).z + parSupportHeight
            topSurPts(1).Set botSurPts(1).x, botSurPts(1).y, botSurPts(1).z + parSupportHeight
            topSurPts(2).Set botSurPts(2).x, botSurPts(2).y, botSurPts(2).z + parSupportHeight
            topSurPts(3).Set botSurPts(3).x, botSurPts(3).y, botSurPts(3).z + parSupportHeight
            
            Dim ObjCollection As Collection
            
            Set ObjCollection = PlaceTrapezoidWithPlanes(m_OutputColl, topSurPts, botSurPts)
    
            m_OutputColl.AddOutput ("Support" & CStr(iCount + 1) & "Plane1"), ObjCollection(1)
            m_OutputColl.AddOutput ("Support" & CStr(iCount + 1) & "Plane2"), ObjCollection(2)
            m_OutputColl.AddOutput ("Support" & CStr(iCount + 1) & "Plane3"), ObjCollection(3)
            m_OutputColl.AddOutput ("Support" & CStr(iCount + 1) & "Plane4"), ObjCollection(4)
            m_OutputColl.AddOutput ("Support" & CStr(iCount + 1) & "Plane5"), ObjCollection(5)
            m_OutputColl.AddOutput ("Support" & CStr(iCount + 1) & "Plane6"), ObjCollection(6)
    
            Set ObjCollection = Nothing
    
            'Increment the Angle
            yAxistoSupport = yAxistoSupport + IncludedAng
        Next iCount

End If

'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0, , , True)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "SimpleVerCylEqpLegsControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    
    
' Insert your code for output 12(Default Surface /plane at Control Point of vessel )
    ' create non-persistent circle to use for creating default surface ---
    
    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Dim objCircle As IngrGeom3D.Circle3d
    
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                0, 0, 0, _
                                0, 0, -1, _
                                0.1 * parVesselDiameter / 2)

'   create persistent default surface plane - the plane can mate ---
    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
' Set the output
    m_OutputColl.AddOutput "DefaultSurface", ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    
    Set objCircle = Nothing
    Set m_oGBSFactory = Nothing
    Set geomFactory = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
   
End Sub
