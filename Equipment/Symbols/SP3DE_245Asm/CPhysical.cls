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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:51 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Friday, Apr 12 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol is created using the following Outputs:
'   i)  Eight standard outputs Consisting of the following:
'       a) Four Insulation aspect outputs,
'       b) Four Physical aspect outputs:
'           A Vessel uses 'PlaceRevolution' and three nozzles which use
'           'CreateNozzlePHWithLength' function, and
'   ii) Variable number of Supports (Maximum two supports) and default surfaces are
'       computed as per the user input.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   26.Sep.2006    RRK  TR84724   If variables of type double are compared, appropriate CmpDblXXX functions are used.
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
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
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselLength = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parVesselCenterHeight = arrayOfInputs(4)
    parFirstSupportLocation = arrayOfInputs(5)
    parSecondSupportLocation = arrayOfInputs(6)
    parSupportThickness = arrayOfInputs(7)
    parSupportLength = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0

' Insert your code for output 5(Vessel)
    Dim cenX As Double, cenY As Double, cenZ As Double
    Dim majorX As Double, majorY As Double
    Dim majorZ As Double, MMRatio As Double, StartAngle As Double, SweepAngle As Double
    Dim norX As Double, norY As Double, norZ As Double
    
    Dim oEqpLine As IngrGeom3D.Line3d
    Dim oEqpLeftArc As IngrGeom3D.EllipticalArc3d
    Dim oEqpRightArc As IngrGeom3D.EllipticalArc3d
'   Vessel is formed by creating a Cylinder and two Elliptical Heads rotating about X-Axis
'   Left Elliptical Dome Curve

    cenX = 0
    cenY = 0
    cenZ = 0
    majorX = 0
    majorY = 0
    majorZ = parVesselDiameter / 2
    MMRatio = 0.5
    StartAngle = PI / 2
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Set oEqpLeftArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, majorX, majorY, majorZ, MMRatio, _
                                                        StartAngle, SweepAngle)
'   Right Elliptical Dome Curve
    cenX = parVesselLength
    cenY = 0
    cenZ = 0
    majorX = 0
    majorY = 0
    majorZ = parVesselDiameter / 2
    MMRatio = 0.5
    StartAngle = PI
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    Set oEqpRightArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, majorX, majorY, majorZ, MMRatio, _
                                                        StartAngle, SweepAngle)
                                                            
    Dim startPoint As New AutoMath.DPosition
    Dim EndPoint As New AutoMath.DPosition

    Dim axisVect As New AutoMath.DVector
    Dim VesselRevCenPt As New AutoMath.DPosition
    
    
    startPoint.Set 0, 0, 0
    EndPoint.Set parVesselLength, 0, 0

    axisVect.Set 1, 0, 0
    VesselRevCenPt.Set parVesselLength / 2, 0, 0
'   Create Horizontal Drum
    Set ObjVessel = PlaceCylinder(m_OutputColl, startPoint, EndPoint, parVesselDiameter, True)
    Set startPoint = Nothing
    Set EndPoint = Nothing

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVessel
    Set ObjVessel = Nothing
    
'   Revolve Left Elliptical Head about X-Axis
    Set ObjVesselLEH = PlaceRevolution(m_OutputColl, oEqpLeftArc, axisVect, VesselRevCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselLEH
    Set ObjVesselLEH = Nothing
    
   '   Revolve Right Elliptical Head about X-Axis
    Set ObjVesselREH = PlaceRevolution(m_OutputColl, oEqpRightArc, axisVect, VesselRevCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselREH
    Set ObjVesselREH = Nothing
    Set axisVect = Nothing
    Set VesselRevCenPt = Nothing
    
    If CmpDblGreaterthan(parVesselCenterHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportThickness, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) Then
        Dim ObjSupport As Object
        Dim ObjDefaultSurface As Object
        Dim stPoint   As New AutoMath.DPosition
        Dim enPoint   As New AutoMath.DPosition
        If CmpDblGreaterthan(parFirstSupportLocation, LINEAR_TOLERANCE) Then
'           Place First Support
            stPoint.Set parFirstSupportLocation - parSupportThickness / 2, _
                     -parSupportLength / 2, _
                     -parVesselCenterHeight
            enPoint.Set stPoint.x + parSupportThickness, _
                    stPoint.y + parSupportLength, _
                    stPoint.z + parVesselCenterHeight
            Set ObjSupport = PlaceBox(m_OutputColl, stPoint, enPoint)
'           Set the output
            m_OutputColl.AddOutput "Support_", ObjSupport
            Set ObjSupport = Nothing

'           Create First DefaultSurface
            Dim lineStrPoints(0 To 11)  As Double
            lineStrPoints(0) = parFirstSupportLocation - parSupportThickness / 2
            lineStrPoints(1) = -parSupportLength / 2
            lineStrPoints(2) = -parVesselCenterHeight
        
            lineStrPoints(3) = lineStrPoints(0) + parSupportThickness
            lineStrPoints(4) = lineStrPoints(1)
            lineStrPoints(5) = lineStrPoints(2)
            
            lineStrPoints(6) = lineStrPoints(3)
            lineStrPoints(7) = lineStrPoints(4) + parSupportLength
            lineStrPoints(8) = lineStrPoints(5)
            
            lineStrPoints(9) = lineStrPoints(6) - parSupportThickness
            lineStrPoints(10) = lineStrPoints(7)
            lineStrPoints(11) = lineStrPoints(8)
        
            Set ObjDefaultSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)
        
'           Set the output
            m_OutputColl.AddOutput "DefaultSurface_", ObjDefaultSurface
            Set ObjDefaultSurface = Nothing
        End If
        
        If CmpDblGreaterthan(parSecondSupportLocation, LINEAR_TOLERANCE) Then
'           Place Second Support
            stPoint.Set parFirstSupportLocation + parSecondSupportLocation _
                                - parSupportThickness / 2, _
                         -parSupportLength / 2, _
                         -parVesselCenterHeight
            enPoint.Set stPoint.x + parSupportThickness, _
                        stPoint.y + parSupportLength, _
                        stPoint.z + parVesselCenterHeight
            Set ObjSupport = PlaceBox(m_OutputColl, stPoint, enPoint)
'           Set the output
            m_OutputColl.AddOutput "Support_", ObjSupport
            Set ObjSupport = Nothing
            
'           Create Second DefaultSurface
            lineStrPoints(0) = parFirstSupportLocation + parSecondSupportLocation _
                                - parSupportThickness / 2
            lineStrPoints(1) = -parSupportLength / 2
            lineStrPoints(2) = -parVesselCenterHeight
        
            lineStrPoints(3) = lineStrPoints(0) + parSupportThickness
            lineStrPoints(4) = lineStrPoints(1)
            lineStrPoints(5) = lineStrPoints(2)
            
            lineStrPoints(6) = lineStrPoints(3)
            lineStrPoints(7) = lineStrPoints(4) + parSupportLength
            lineStrPoints(8) = lineStrPoints(5)
            
            lineStrPoints(9) = lineStrPoints(6) - parSupportThickness
            lineStrPoints(10) = lineStrPoints(7)
            lineStrPoints(11) = lineStrPoints(8)
        
            Set ObjDefaultSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)
        
'           Set the output
            m_OutputColl.AddOutput "DefaultSurface_", ObjDefaultSurface
            Set ObjDefaultSurface = Nothing
        End If
        Set stPoint = Nothing
        Set enPoint = Nothing
        Set geomFactory = Nothing
    End If
    


     
'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.05)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "SimpleHorCylVesselControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing


    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


