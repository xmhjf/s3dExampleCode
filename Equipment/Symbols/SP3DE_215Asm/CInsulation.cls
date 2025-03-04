VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:16:20 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
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
'   25.Jul.2003     SSP                  Corrected the center location of Ellipses for correct Insulation.
'                                           (TR-44842  Update Equipment symbols, so as to position the equipment by constraints  )
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   16.May.2005     kkk         CR-76070: Removed the Nozzle related parameters and code to create insulation for nozzles.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim parNumberOfSupports  As Double   'Number of Support
    Dim parSupportHeight As Double
    Dim parSupportLength As Double
    Dim parSupportThickness As Double
    Dim parVesselStartPoint  As Double
    Dim parInsulationThickness As Double
    
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
    
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

' Insert your code for output 1(Vessel (Shell))
    Dim cenX As Double, cenY As Double, cenZ As Double
    Dim ObjVesselIns As Object
    Dim MajorX As Double, MajorY As Double
    Dim MajorZ As Double, mMRatio As Double, StartAngle As Double, SweepAngle As Double
    Dim norX As Double, norY As Double, norZ As Double
    Dim startPoint As New AutoMath.DPosition
    Dim oEqpLine As IngrGeom3D.Line3d
    Dim oEqpArcTop As IngrGeom3D.EllipticalArc3d
    Dim oEqpArcBottom As IngrGeom3D.EllipticalArc3d
    Dim oEqpComplexStr As IngrGeom3D.ComplexString3d
' Vessel is formed by creating a Complex string and rotating about Z-Axis
'Top Convex Dome Curve
    cenX = 0
    cenY = 0
    cenZ = parVesselHeight + parVesselStartPoint
    MajorX = parVesselDiameter / 2 + parInsulationThickness
    MajorY = 0
    MajorZ = 0
    mMRatio = (parVesselDiameter / 4 + parInsulationThickness) / (parVesselDiameter / 2 + parInsulationThickness)
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
    Linepts(0) = parVesselDiameter / 2 + parInsulationThickness
    Linepts(1) = 0
    Linepts(2) = parVesselHeight + parVesselStartPoint
    Linepts(3) = parVesselDiameter / 2 + parInsulationThickness
    Linepts(4) = 0
    Linepts(5) = parVesselStartPoint
    Set oEqpLine = geomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                        Linepts(3), Linepts(4), Linepts(5))
    'Bottom Dome Curve
    cenX = 0
    cenY = 0
    cenZ = parVesselStartPoint
    MajorX = parVesselDiameter / 2 + parInsulationThickness
    MajorY = 0
    MajorZ = 0
    mMRatio = (parVesselDiameter / 4 + parInsulationThickness) / (parVesselDiameter / 2 + parInsulationThickness)
    StartAngle = PI
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    Set oEqpArcBottom = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
                                                            
    Dim EleCollection           As Collection
    Dim axisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
    
    Set EleCollection = New Collection
    EleCollection.Add oEqpArcTop
    EleCollection.Add oEqpLine
    EleCollection.Add oEqpArcBottom
    
    startPoint.Set 0, 0, (parVesselHeight + parVesselStartPoint + parVesselDiameter / 4 + parInsulationThickness)
    Set oEqpComplexStr = PlaceTrCString(startPoint, EleCollection)
    axisVect.Set 0, 0, 1
    revCenPt.Set 0, 0, (parVesselHeight / 2 + parVesselStartPoint)
    'Revolve it about Z-Axiz
    Set ObjVesselIns = PlaceRevolution(m_OutputColl, oEqpComplexStr, axisVect, revCenPt, PI * 2, True)

' Set the output
    m_OutputColl.AddOutput ("VesselInsul"), ObjVesselIns
    Set ObjVesselIns = Nothing
    
    Set axisVect = Nothing
    Set revCenPt = Nothing
    
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
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
