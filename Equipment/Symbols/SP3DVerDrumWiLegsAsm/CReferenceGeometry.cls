VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReferenceGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:40 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         svsmylav
'   Creation Date:  Thusday, June 5 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol, Horizontal Drum , is prepared as per DOW Chemical Company manual.
'   All assumed values are as per the DOW documents unless otherwise specified. For more details
'   refer DOW document D-182.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit
Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "ReferenceGeometry:" 'Used for error messages

Dim PI As Double

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oGeomHelper = New SymbolServices
    PI = Atn(1) * 4
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    
   
    Dim parVesselDiameter   As Double
    Dim parVesselTantoTan   As Double
    Dim parCPtoSupport  As Double
    Dim parInsulationThickness  As Double
    
    Dim parSupportOrientation   As Double
    Dim parSupportLength    As Double
    Dim parSupportWidth As Double
    Dim parSupportHeight    As Double
    
    Dim parBackingPlateThickness    As Double
    Dim parBackingPlateDepth    As Double
    
    Dim parLengthBaseplate  As Double
    Dim parWidthBaseplate   As Double
    Dim parBasePlateThickness   As Double
    
    Dim parXboltHole    As Double
    Dim parYboltHole    As Double
    
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselTantoTan = arrayOfInputs(3)
    parCPtoSupport = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parSupportOrientation = arrayOfInputs(6)
    parSupportLength = arrayOfInputs(7)
    parSupportWidth = arrayOfInputs(8)
    parSupportHeight = arrayOfInputs(9)
    parBackingPlateThickness = arrayOfInputs(10)
    parBackingPlateDepth = arrayOfInputs(11)
    parLengthBaseplate = arrayOfInputs(12)
    parWidthBaseplate = arrayOfInputs(13)
    parBasePlateThickness = arrayOfInputs(14)
    parXboltHole = arrayOfInputs(15)
    parYboltHole = arrayOfInputs(16)
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    iOutput = 0
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim VesselOuterDiameter     As Double
    Dim shellThickness As Double
    Dim VesselRadius As Double
    
    Set geomFactory = New IngrGeom3D.GeometryFactory
    shellThickness = 0.0508         'Assumed value for Shell thickness
    VesselOuterDiameter = parVesselDiameter + 2 * shellThickness
    VesselRadius = VesselOuterDiameter / 2
                                            
    'Control Point
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, _
                                                                                                                0, 0, 0, 0.02, , , False)
    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    
    'output
    iOutput = iOutput + 1

    m_OutputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    
    'Default Surface
    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Dim oCircle As IngrGeom3D.Circle3d
    
    Dim normalX As Double
    Dim normalY As Double
    Dim normalZ As Double
    Dim xValue As Double
    Dim yValue As Double
    Dim zValue As Double
    
    xValue = 0
    yValue = 0
    zValue = 0
    normalX = 0
    normalY = 0
    normalZ = -1
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                                xValue, yValue, zValue, _
                                                                                normalX, normalY, normalZ, _
                                                                                VesselOuterDiameter / 4)
    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, oCircle)
    
    'Set the Output
    iOutput = iOutput + 1

    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set oCircle = Nothing
    
    'Place a Vertical Line along the axis of the Drum
    Dim ObjVertLine As IngrGeom3D.Line3d
    Dim startPoint(0 To 2) As Double
    Dim endPoint(0 To 2) As Double
    
    startPoint(0) = 0
    startPoint(1) = 0
    startPoint(2) = 0
    endPoint(0) = 0
    endPoint(1) = 0
    endPoint(2) = VesselOuterDiameter / 4
    Set ObjVertLine = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, startPoint(0), startPoint(1), startPoint(2), _
                                                                                                    endPoint(0), endPoint(1), endPoint(2))
    'Set the Output
    iOutput = iOutput + 1

    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVertLine
    Set ObjVertLine = Nothing

'Insert your code for Default Surface at the Supports Bottom
    'Create a line String and use it to create Default surface
    'Line String representing bottom of support
    Dim ObjBaseSupport As IngrGeom3D.Projection3d
    Dim BasePlateProjW As Double
    Dim BasePlateProjL As Double
    Dim SubtendedAng As Double
    Dim Radius As Double
    Dim SupportTopZee As Double
    Dim dlinepts(0 To 14) As Double
    Dim RadiustoEdge As Double
    Dim oLineStr1 As IngrGeom3D.LineString3d
    
    RadiustoEdge = VesselRadius + parBackingPlateThickness
    BasePlateProjW = parLengthBaseplate / 2 - parSupportLength / 2
    BasePlateProjL = parWidthBaseplate / 2 - parSupportWidth / 2
    SupportTopZee = parSupportHeight + parBasePlateThickness - parCPtoSupport
    
    RadiustoEdge = RadiustoEdge + parSupportWidth / 2 - parWidthBaseplate / 2
    SubtendedAng = Atn((parLengthBaseplate / 2) / RadiustoEdge)
    Radius = Sqr((RadiustoEdge ^ 2) + (parLengthBaseplate / 2) ^ 2)
    
    'Point 1
    dlinepts(0) = Radius * Sin(parSupportOrientation - SubtendedAng)
    dlinepts(1) = Radius * Cos(parSupportOrientation - SubtendedAng)
    dlinepts(2) = SupportTopZee - parSupportHeight - parBasePlateThickness
    
    'Point 2
    dlinepts(3) = Radius * Sin(parSupportOrientation + SubtendedAng)
    dlinepts(4) = Radius * Cos(parSupportOrientation + SubtendedAng)
    dlinepts(5) = SupportTopZee - parSupportHeight - parBasePlateThickness
    
    'Point 3
    SubtendedAng = Atn((parLengthBaseplate / 2) / (RadiustoEdge + parWidthBaseplate))
    Radius = Sqr((parLengthBaseplate / 2) ^ 2 + (RadiustoEdge + parWidthBaseplate) ^ 2)
    
    dlinepts(6) = Radius * Sin(parSupportOrientation + SubtendedAng)
    dlinepts(7) = Radius * Cos(parSupportOrientation + SubtendedAng)
    dlinepts(8) = SupportTopZee - parSupportHeight - parBasePlateThickness
    
    'Point 4
    dlinepts(9) = Radius * Sin(parSupportOrientation - SubtendedAng)
    dlinepts(10) = Radius * Cos(parSupportOrientation - SubtendedAng)
    dlinepts(11) = SupportTopZee - parSupportHeight - parBasePlateThickness
    
    'Point 5
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = dlinepts(1)
    dlinepts(14) = dlinepts(2)
    
    Set oLineStr1 = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dlinepts)
    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, oLineStr1)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set oLineStr1 = Nothing
    Set geomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

