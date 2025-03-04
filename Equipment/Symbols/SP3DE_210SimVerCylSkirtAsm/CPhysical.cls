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
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         VRK
'   Creation Date:  Wednesday,March 14 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has four Outputs:
'      In which Default Surface is created at the Origin.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
     Set m_oSymGeomHelper = New SymbolServices
     PI = 4 * Atn(1)
     
     Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    'Define Vessel Head types
    Const EndType_2TO1 = 1
    Const EndType_CONE = 2
    Const EndType_DOME = 3
    Const EndType_FnD = 4
    Const EndType_FLAT = 5
    Const EndType_FLGD = 6
    Const EndType_HEMI = 7
    Const EndType_NONE = 8
    Const EndType_TORC = 9
    Const EndType_TORS = 10
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double
    Dim objVesselBody As Object
    
    Dim parEnd1Type As Long
    Dim parEnd2Type As Long
    Dim parVesselTantoTan As Double
    Dim parVesselDiameter As Double
    Dim parSkirtTopDiameter As Double
    Dim parSkirtBottomDiameter As Double
    Dim parSkirtHeight As Double
    Dim parSupportReferencePoint As Double
    Dim parEndHeadConeTopDiameter As Double
    Dim parEndHeadConeHeight As Double
    Dim parEndHeadKnuckleRadius As Double
    Dim parEndHeadDomeradius As Double
    Dim parEndHeadFlangedThick1 As Double
    Dim parEndHeadFlangedThick2 As Double
    Dim parEndHeadSphericalRadius As Double
    Dim parInsulationThickness As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parEnd1Type = arrayOfInputs(2)
    parEnd2Type = arrayOfInputs(3)
    parVesselTantoTan = arrayOfInputs(4)            'P2
    parVesselDiameter = arrayOfInputs(5)            'P3
    parSkirtTopDiameter = arrayOfInputs(6)          'P5
    parSkirtBottomDiameter = arrayOfInputs(7)       'P6
    parSkirtHeight = arrayOfInputs(8)               'P4
    parSupportReferencePoint = arrayOfInputs(9)     'P7
    parEndHeadConeTopDiameter = arrayOfInputs(10)   'P51
    parEndHeadConeHeight = arrayOfInputs(11)        'P50
    parEndHeadKnuckleRadius = arrayOfInputs(12)     'P51 for TORC and TORS
    parEndHeadDomeradius = arrayOfInputs(13)        'P50 for Dome
    parEndHeadFlangedThick1 = arrayOfInputs(14)     'P50 for FLGD
    parEndHeadFlangedThick2 = arrayOfInputs(15)     'P51 for FLGD
    parEndHeadSphericalRadius = arrayOfInputs(16)   'P50
    parInsulationThickness = arrayOfInputs(17)
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    'Origin is at DP2 of Top plane of Skirt.
    'Place Default Surface
    Dim geomFactory As GeometryFactory
    Set geomFactory = New GeometryFactory
    Dim ObjBotDefaultSurface As IngrGeom3D.Plane3d
    Dim oCircle As IngrGeom3D.Circle3d
    Dim xValue As Double
    Dim yValue As Double
    Dim zValue As Double
    Dim normalX As Double
    Dim normalY As Double
    Dim normalZ As Double
    xValue = 0
    yValue = 0
    zValue = 0

    normalX = 0
    normalY = 0
    normalZ = -1
    Dim dDefaultSurfaceDiameter As Double
    
    If (CmpDblGreaterthan(parSupportReferencePoint, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSkirtTopDiameter, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSkirtBottomDiameter, LINEAR_TOLERANCE)) Then
        dDefaultSurfaceDiameter = parSkirtTopDiameter
    Else
        dDefaultSurfaceDiameter = 0.25 * parVesselDiameter
   End If
    
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                xValue, yValue, zValue, _
                                                normalX, normalY, normalZ, _
                                                dDefaultSurfaceDiameter / 2)
    Set ObjBotDefaultSurface = geomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, oCircle)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBotDefaultSurface
    Set ObjBotDefaultSurface = Nothing
    Set oCircle = Nothing
    Set geomFactory = Nothing
    
    'Insert your code for Vessel Body
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    stPoint.Set 0, 0, 0
    enPoint.Set 0, 0, parVesselTantoTan
    
    Set objVesselBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVesselBody
    Set objVesselBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    'Insert your code for Truncated Cone Skirt
    If (CmpDblGreaterthan(parSkirtHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSkirtTopDiameter, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSkirtBottomDiameter, LINEAR_TOLERANCE)) Then
        Dim objSkirt As Object
        Dim baseCenter As New AutoMath.DPosition
        Dim topCenter As New AutoMath.DPosition
        topCenter.Set 0, 0, -(parSupportReferencePoint - parSkirtHeight) 'P7-P4
        baseCenter.Set 0, 0, -parSupportReferencePoint
    
        Set objSkirt = PlaceCone(m_OutputColl, baseCenter, topCenter, parSkirtBottomDiameter / 2, parSkirtTopDiameter / 2, False)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "SupportSkirt", objSkirt
        Set objSkirt = Nothing
        Set baseCenter = Nothing
        Set topCenter = Nothing
    End If
    
    '==========================================
    'Construction of Equipment Foundation Port
    '==========================================
    'Place Circular FoundationPort
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
    
    'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = -parSupportReferencePoint
    
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#
    
    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_OutputColl.ResourceManager)

    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub






