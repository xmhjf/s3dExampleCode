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
'   Creation Date:  Monday,March 26 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has one Output:
'           Horizontal Vessel Tank
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Set m_oSymGeomHelper = New SymbolServices
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorHandler
    
    'Define End types
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
    
    Dim parVesselTantoTan  As Double
    Dim parVesselDiameter As Double
    Dim parEnd1Type As Long
    Dim parEnd2Type As Long
    Dim parInsulationThickness As Double
    Dim parEndHeadConeTopDiameter As Double
    Dim parEndHeadConeHeight As Double
    Dim parEndHeadKnuckleRadius As Double
    Dim parEndHeadDomeradius As Double
    Dim parEndHeadFlangedThick1 As Double
    Dim parEndHeadFlangedThick2 As Double
    Dim parEndHeadSphericalRadius As Double
    Dim CenterPos As IJDPosition
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Set CenterPos = New DPosition

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselTantoTan = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    parEnd1Type = arrayOfInputs(5)
    parEnd2Type = arrayOfInputs(6)
    parEndHeadConeTopDiameter = arrayOfInputs(7)    'P51
    parEndHeadConeHeight = arrayOfInputs(8)         'P50
    parEndHeadKnuckleRadius = arrayOfInputs(9)      'P51 for TORC n TORS
    parEndHeadDomeradius = arrayOfInputs(10)        'P50
    parEndHeadFlangedThick1 = arrayOfInputs(11)     'P50 for FLGD
    parEndHeadFlangedThick2 = arrayOfInputs(12)     'P51 for FLGD
    parEndHeadSphericalRadius = arrayOfInputs(13)   'P50 for TORS
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_outputColl
    
    CenterPos.Set 0, 0, 0
    stPoint.Set 0, 0, 0
    enPoint.Set parVesselTantoTan, 0, 0
        
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parVesselDiameter

    Set stPoint = Nothing
    Set enPoint = Nothing
    Set CenterPos = Nothing
    
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub





