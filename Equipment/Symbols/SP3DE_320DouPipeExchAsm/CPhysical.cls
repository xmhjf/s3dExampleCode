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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:03:05 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RUK
'   Creation Date:  Tuesday, Feb 27 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Nine Outputs:
'      In which One Default Surface is created at Origin.
'      Two Rectangular Foundation Ports are created at the bottom of Supports
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

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

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim ObjPipe1    As Object
    Dim ObjPipe2    As Object
    Dim ObjBlock    As Object
    Dim ObjSupport1 As Object
    Dim ObjSupport2 As Object
       
    Dim parBlockEndFromPP As Double         'P1
    Dim parPipeEndFromPP As Double          'P2
    Dim parPipeCentoCen As Double           'P3
    Dim parPipeDiameter As Double           'P4
    Dim parBlockLength As Double            'P5
    Dim parBlockHeight As Double            'P6
    Dim parBlockWidth As Double             'P7
    Dim parBundlePullingLength As Double    'P8
    Dim parSupport1CenFromPP As Double      'P9
    Dim parSupportCentoCen As Double        'P10
    Dim parSupportHeight As Double          'P11
    Dim parSupportThickness As Double       'P12
    Dim parSupportWidth As Double           'P13
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double
   
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBlockEndFromPP = arrayOfInputs(2)
    parPipeEndFromPP = arrayOfInputs(3)
    parPipeCentoCen = arrayOfInputs(4)
    parPipeDiameter = arrayOfInputs(5)
    parBlockLength = arrayOfInputs(6)
    parBlockHeight = arrayOfInputs(7)
    parBlockWidth = arrayOfInputs(8)
    parBundlePullingLength = arrayOfInputs(9)
    parSupport1CenFromPP = arrayOfInputs(10)
    parSupportCentoCen = arrayOfInputs(11)
    parSupportHeight = arrayOfInputs(12)
    parSupportThickness = arrayOfInputs(13)
    parSupportWidth = arrayOfInputs(14)
    parXBoltHole = arrayOfInputs(15)
    parYBoltHole = arrayOfInputs(16)
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    'Origin is at DP/PP1
    'Insert your code for Default Surface (Default surface at origin)
    Dim objCircle As IngrGeom3D.Circle3d
    Dim ObjDefaultSurface As Object

    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                    0, 0, 0, _
                                                    1, 0, 0, _
                                                    0.8 * parPipeDiameter / 2)
    'Create persistent default surface plane - the plane can mate ---
    Set ObjDefaultSurface = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set objCircle = Nothing
      
    'Insert your code for Pipe1
    stPoint.Set -parPipeEndFromPP, 0, parPipeCentoCen / 2
    enPoint.Set parBlockEndFromPP - parBlockLength, 0, parPipeCentoCen / 2
    
    Set ObjPipe1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parPipeDiameter, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipe1
    Set ObjPipe1 = Nothing
    
    'Insert your code for Pipe2
    stPoint.Set -parPipeEndFromPP, 0, -parPipeCentoCen / 2
    enPoint.Set parBlockEndFromPP - parBlockLength, 0, -parPipeCentoCen / 2
    
    Set ObjPipe2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parPipeDiameter, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipe2
    Set ObjPipe2 = Nothing
    
    'Insert your code for Block
    stPoint.Set parBlockEndFromPP - parBlockLength, parBlockWidth / 2, parBlockHeight / 2
    enPoint.Set parBlockEndFromPP, -parBlockWidth / 2, -parBlockHeight / 2
    
    Set ObjBlock = PlaceBox(m_OutputColl, stPoint, enPoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBlock
    Set ObjBlock = Nothing
    
    'Inser your code for creating the edges for Block
    Dim iCount As Integer
    Dim Points(0 To 23) As Double
    Dim ObjEdgeColl As New Collection
    
    stPoint.Set parBlockEndFromPP - parBlockLength, 0, 0
    
    'Points at corners of block
    
    '  0,1,2-----------------12,13,14
    '       '               '
    '       '  6,7,8-----------------18,19,20
    '       '       '        '      '
    '       '       '        '      '
    '       '       '        '      '
    '  3,4,5--------'---------15,16,17
    '               '               '
    '       9,10,11 -----------------21,22,23
    
    For iCount = 0 To 12 Step 12
        Points(0 + iCount) = stPoint.x
        Points(1 + iCount) = stPoint.y + parBlockWidth / 2
        Points(2 + iCount) = stPoint.z + parBlockHeight / 2

        Points(3 + iCount) = stPoint.x
        Points(4 + iCount) = stPoint.y + parBlockWidth / 2
        Points(5 + iCount) = stPoint.z - parBlockHeight / 2
        
        Points(6 + iCount) = stPoint.x
        Points(7 + iCount) = stPoint.y - parBlockWidth / 2
        Points(8 + iCount) = stPoint.z + parBlockHeight / 2
        
        Points(9 + iCount) = stPoint.x
        Points(10 + iCount) = stPoint.y - parBlockWidth / 2
        Points(11 + iCount) = stPoint.z - parBlockHeight / 2
        
        stPoint.x = stPoint.x + parBlockLength
    Next iCount
    
    'Edges
    'Edge1 Point1 to Point2
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                Points(0), Points(1), Points(2), _
                                Points(3), Points(4), Points(5))
    'Edge2 Point1 to Point3
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                Points(0), Points(1), Points(2), _
                                Points(6), Points(7), Points(8))
    'Edge3 Point1 to Point5
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                Points(0), Points(1), Points(2), _
                                Points(12), Points(13), Points(14))
    'Edge4 Point4 to Point2
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                Points(9), Points(10), Points(11), _
                                Points(3), Points(4), Points(5))
    'Edge5 Point4 to Point3
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                Points(9), Points(10), Points(11), _
                                Points(6), Points(7), Points(8))
    'Edge6 Point4 to Point8
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                Points(9), Points(10), Points(11), _
                                Points(21), Points(22), Points(23))
    'Edge7 Point6 to Point2
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                Points(15), Points(16), Points(17), _
                                Points(3), Points(4), Points(5))
    'Edge8 Point6 to Point5
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                Points(15), Points(16), Points(17), _
                                Points(12), Points(13), Points(14))
    'Edge9 Point6 to Point8
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                Points(15), Points(16), Points(17), _
                                Points(21), Points(22), Points(23))
    'Edge10 Point7 to Point3
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                Points(18), Points(19), Points(20), _
                                Points(6), Points(7), Points(8))
    'Edge11 Point7 to Point5
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                Points(18), Points(19), Points(20), _
                                Points(12), Points(13), Points(14))
    'Edge12 Point7 to Point8
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                Points(18), Points(19), Points(20), _
                                Points(21), Points(22), Points(23))
    'Set the Output
    iOutput = iOutput + 1
    For iCount = 1 To ObjEdgeColl.Count
        m_OutputColl.AddOutput "Edges_", ObjEdgeColl(iCount)
    Next iCount
    Set ObjEdgeColl = Nothing
    
    'Creating the Points on each surface of the Block
    Dim ObjPointColl As Collection
    Set ObjPointColl = New Collection
    
    stPoint.Set parBlockEndFromPP - parBlockLength / 2, 0, 0
    
    'Points on Right and Left surfaces
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                            stPoint.x + parBlockLength / 2, stPoint.y, stPoint.z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                            stPoint.x - parBlockLength / 2, stPoint.y, stPoint.z)
    'Points on Front and Back surfaces
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                            stPoint.x, stPoint.y - parBlockWidth / 2, stPoint.z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                            stPoint.x, stPoint.y + parBlockWidth / 2, stPoint.z)
    'Points on Top and Bottom surfaces
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                            stPoint.x, stPoint.y, stPoint.z + parBlockHeight / 2)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                            stPoint.x, stPoint.y, stPoint.z - parBlockHeight / 2)
    'Set the output
    iOutput = iOutput + 1
    For iCount = 1 To ObjPointColl.Count
        m_OutputColl.AddOutput "Points_", ObjPointColl(iCount)
    Next iCount
    Set ObjPointColl = Nothing
    
    'Insert your code for creating Supports
    If (parSupportThickness > 0) Then
        If (parSupport1CenFromPP > 0) Then
            stPoint.Set parSupport1CenFromPP - parSupportThickness / 2, _
                            parSupportWidth / 2, _
                            parSupportHeight / 2
            enPoint.Set parSupport1CenFromPP + parSupportThickness / 2, _
                            -parSupportWidth / 2, _
                            -parSupportHeight / 2
    
            Set ObjSupport1 = PlaceBox(m_OutputColl, stPoint, enPoint)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "Supports_", ObjSupport1
            Set ObjSupport1 = Nothing
        End If
            
        If (parSupport1CenFromPP > 0) And (parSupportCentoCen > 0) Then
            stPoint.Set parSupport1CenFromPP + parSupportCentoCen - parSupportThickness / 2, _
                            parSupportWidth / 2, _
                            parSupportHeight / 2
            enPoint.Set parSupport1CenFromPP + parSupportCentoCen + parSupportThickness / 2, _
                            -parSupportWidth / 2, _
                            -parSupportHeight / 2
    
            Set ObjSupport2 = PlaceBox(m_OutputColl, stPoint, enPoint)
            ' Set the output
            m_OutputColl.AddOutput "Supports_", ObjSupport2
            Set ObjSupport2 = Nothing
        End If
    End If
      
    'Create Equipment Foundation Ports  at the bottom of Supports
    'Coordinate system is at the bottom of support.
    '
    '                      |-------|
    '   Y                  |       |
    '   ^                  |       |
    '   |                  |       |
    '   |                  |       |
    '   |                  |       |
    '   -----> X           |       |
    '   Symbol CS          |       |
    '                      |       |
    '                      |   X   |
    '                      |   ^   |
    '                      |   |   |
    '                      |   |   |
    '                      |---|---|----> Y
    '                   Port CS
    
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
    
    'Create FoundationPort1 at the bottom of Support1
    'The origin of the port is taken to be at the center point of the support base.
    dOrigin(0) = parSupport1CenFromPP
    dOrigin(1) = 0
    dOrigin(2) = -parSupportHeight / 2
    
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#

    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_OutputColl.ResourceManager)
    Dim holes() As Variant
    Call ObjFoundationPort.GetHoles(holes())
    
    holes(0, 1) = -parXBoltHole
    holes(0, 2) = -parYBoltHole
    holes(1, 1) = parXBoltHole
    holes(1, 2) = -parYBoltHole
    holes(2, 1) = parXBoltHole
    holes(2, 2) = parYBoltHole
    holes(3, 1) = -parXBoltHole
    holes(3, 2) = parYBoltHole
    
    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                             dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
    Call ObjFoundationPort.SetHoles(holes)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    
    'Create FoundationPort2 at the bottom Support2
    'The origin of the port is taken to be at the ccenter point of the support base.
    dOrigin(0) = parSupport1CenFromPP + parSupportCentoCen
    dOrigin(1) = 0
    dOrigin(2) = -parSupportHeight / 2
    
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#

    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort2", _
                                                    False, m_OutputColl.ResourceManager)
    
    Call ObjFoundationPort.GetHoles(holes())

    holes(0, 1) = -parXBoltHole
    holes(0, 2) = -parYBoltHole
    holes(1, 1) = parXBoltHole
    holes(1, 2) = -parYBoltHole
    holes(2, 1) = parXBoltHole
    holes(2, 2) = parYBoltHole
    holes(3, 1) = -parXBoltHole
    holes(3, 2) = parYBoltHole
    
    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))

    Call ObjFoundationPort.SetHoles(holes)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
