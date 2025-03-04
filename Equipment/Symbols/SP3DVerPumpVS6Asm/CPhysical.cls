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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:00 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RUK
'   Creation Date:  Monday, September 03 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Nine Outputs:
'      In which Default Surface is created at the bottom of the Base Plate.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    PI = 4 * Atn(1)
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parBaseBotToDischarge As Double
    Dim parBaseBotToSuction As Double
    Dim parHeadLength As Double
    Dim parMotorLength As Double
    Dim parBarrelLength As Double
    Dim parBaseThickness As Double
    Dim parBarrelDiameter As Double
    Dim parHeadDiameter As Double
    Dim parMotorDiameter As Double
    Dim parBasePlateWidth As Double
    Dim parDischProjection As Double
    Dim parSuctionProjection As Double
    Dim parPullEnvolopeLen As Double
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBaseBotToDischarge = arrayOfInputs(2)    'A
    parBaseBotToSuction = arrayOfInputs(3)      'B
    parHeadLength = arrayOfInputs(4)            'C
    parMotorLength = arrayOfInputs(5)           'D
    parBarrelLength = arrayOfInputs(6)          'E
    parBaseThickness = arrayOfInputs(7)         'F
    parBarrelDiameter = arrayOfInputs(8)        'G
    parHeadDiameter = arrayOfInputs(9)          'H
    parMotorDiameter = arrayOfInputs(10)        'J
    parBasePlateWidth = arrayOfInputs(11)       'K
    parDischProjection = arrayOfInputs(12)      'L
    parSuctionProjection = arrayOfInputs(13)    'M
    parPullEnvolopeLen = arrayOfInputs(14)      'P
    parXBoltHole = arrayOfInputs(15)
    parYBoltHole = arrayOfInputs(16)
    
    'Origin is at PP1
    iOutput = 0

    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition
    Dim oVector As AutoMath.DVector
    Dim oCompStr As IngrGeom3D.ComplexString3d
    Dim oLine As Object
    Dim oMajor As AutoMath.DPosition
    Dim oArc As Object
    
    Dim dMMRatio As Double
    Dim iCount As Integer
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oCenter = New DPosition
    Set oVector = New DVector
    Set oMajor = New DPosition
    
    'Create the Barrel or Pipe
    Dim objBarrel As Object
    Dim oBarrColl As Collection
    
    Set oCompStr = New ComplexString3d
    Set oBarrColl = New Collection
    
    'Line 1
    oStPoint.Set 0, 0, 0
    oEnPoint.Set parBarrelDiameter / 2, 0, 0
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oBarrColl.Add oLine
    Set oLine = Nothing
    
    'Line 2
    oStPoint.Set parBarrelDiameter / 2, 0, 0
    oEnPoint.Set parBarrelDiameter / 2, 0, -0.9 * parBarrelLength
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oBarrColl.Add oLine
    Set oLine = Nothing
    
    'Arc
    oCenter.Set 0, 0, -0.9 * parBarrelLength
    oMajor.Set parBarrelDiameter / 2, 0, 0
    oVector.Set 0, 1, 0
    dMMRatio = (0.1 * parBarrelLength) / (parBarrelDiameter / 2)
    Set oArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                        Nothing, oCenter.x, oCenter.y, oCenter.z, _
                        oVector.x, oVector.y, oVector.z, _
                        oMajor.x, oMajor.y, oMajor.z, _
                        dMMRatio, 0, PI / 2)
    oBarrColl.Add oArc
    Set oArc = Nothing
    
    'Non-persistant Complex String
    oStPoint.Set 0, 0, 0
    Set oCompStr = PlaceTrCString(oStPoint, oBarrColl)
    oVector.Set 0, 0, 1
    oCenter.Set 0, 0, 0
    Set objBarrel = PlaceRevolution(m_OutputColl, oCompStr, oVector, oCenter, _
                                                2 * PI, True)
        
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBarrel
    Set objBarrel = Nothing
    
    For iCount = 1 To oBarrColl.Count
        oBarrColl.Remove 1
    Next iCount
    Set oBarrColl = Nothing
    Set oCompStr = Nothing
    
    'Create the Baseplate (Output 3)
    Dim objBasePlateColl As Collection
    Dim oTopSurPts(0 To 3) As IJDPosition
    Dim oBotSurPts(0 To 3) As IJDPosition
        
    Set objBasePlateColl = New Collection
    For iCount = 0 To 3
        Set oTopSurPts(iCount) = New DPosition
        Set oBotSurPts(iCount) = New DPosition
    Next iCount
    
    oTopSurPts(0).Set parBasePlateWidth / 2, parBasePlateWidth / 2, parBaseThickness
    oTopSurPts(1).Set -oTopSurPts(0).x, oTopSurPts(0).y, oTopSurPts(0).z
    oTopSurPts(2).Set -oTopSurPts(0).x, -oTopSurPts(0).y, oTopSurPts(0).z
    oTopSurPts(3).Set oTopSurPts(0).x, -oTopSurPts(0).y, oTopSurPts(0).z
    
    oBotSurPts(0).Set oTopSurPts(0).x, oTopSurPts(0).y, 0
    oBotSurPts(1).Set -oBotSurPts(0).x, oBotSurPts(0).y, oBotSurPts(0).z
    oBotSurPts(2).Set -oBotSurPts(0).x, -oBotSurPts(0).y, oBotSurPts(0).z
    oBotSurPts(3).Set oBotSurPts(0).x, -oBotSurPts(0).y, oBotSurPts(0).z
    
    Set objBasePlateColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopSurPts, _
                                    oBotSurPts)
    'Set the Output
    iOutput = iOutput + 1
    For iCount = 1 To objBasePlateColl.Count
        m_OutputColl.AddOutput "BasePlate_", objBasePlateColl.Item(iCount)
    Next iCount
    
    'Create the Default Surface (Output 4)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBasePlateColl.Item(2)
    
    For iCount = 1 To objBasePlateColl.Count
        objBasePlateColl.Remove 1
    Next iCount
    Set objBasePlateColl = Nothing
    For iCount = 0 To 3
        Set oTopSurPts(iCount) = Nothing
        Set oBotSurPts(iCount) = Nothing
    Next iCount
    
    'Create Edges for Baseplate
    Dim objEdgeColl As Collection
    Set objEdgeColl = New Collection
    oStPoint.Set parBasePlateWidth / 2, parBasePlateWidth / 2, parBaseThickness
    oEnPoint.Set -parBasePlateWidth / 2, -parBasePlateWidth / 2, 0
    Set objEdgeColl = CreateBoxEdges(m_OutputColl, oStPoint, oEnPoint)
    'Set the Output
    iOutput = iOutput + 1
    For iCount = 1 To objEdgeColl.Count
        m_OutputColl.AddOutput "Edges_", objEdgeColl.Item(iCount)
    Next iCount
    For iCount = 1 To objEdgeColl.Count
        objEdgeColl.Remove 1
    Next iCount
    Set objEdgeColl = Nothing
    
    'Create the Points on each face of the Baseplate
    Dim objPointColl As Collection
    Set objPointColl = New Collection
    
    oStPoint.Set parBasePlateWidth / 2, parBasePlateWidth / 2, parBaseThickness
    oEnPoint.Set -parBasePlateWidth / 2, -parBasePlateWidth / 2, 0
    Set objPointColl = CreatePointsOnBoxFaces(m_OutputColl, oStPoint, oEnPoint)
    'Set the Output
    iOutput = iOutput + 1
    For iCount = 1 To objPointColl.Count
        m_OutputColl.AddOutput "Points_", objPointColl.Item(iCount)
    Next iCount
    For iCount = 1 To objPointColl.Count
        objPointColl.Remove 1
    Next iCount
    Set objPointColl = Nothing
    
    'Create the Head or Driver Stand
    Dim objHead As Object
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, 0, parHeadLength
    Set objHead = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            parHeadDiameter, True)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHead
    Set objHead = Nothing
    
    'Create the Motor
    Dim objMotor As Object
    Dim oMotColl As Collection
    
    Set oCompStr = New ComplexString3d
    Set oMotColl = New Collection
    
    'Line 1
    oStPoint.Set 0, 0, parHeadLength
    oEnPoint.Set parMotorDiameter / 2, 0, parHeadLength
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oMotColl.Add oLine
    Set oLine = Nothing
    
    'Line 2
    oStPoint.Set parMotorDiameter / 2, 0, parHeadLength
    oEnPoint.Set parMotorDiameter / 2, 0, parHeadLength + 0.9 * parMotorLength
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oMotColl.Add oLine
    Set oLine = Nothing
    
    'Arc
    Set oMajor = New DPosition
    
    oCenter.Set 0, 0, parHeadLength + 0.9 * parMotorLength
    oMajor.Set parMotorDiameter / 2, 0, 0
    oVector.Set 0, -1, 0
    dMMRatio = (0.1 * parMotorLength) / (parMotorDiameter / 2)
    Set oArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                        Nothing, oCenter.x, oCenter.y, oCenter.z, _
                        oVector.x, oVector.y, oVector.z, _
                        oMajor.x, oMajor.y, oMajor.z, _
                        dMMRatio, 0, PI / 2)
    oMotColl.Add oArc
    Set oArc = Nothing
    
    'Non-persistant Complex String
    oStPoint.Set 0, 0, parHeadLength
    Set oCompStr = PlaceTrCString(oStPoint, oMotColl)
    
    oVector.Set 0, 0, 1
    oCenter.Set 0, 0, parHeadLength
    Set objMotor = PlaceRevolution(m_OutputColl, oCompStr, oVector, oCenter, _
                                                2 * PI, True)
    'Set the Output
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMotor
    Set oMajor = Nothing
    For iCount = 1 To oMotColl.Count
        oMotColl.Remove 1
    Next iCount
    Set oCompStr = Nothing
    Set oMotColl = Nothing
    Set objMotor = Nothing
    
    'Remove the References
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenter = Nothing
    Set oVector = Nothing
    Set oGeomFactory = Nothing
   
    'Create Equipment Foundation Port at the Bottom of Base Plate
    '
    '                      |-------|
    '   Y                  |       |
    '   ^                  |       |
    '   |                  |   X   |
    '   |                  |   ^   |
    '   |                  |   |   |
    '   -----> X           |   |   |Port CS
    '   Symbol CS          |   |----> Y
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |-------|
    
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
   'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = 0
    
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
    Set NozzlePHFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
