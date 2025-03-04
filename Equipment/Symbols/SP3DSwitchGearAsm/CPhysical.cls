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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:39 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: Veena
'   Creation Date:  Friday, Sep 29 2006
'
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   22.Apr.2008     VRK                CR-120097  Need to set proper units to the cable nozzle diameter and area
'******************************************************************************

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

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim ObjFoundationPort As IJEqpFoundationPort
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim ObjSwitchGearBody As Object
    Dim ObjConduitPort As IJConduitPortOcc
    Dim ObjCablePort As IJCableNozzle
    Dim objPoint As Object
    
    Dim iOutput     As Double
    
    Dim parSwitchGearHeight As Double
    Dim parSwitchGearWidth As Double
    Dim parSwitchGearLength As Double
    Dim parXboltHole As Double
    Dim parYboltHole As Double
    
    Dim pRightCenterTopFace   As IJDPosition
    Dim pConduitPort   As IJDPosition
    Dim pCablePort     As IJDPosition

    Dim CenterPos       As IJDPosition
    Dim pBody1   As IJDPosition
    Dim oDir           As IJDVector
    Dim oRadialOrient As IJDVector
    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim oTopCircle As Object

    Set pRightCenterTopFace = New DPosition
    Set pConduitPort = New DPosition
    Set pCablePort = New DPosition
    Set CenterPos = New DPosition
    Set pBody1 = New DPosition
    Set oDir = New DVector
    Set oRadialOrient = New DVector
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parSwitchGearHeight = arrayOfInputs(2)
    parSwitchGearWidth = arrayOfInputs(3)
    parSwitchGearLength = arrayOfInputs(4)
    parXboltHole = arrayOfInputs(5)
    parYboltHole = arrayOfInputs(6)
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_outputColl

    oDir.Set 0, 0, 1
    
    CenterPos.Set 0, 0, 0
    
    'Create the SwitchGear Body
    pBody1.Set CenterPos.x + parSwitchGearLength, CenterPos.y + parSwitchGearWidth, CenterPos.z + parSwitchGearHeight
    Set ObjSwitchGearBody = PlaceBox(m_outputColl, CenterPos, pBody1)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjSwitchGearBody
    Set ObjSwitchGearBody = Nothing
    
    'The RightCenter of the top face
    pRightCenterTopFace.Set CenterPos.x, CenterPos.y + parSwitchGearWidth / 8, CenterPos.z + parSwitchGearHeight
   
       
    'Create the Conduit and Cable Ports
    pConduitPort.Set pRightCenterTopFace.x + parSwitchGearLength / 5#, pRightCenterTopFace.y, pRightCenterTopFace.z
    Set ObjConduitPort = CreateConduitNozzlePH(pConduitPort, oDir, m_outputColl, oPartFclt, 1)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
    
    pConduitPort.Set pRightCenterTopFace.x + (2# * parSwitchGearLength) / 5#, pRightCenterTopFace.y, pRightCenterTopFace.z
    Set ObjConduitPort = CreateConduitNozzlePH(pConduitPort, oDir, m_outputColl, oPartFclt, 2)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
 
    pConduitPort.Set pRightCenterTopFace.x + (3# * parSwitchGearLength) / 5#, pRightCenterTopFace.y, pRightCenterTopFace.z
    Set ObjConduitPort = CreateConduitNozzlePH(pConduitPort, oDir, m_outputColl, oPartFclt, 3)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
 
    pConduitPort.Set pRightCenterTopFace.x + (4# * parSwitchGearLength) / 5#, pRightCenterTopFace.y, pRightCenterTopFace.z
    Set ObjConduitPort = CreateConduitNozzlePH(pConduitPort, oDir, m_outputColl, oPartFclt, 4)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing

    pCablePort.Set pRightCenterTopFace.x + parSwitchGearLength / 5#, pRightCenterTopFace.y, pRightCenterTopFace.z
    Set ObjCablePort = CreateCableNozzle(pCablePort, oDir, m_outputColl, oPartFclt, 5)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCablePort
    Set ObjCablePort = Nothing
    
    pCablePort.Set pRightCenterTopFace.x + (2# * parSwitchGearLength) / 5#, pRightCenterTopFace.y, pRightCenterTopFace.z
    Set ObjCablePort = CreateCableNozzle(pCablePort, oDir, m_outputColl, oPartFclt, 6)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCablePort
    Set ObjCablePort = Nothing
    
     pCablePort.Set pRightCenterTopFace.x + (3# * parSwitchGearLength) / 5#, pRightCenterTopFace.y, pRightCenterTopFace.z
    Set ObjCablePort = CreateCableNozzle(pCablePort, oDir, m_outputColl, oPartFclt, 7)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCablePort
    Set ObjCablePort = Nothing
 
    pCablePort.Set pRightCenterTopFace.x + (4# * parSwitchGearLength) / 5#, pRightCenterTopFace.y, pRightCenterTopFace.z
    Set ObjCablePort = CreateCableNozzle(pCablePort, oDir, m_outputColl, oPartFclt, 8)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCablePort
    Set ObjCablePort = Nothing
    
  'Cable Tray with center cableport
    
    Dim ObjCableTrayPort   As IJCableTrayPortOcc
    Dim pCableTrayPort   As IJDPosition
    Set pCableTrayPort = New DPosition
    Set oDir = Nothing
    
    Set oDir = New DVector
    oDir.Set 0, 0, 1
    oRadialOrient.Set 0, -1, 0
    pCableTrayPort.Set CenterPos.x + parSwitchGearLength / 2, CenterPos.y + (3 * parSwitchGearWidth / 4), CenterPos.z + parSwitchGearHeight
  
    'Insert your code for output 6(Cable Tray Port)
    Set ObjCableTrayPort = CreateCableTrayPortPH(oPartFclt, 9, pCableTrayPort, oDir, oRadialOrient, m_outputColl)

    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCableTrayPort
    Set ObjCableTrayPort = Nothing
    Set oDir = Nothing
    
    Set oDir = New DVector
    oDir.Set 0, 0, 1
    pCablePort.Set CenterPos.x + parSwitchGearLength / 2, CenterPos.y + (3 * parSwitchGearWidth / 4), CenterPos.z + parSwitchGearHeight
    Set ObjCablePort = CreateCableNozzle(pCablePort, oDir, m_outputColl, oPartFclt, 10)

    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCablePort
    Set ObjCablePort = Nothing

    'Insert your code for output (FoundationPort1)
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(1 To 3) As Double
    Dim dXaxis(1 To 3) As Double
    Dim dZaxis(1 To 3) As Double
    
    'Coordinate sysytem is at the bottom of support.
    '
    '               |-------|
    '   Y           |       |
    '   ^           |       |
    '   |           |       |
    '   |           |       |
    '   |           |       |
    '   -----> X    |       |
    '   Symbol CS   |       |
    '               |       |
    '               |   X   |
    '               |   ^   |
    '               |   |   |
    '               |   |   |
    '               |---|---|----> Y
    '                   Port CS
    
    dOrigin(1) = 0#
    dOrigin(2) = 0#
    dOrigin(3) = 0#
    
    dXaxis(1) = 0#
    dXaxis(2) = 1#
    dXaxis(3) = 0#
    
    dZaxis(1) = 0#
    dZaxis(2) = 0#
    dZaxis(3) = -1#
    
    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_outputColl.ResourceManager)


    Dim holes() As Variant
    Call ObjFoundationPort.GetHoles(holes())
    
    holes(0, 1) = parXboltHole
    holes(0, 2) = -(parSwitchGearWidth - parYboltHole)
    holes(1, 1) = parXboltHole
    holes(1, 2) = -parYboltHole
    holes(2, 1) = parSwitchGearLength - parXboltHole
    holes(2, 2) = -parYboltHole
    holes(3, 1) = parSwitchGearLength - parXboltHole
    holes(3, 2) = -(parSwitchGearWidth - parYboltHole)
    Call ObjFoundationPort.PutCS(dOrigin(1), dOrigin(2), dOrigin(3), _
                             dXaxis(1), dXaxis(2), dXaxis(3), _
                            dZaxis(1), dZaxis(2), dZaxis(3))
    
    Call ObjFoundationPort.SetHoles(holes)

    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
  
 
    'Circle to indicate the cableport
    iOutput = iOutput + 1
    Set oTopCircle = geomFactory.Circles3d.CreateBy3Points(m_outputColl.ResourceManager, (parSwitchGearLength / 2) - 0.01, (3 * parSwitchGearWidth / 4), parSwitchGearHeight, _
                                                                   (parSwitchGearLength / 2), (3 * parSwitchGearWidth / 4) + 0.01, parSwitchGearHeight, _
                                                                  (parSwitchGearLength / 2) + 0.01, (3 * parSwitchGearWidth / 4), parSwitchGearHeight)


   m_outputColl.AddOutput arrayOfOutputs(iOutput), oTopCircle
   Set oTopCircle = Nothing
  
    
    'Create salient points on each face of the body
    
    CenterPos.Set 0, parSwitchGearWidth / 2, parSwitchGearHeight / 2 'FrontFace Center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing

    CenterPos.Set parSwitchGearLength / 2, parSwitchGearWidth, parSwitchGearHeight / 2 'Left Face Center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing

    CenterPos.Set parSwitchGearLength, parSwitchGearWidth / 2, parSwitchGearHeight / 2 'Back FaceCenter
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    CenterPos.Set parSwitchGearLength / 2, 0, parSwitchGearHeight / 2 'Right Face Center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
     
    CenterPos.Set parSwitchGearLength / 2, parSwitchGearWidth / 2, parSwitchGearHeight 'Top FaceCenter
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
     
    CenterPos.Set parSwitchGearLength / 2, parSwitchGearWidth / 2, 0 'Bottom Face center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing

 'Create line by two pints:
            
        
    Dim Points(0 To 23) As Double
    Dim line As Object
    
    Points(0) = 0
    Points(1) = 0
    Points(2) = 0
    
    Points(3) = 0
    Points(4) = parSwitchGearWidth
    Points(5) = 0
    
    Points(6) = 0
    Points(7) = parSwitchGearWidth
    Points(8) = parSwitchGearHeight
    
    Points(9) = 0
    Points(10) = 0
    Points(11) = parSwitchGearHeight
    
    Points(12) = parSwitchGearLength
    Points(13) = 0
    Points(14) = 0
    
    Points(15) = parSwitchGearLength
    Points(16) = parSwitchGearWidth
    Points(17) = 0
    
    Points(18) = parSwitchGearLength
    Points(19) = parSwitchGearWidth
    Points(20) = parSwitchGearHeight
    
    Points(21) = parSwitchGearLength
    Points(22) = 0
    Points(23) = parSwitchGearHeight
       
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, Points(0), Points(1), Points(2), Points(3), Points(4), Points(5))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, Points(3), Points(4), Points(5), Points(6), Points(7), Points(8))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, Points(6), Points(7), Points(8), Points(9), Points(10), Points(11))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, Points(9), Points(10), Points(11), Points(0), Points(1), Points(2))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, Points(0), Points(1), Points(2), Points(12), Points(13), Points(14))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, Points(12), Points(13), Points(14), Points(21), Points(22), Points(23))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, Points(21), Points(22), Points(23), Points(9), Points(10), Points(11))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, Points(18), Points(19), Points(20), Points(6), Points(7), Points(8))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
        
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, Points(18), Points(19), Points(20), Points(15), Points(16), Points(17))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, Points(3), Points(4), Points(5), Points(15), Points(16), Points(17))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
       
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, Points(15), Points(16), Points(17), Points(12), Points(13), Points(14))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, Points(21), Points(22), Points(23), Points(18), Points(19), Points(20))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
  
   Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
     
End Sub

Private Function CreateCableNozzle(oBasePt As AutoMath.DPosition, oAxis As AutoMath.DVector, ByVal objOutputColl As Object, ByRef oPartFclt As PartFacelets.IJDPart, dNozzleIndex As Long) As IJCableNozzle
        Const METHOD = "CreateCableNozzle:"
 
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim iDistribPort As IJDistribPort
    Dim NozzlePHFactory As NozzlePHFactory
    Dim ObjCableNozzle As CableNozzle
    
    Set NozzlePHFactory = New NozzlePHFactory
    Set ObjCableNozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, dNozzleIndex, _
                                                             objOutputColl.ResourceManager)
                                
    Set iLogicalDistPort = ObjCableNozzle
    Set iDistribPort = ObjCableNozzle
    iLogicalDistPort.SetCenterLocation oBasePt
    iDistribPort.SetDirectionVector oAxis
    Set CreateCableNozzle = ObjCableNozzle
   
    Set NozzlePHFactory = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set ObjCableNozzle = Nothing

    Exit Function
    
ErrHandler:
  ReportUnanticipatedError MODULE, METHOD
End Function







