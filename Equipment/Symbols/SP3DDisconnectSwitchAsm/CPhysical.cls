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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: Veena
'   Creation Date:  Monday, oct 16 2006
'
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************


Option Explicit
Private m_oSymGeomHelper As IJSymbolGeometryHelper

Private Const MODULE = "CPhysical:" 'Used for error messages

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
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double

    Dim parElecEquipmentHeight As Double
    Dim parElecEquipmentWidth As Double
    Dim parElecEquipmentLength As Double
    Dim parSwitchLeverHeight As Double
    Dim parSwitchLeverWidth As Double
    Dim parSwitchLeverLength As Double
    
    Dim ObjDisContSwitchBody As Object
    Dim ObjProjectedKnob As Object
    Dim ObjConduitPort As IJConduitPortOcc
    Dim ObjCablePort As IJCableNozzle
    Dim objPoint As Object
     
    Dim pConduitPort   As IJDPosition
    Dim pCablePort     As IJDPosition
    
    Dim CenterPos       As IJDPosition
    Dim pBody1   As IJDPosition
    Dim oDir           As IJDVector
    Dim oRadialOrient As IJDVector
    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim pBodyCorner1   As IJDPosition
    Dim pBodyCorner2   As IJDPosition
    
    Set pConduitPort = New DPosition
    Set pCablePort = New DPosition
    Set pBodyCorner1 = New DPosition
    Set pBodyCorner2 = New DPosition
    Set oDir = New DVector
    Set oRadialOrient = New DVector
    Set CenterPos = New DPosition
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parElecEquipmentHeight = arrayOfInputs(2)
    parElecEquipmentWidth = arrayOfInputs(3)
    parElecEquipmentLength = arrayOfInputs(4)
    parSwitchLeverHeight = arrayOfInputs(5)
    parSwitchLeverWidth = arrayOfInputs(6)
    parSwitchLeverLength = arrayOfInputs(7)

    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_outputColl

    'Taking Leftface bottom line as origin
    pBodyCorner1.Set parElecEquipmentWidth, -parElecEquipmentLength / 2, 0
    pBodyCorner2.Set pBodyCorner1.x - parElecEquipmentWidth, pBodyCorner1.y + parElecEquipmentLength, pBodyCorner1.z + parElecEquipmentHeight

    'Create the DisConnect Switch Body
    Set ObjDisContSwitchBody = PlaceBox(m_outputColl, pBodyCorner1, pBodyCorner2)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDisContSwitchBody
    Set ObjDisContSwitchBody = Nothing
    
    'Insert your code for outputs 2(Knob of Control Switch)
    Dim KnobPoints(0 To 17) As Double
    Dim line As Object
    'Point 1
    KnobPoints(0) = parElecEquipmentWidth
    KnobPoints(1) = -parElecEquipmentLength / 4
    KnobPoints(2) = parElecEquipmentHeight / 2
    'Point2
    KnobPoints(3) = KnobPoints(0)
    KnobPoints(4) = 0
    KnobPoints(5) = parElecEquipmentHeight / 2 + parSwitchLeverHeight / 2
    'Point3
    KnobPoints(6) = KnobPoints(0)
    KnobPoints(7) = KnobPoints(1) + parSwitchLeverLength
    KnobPoints(8) = KnobPoints(5)
    'Point4
    KnobPoints(9) = KnobPoints(0)
    KnobPoints(10) = KnobPoints(7)
    KnobPoints(11) = KnobPoints(8) - parSwitchLeverHeight
    'Point5
    KnobPoints(12) = KnobPoints(0)
    KnobPoints(13) = 0
    KnobPoints(14) = KnobPoints(11)
    'Point6
    KnobPoints(15) = KnobPoints(0)
    KnobPoints(16) = KnobPoints(1)
    KnobPoints(17) = KnobPoints(2)

    oDir.Set 1, 0, 0
    Set line = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, KnobPoints)

    Set ObjProjectedKnob = PlaceProjection(m_outputColl, line, oDir, parSwitchLeverWidth, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjProjectedKnob
    Set ObjProjectedKnob = Nothing
    
    oDir.Set 0, 0, -1
    'Create the Conduit and Cable Ports
    pConduitPort.Set parElecEquipmentWidth / 2, 0, parElecEquipmentHeight
    Set ObjConduitPort = CreateConduitNozzlePH(pConduitPort, oDir, m_outputColl, oPartFclt, 1)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
   
    pCablePort.Set pConduitPort.x, pConduitPort.y, pConduitPort.z
    Set ObjCablePort = CreateCableNozzle(pCablePort, oDir, m_outputColl, oPartFclt, 2)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCablePort
    Set ObjCablePort = Nothing

    
    'Create salient points on each face of the body
    CenterPos.Set parElecEquipmentWidth / 2, parElecEquipmentLength / 2, parElecEquipmentHeight / 2 'BackFace Center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing

    CenterPos.Set parElecEquipmentWidth / 2, -parElecEquipmentLength / 2, parElecEquipmentHeight / 2 'Front Face Center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing

    CenterPos.Set parElecEquipmentWidth / 2, 0, parElecEquipmentHeight  'TopFaceCenter
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    CenterPos.Set parElecEquipmentWidth / 2, 0, 0  'Bottom Face Center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
     
    CenterPos.Set 0, 0, parElecEquipmentHeight / 2 'Left FaceCenter
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
     
    CenterPos.Set parElecEquipmentWidth, 0, parElecEquipmentHeight / 2 'Right Face center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
     'Create line by two pints:
                  
    Dim Points(0 To 23) As Double
    
    Points(0) = parElecEquipmentWidth
    Points(1) = -parElecEquipmentLength / 2
    Points(2) = 0
    
    Points(3) = 0
    Points(4) = -parElecEquipmentLength / 2
    Points(5) = 0
    
    Points(6) = 0
    Points(7) = -parElecEquipmentLength / 2
    Points(8) = parElecEquipmentHeight
    
    Points(9) = parElecEquipmentWidth
    Points(10) = -parElecEquipmentLength / 2
    Points(11) = parElecEquipmentHeight
    
    Points(12) = parElecEquipmentWidth
    Points(13) = parElecEquipmentLength / 2
    Points(14) = 0
    
    Points(15) = 0
    Points(16) = parElecEquipmentLength / 2
    Points(17) = 0
    
    Points(18) = 0
    Points(19) = parElecEquipmentLength / 2
    Points(20) = parElecEquipmentHeight
    
    Points(21) = parElecEquipmentWidth
    Points(22) = parElecEquipmentLength / 2
    Points(23) = parElecEquipmentHeight
       
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
    Set oDir = Nothing
    Set geomFactory = Nothing
  Exit Sub
    
  

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub
Private Function CreateCableNozzle(oBasePt As AutoMath.DPosition, oAxis As AutoMath.DVector, ByVal objOutputColl As Object, ByRef oPartFclt As PartFacelets.IJDPart, dNozzleIndex As Long) As IJCableNozzle
        Const METHOD = "CreateCableNozzle:"
 
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim iDistribPort As IJDistribPort
    Dim NozzlePHFactory As NozzlePHFactory
    Dim ObjCablenozzle As CableNozzle
    
    Set NozzlePHFactory = New NozzlePHFactory
    
    
    Set ObjCablenozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, dNozzleIndex, _
                                                             objOutputColl.ResourceManager)
                                
    Set iLogicalDistPort = ObjCablenozzle
    Set iDistribPort = ObjCablenozzle
    
    iLogicalDistPort.SetCenterLocation oBasePt
    
    iDistribPort.SetDirectionVector oAxis
  
    Set CreateCableNozzle = ObjCablenozzle
    
    Set NozzlePHFactory = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set ObjCablenozzle = Nothing

    Exit Function
    
ErrorHandler:
   Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext

End Function


