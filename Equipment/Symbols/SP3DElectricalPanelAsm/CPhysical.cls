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
'   Creation Date:  Friday, Sep 15 2006
'
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double

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
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim ObjPanelboardBody As Object
    Dim ObjDefaultSurface As Object
    Dim ObjConduitPort As IJConduitPortOcc
    Dim ObjCablePort As IJCableNozzle
    Dim objPoint As Object
    
    Dim iOutput     As Double
    Dim parPanelboardHeight As Double
    Dim parPanelboardWidth As Double
    Dim parPanelboardLength As Double
    
    Dim pRightCenterTopFace   As IJDPosition
    Dim pConduitPort   As IJDPosition
    Dim pCablePort     As IJDPosition
    Dim pBodyCorner1   As IJDPosition
    Dim pBodyCorner2   As IJDPosition
    Dim oDir           As IJDVector
  
    Set pRightCenterTopFace = New DPosition
    Set pConduitPort = New DPosition
    Set pCablePort = New DPosition
    Set pBodyCorner1 = New DPosition
    Set pBodyCorner2 = New DPosition
    Set oDir = New DVector
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPanelboardHeight = arrayOfInputs(2)
    parPanelboardWidth = arrayOfInputs(3)
    parPanelboardLength = arrayOfInputs(4)
      
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_outputColl

    oDir.Set 0, 0, 1
 
    'CenterPoint of the Backface is the origin(0,0,0) of the Panelboard.
    'Build the symbol accordingly.
    
    'Create the Panelboard Body
    pBodyCorner1.Set -parPanelboardLength, -parPanelboardWidth / 2, -parPanelboardHeight / 2
    pBodyCorner2.Set pBodyCorner1.x + parPanelboardLength, pBodyCorner1.y + parPanelboardWidth, pBodyCorner1.z + parPanelboardHeight
    Set ObjPanelboardBody = PlaceBox(m_outputColl, pBodyCorner1, pBodyCorner2)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPanelboardBody
    Set ObjPanelboardBody = Nothing
    
    'The RightCenter of the top face
    pRightCenterTopFace.Set -parPanelboardLength / 2, -parPanelboardWidth / 2, parPanelboardHeight / 2
    
    'Create the Conduit and Cable Ports
    pConduitPort.Set pRightCenterTopFace.x, pRightCenterTopFace.y + parPanelboardWidth / 5#, pRightCenterTopFace.z
    Set ObjConduitPort = CreateConduitNozzlePH(pConduitPort, oDir, m_outputColl, oPartFclt, 1)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
    
    pConduitPort.Set pRightCenterTopFace.x, pRightCenterTopFace.y + (2# * parPanelboardWidth) / 5#, pRightCenterTopFace.z
    Set ObjConduitPort = CreateConduitNozzlePH(pConduitPort, oDir, m_outputColl, oPartFclt, 2)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
 
    pConduitPort.Set pRightCenterTopFace.x, pRightCenterTopFace.y + (3# * parPanelboardWidth) / 5#, pRightCenterTopFace.z
    Set ObjConduitPort = CreateConduitNozzlePH(pConduitPort, oDir, m_outputColl, oPartFclt, 3)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
 
    pConduitPort.Set pRightCenterTopFace.x, pRightCenterTopFace.y + (4# * parPanelboardWidth) / 5#, pRightCenterTopFace.z
    Set ObjConduitPort = CreateConduitNozzlePH(pConduitPort, oDir, m_outputColl, oPartFclt, 4)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing

    pCablePort.Set pRightCenterTopFace.x, pRightCenterTopFace.y + parPanelboardWidth / 5#, pRightCenterTopFace.z
    Set ObjCablePort = CreateCableNozzle(pCablePort, oDir, m_outputColl, oPartFclt, 5)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCablePort
    Set ObjCablePort = Nothing
    
    pCablePort.Set pRightCenterTopFace.x, pRightCenterTopFace.y + (2# * parPanelboardWidth) / 5#, pRightCenterTopFace.z
    Set ObjCablePort = CreateCableNozzle(pCablePort, oDir, m_outputColl, oPartFclt, 6)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCablePort
    Set ObjCablePort = Nothing
 
    pCablePort.Set pRightCenterTopFace.x, pRightCenterTopFace.y + (3# * parPanelboardWidth) / 5#, pRightCenterTopFace.z
    Set ObjCablePort = CreateCableNozzle(pCablePort, oDir, m_outputColl, oPartFclt, 7)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCablePort
    Set ObjCablePort = Nothing
 
    pCablePort.Set pRightCenterTopFace.x, pRightCenterTopFace.y + (4# * parPanelboardWidth) / 5#, pRightCenterTopFace.z
    Set ObjCablePort = CreateCableNozzle(pCablePort, oDir, m_outputColl, oPartFclt, 8)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCablePort
    Set ObjCablePort = Nothing

   'Create salient points on each face of the body
    Dim CenterPos As IJDPosition
    Set CenterPos = New DPosition
    
    CenterPos.Set 0, 0, 0 'Back Face Center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing

    CenterPos.Set -parPanelboardLength, 0, 0 'Front Face Center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing

    CenterPos.Set -parPanelboardLength / 2, -parPanelboardWidth / 2, 0 'Right Face Center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    CenterPos.Set -parPanelboardLength / 2, parPanelboardWidth / 2, 0 'Left Face Center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
     
    CenterPos.Set -parPanelboardLength / 2, 0, -parPanelboardHeight / 2 'Bottom Face Center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
     
    CenterPos.Set -parPanelboardLength / 2, 0, parPanelboardHeight / 2 'Top Face Center
    Set objPoint = CreatePoint(m_outputColl, CenterPos, Nothing)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing

    'Create line by two pints:
            
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim Points(0 To 23) As Double
    Dim line As Object
    
    Points(0) = -parPanelboardLength
    Points(1) = -parPanelboardWidth / 2
    Points(2) = -parPanelboardHeight / 2
    
    Points(3) = -parPanelboardLength
    Points(4) = parPanelboardWidth / 2
    Points(5) = -parPanelboardHeight / 2
    
    Points(6) = -parPanelboardLength
    Points(7) = parPanelboardWidth / 2
    Points(8) = parPanelboardHeight / 2
    
    Points(9) = -parPanelboardLength
    Points(10) = -parPanelboardWidth / 2
    Points(11) = parPanelboardHeight / 2
    
    Points(12) = 0
    Points(13) = -parPanelboardWidth / 2
    Points(14) = -parPanelboardHeight / 2
    
    Points(15) = 0
    Points(16) = parPanelboardWidth / 2
    Points(17) = -parPanelboardHeight / 2
    
    Points(18) = 0
    Points(19) = parPanelboardWidth / 2
    Points(20) = parPanelboardHeight / 2
    
    Points(21) = 0
    Points(22) = -parPanelboardWidth / 2
    Points(23) = parPanelboardHeight / 2
       
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

    Dim SurfacePoints(0 To 11) As Double
     
    SurfacePoints(0) = 0
    SurfacePoints(1) = -parPanelboardWidth / 2
    SurfacePoints(2) = -parPanelboardHeight / 2
    
    SurfacePoints(3) = 0
    SurfacePoints(4) = parPanelboardWidth / 2
    SurfacePoints(5) = -parPanelboardHeight / 2
    
    SurfacePoints(6) = 0
    SurfacePoints(7) = parPanelboardWidth / 2
    SurfacePoints(8) = parPanelboardHeight / 2
    
    SurfacePoints(9) = 0
    SurfacePoints(10) = -parPanelboardWidth / 2
    SurfacePoints(11) = parPanelboardHeight / 2

    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByPoints(m_outputColl.ResourceManager, 4, SurfacePoints)
    
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
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
  ReportUnanticipatedError2 MODULE, METHOD
End Function











