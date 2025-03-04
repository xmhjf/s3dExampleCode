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
'   Author:         MA
'   Creation Date:  Tuesday, August 21 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Six Outputs:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
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
    
    Dim parFixtureLength As Double
    Dim parFixtureWidth  As Double
    Dim parFixtureHeight As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFixtureLength = arrayOfInputs(2)    'L
    parFixtureWidth = arrayOfInputs(3)     'W
    parFixtureHeight = arrayOfInputs(4)    'H
    
    iOutput = 0
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    'Create Fixture Body (Output 1)
    If CmpDblGreaterthan(parFixtureLength, LINEAR_TOLERANCE) And _
            CmpDblGreaterthan(parFixtureWidth, LINEAR_TOLERANCE) And _
            CmpDblGreaterthan(parFixtureHeight, LINEAR_TOLERANCE) Then
            
        Dim objFixtureBody As Object
        oStPoint.Set -parFixtureLength / 2, -parFixtureWidth / 2, -parFixtureHeight / 2
        oEnPoint.Set parFixtureLength / 2, parFixtureWidth / 2, parFixtureHeight / 2
        
        Set objFixtureBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFixtureBody
        Set objFixtureBody = Nothing
    End If
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
   
    Dim oPortPos As AutoMath.DPosition
    Dim oPortAxis As AutoMath.DVector
    Set oPortPos = New AutoMath.DPosition
    Set oPortAxis = New AutoMath.DVector
    
    'Create the Conduit Port 1 (Output 2)
    Dim objConduitPort1 As IJConduitPortOcc
    
    oPortPos.Set -parFixtureLength / 2, 0, 0
    oPortAxis.Set -1, 0, 0
    
    Set objConduitPort1 = CreateConduitNozzlePH(oPortPos, oPortAxis, m_OutputColl, oPartFclt, 1)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objConduitPort1
    Set objConduitPort1 = Nothing
    
    'Create the Conduit Port 2 (Output 3)
    Dim objConduitPort2 As IJConduitPortOcc

    oPortPos.Set parFixtureLength / 2, 0, 0
    oPortAxis.Set 1, 0, 0
    
    Set objConduitPort2 = CreateConduitNozzlePH(oPortPos, oPortAxis, m_OutputColl, oPartFclt, 2)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objConduitPort2
    Set objConduitPort2 = Nothing
    
    'Create the Cable Port 1(Output 4)
    Dim objCableNozzle1 As CableNozzle
    Dim iDistribPort As IJDistribPort
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim oNozzlePHFactory As NozzlePHFactory
    Set oNozzlePHFactory = New NozzlePHFactory
    
    Set objCableNozzle1 = oNozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 3, _
                                                            m_OutputColl.ResourceManager)
    
    
    oPortPos.Set -parFixtureLength / 2, 0, 0
    oPortAxis.Set -1, 0, 0
    
    Set iLogicalDistPort = objCableNozzle1
    iLogicalDistPort.SetCenterLocation oPortPos
    Set iDistribPort = objCableNozzle1
    iDistribPort.SetDirectionVector oPortAxis
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableNozzle1
    Set objCableNozzle1 = Nothing
    
    'Create the Cable Port 2(Output 5)
    Dim objCableNozzle2 As CableNozzle
    
    Set objCableNozzle2 = oNozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 4, _
                                                            m_OutputColl.ResourceManager)
    
    oPortPos.Set parFixtureLength / 2, 0, 0
    oPortAxis.Set 1, 0, 0
    
    Set iLogicalDistPort = objCableNozzle2
    iLogicalDistPort.SetCenterLocation oPortPos
    Set iDistribPort = objCableNozzle2
    iDistribPort.SetDirectionVector oPortAxis

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableNozzle2
    Set objCableNozzle2 = Nothing
    
    Set oNozzlePHFactory = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set oPortPos = Nothing
    Set oPortAxis = Nothing
    
   'Create the Default Surface (Output 6)
    If CmpDblGreaterthan(parFixtureLength, LINEAR_TOLERANCE) And _
       CmpDblGreaterthan(parFixtureWidth, LINEAR_TOLERANCE) And _
       CmpDblGreaterthan(parFixtureHeight, LINEAR_TOLERANCE) Then
        
        Dim objPlane As IngrGeom3D.Plane3d
        Dim dLineStrPoints(0 To 14) As Double
        Dim oLineString As IngrGeom3D.LineString3d
        dLineStrPoints(0) = -parFixtureLength / 2
        dLineStrPoints(1) = -parFixtureWidth / 2
        dLineStrPoints(2) = parFixtureHeight / 2
        
        dLineStrPoints(3) = -parFixtureLength / 2
        dLineStrPoints(4) = parFixtureWidth / 2
        dLineStrPoints(5) = parFixtureHeight / 2
        
        dLineStrPoints(6) = parFixtureLength / 2
        dLineStrPoints(7) = parFixtureWidth / 2
        dLineStrPoints(8) = parFixtureHeight / 2
        
        dLineStrPoints(9) = parFixtureLength / 2
        dLineStrPoints(10) = -parFixtureWidth / 2
        dLineStrPoints(11) = parFixtureHeight / 2
        
        dLineStrPoints(12) = -parFixtureLength / 2
        dLineStrPoints(13) = -parFixtureWidth / 2
        dLineStrPoints(14) = parFixtureHeight / 2
        
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        
        'Create persistent default surface plane - the plane can mate ---
        
        Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                                           (m_OutputColl.ResourceManager, oLineString)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
        Set objPlane = Nothing
        Set oLineString = Nothing
    End If
    
    Set oGeomFactory = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub


