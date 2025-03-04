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
'   Author:         VRK
'   Creation Date:  Friday, April 27 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Four Outputs:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper

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
    
    Dim parFixtureLength As Double
    Dim parFixtureWidth  As Double
    Dim parFixtureHeight As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFixtureLength = arrayOfInputs(2)    'P1
    parFixtureWidth = arrayOfInputs(3)     'P2
    parFixtureHeight = arrayOfInputs(4)    'P3
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
   
    If CmpDblGreaterthan(parFixtureLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parFixtureWidth, LINEAR_TOLERANCE) And CmpDblGreaterthan(parFixtureHeight, LINEAR_TOLERANCE) Then
        Dim oPoint1 As New AutoMath.DPosition
        Dim oPoint2 As New AutoMath.DPosition
        Dim oPoint3 As New AutoMath.DPosition
        Dim objarc As Object
        Dim objLine As Object
        Dim objFluoFixBody As Object
        Dim objCollection As Collection
        Set objCollection = New Collection
        
        oPoint1.Set 0, 0.5 * parFixtureWidth, 0
        oPoint2.Set 0, 0.4 * parFixtureWidth, 0.4 * parFixtureHeight
        oPoint3.Set 0, 0.2 * parFixtureWidth, 0.6 * parFixtureHeight
        
        Set objarc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, oPoint1.x, oPoint1.y, oPoint1.z, oPoint2.x, oPoint2.y, oPoint2.z, oPoint3.x, oPoint3.y, oPoint3.z)
        objCollection.Add objarc
        Set objarc = Nothing

        oStPoint.Set 0, 0.2 * parFixtureWidth, 0.6 * parFixtureHeight
        oEnPoint.Set 0, 0.2 * parFixtureWidth, parFixtureHeight
        Set objLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                        oEnPoint.x, oEnPoint.y, oEnPoint.z)
        objCollection.Add objLine
        Set objLine = Nothing
        
        oStPoint.Set 0, 0.2 * parFixtureWidth, parFixtureHeight
        oEnPoint.Set 0, -0.2 * parFixtureWidth, parFixtureHeight
        Set objLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                        oEnPoint.x, oEnPoint.y, oEnPoint.z)
        objCollection.Add objLine
        Set objLine = Nothing

        oStPoint.Set 0, -0.2 * parFixtureWidth, parFixtureHeight
        oEnPoint.Set 0, -0.2 * parFixtureWidth, 0.6 * parFixtureHeight
        Set objLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                        oEnPoint.x, oEnPoint.y, oEnPoint.z)
        objCollection.Add objLine
        Set objLine = Nothing
        
        oPoint1.Set 0, -0.2 * parFixtureWidth, 0.6 * parFixtureHeight
        oPoint2.Set 0, -0.4 * parFixtureWidth, 0.4 * parFixtureHeight
        oPoint3.Set 0, -0.5 * parFixtureWidth, 0
        
        Set objarc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, oPoint1.x, oPoint1.y, oPoint1.z, oPoint2.x, oPoint2.y, oPoint2.z, oPoint3.x, oPoint3.y, oPoint3.z)
        objCollection.Add objarc
        Set objarc = Nothing
        Set oPoint1 = Nothing
        Set oPoint2 = Nothing
        Set oPoint3 = Nothing
        
        'Prepare Complex string and revolve it about Z-axis
        Dim oComplexStr As IngrGeom3D.ComplexString3d
        Dim oDirProj As AutoMath.DVector
        Set oDirProj = New AutoMath.DVector
        oStPoint.Set 0, 0.5 * parFixtureWidth, 0
        Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
        oDirProj.Set 1, 0, 0
        Set objFluoFixBody = PlaceProjection(m_OutputColl, oComplexStr, oDirProj, parFixtureLength, False)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFluoFixBody
        Set objFluoFixBody = Nothing
        Dim iCount As Integer
        For iCount = 1 To objCollection.Count
            objCollection.Remove 1
            oComplexStr.RemoveCurve True
         Next iCount
        Set oComplexStr = Nothing
        Set oDirProj = Nothing
        Set objCollection = Nothing
    End If
   
    'Create the Conduit Port (Output 2)
    Dim oPortPos As New AutoMath.DPosition
    Dim objConduitPort As IJConduitPortOcc
    Dim oPortAxis As New AutoMath.DVector

    oPortPos.Set 0.5 * parFixtureLength, 0, parFixtureHeight
    oPortAxis.Set 0, 0, 1
    Set objConduitPort = CreateConduitNozzlePH(oPortPos, oPortAxis, m_OutputColl, oPartFclt, 1)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objConduitPort
    Set objConduitPort = Nothing

    'Create the Cable Port (Output 3)
    Dim objCableNozzle As CableNozzle
    Dim iDistribPort As IJDistribPort
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim oNozzlePHFactory As NozzlePHFactory
    Set oNozzlePHFactory = New NozzlePHFactory
    
    Set objCableNozzle = oNozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 2, _
                                                            m_OutputColl.ResourceManager)
    oPortPos.Set 0.5 * parFixtureLength, 0, parFixtureHeight
                                
    Set iLogicalDistPort = objCableNozzle
    iLogicalDistPort.SetCenterLocation oPortPos
    Set iDistribPort = objCableNozzle
    iDistribPort.SetDirectionVector oPortAxis

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableNozzle
    Set objCableNozzle = Nothing
    Set oNozzlePHFactory = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set oPortPos = Nothing
    Set oPortAxis = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
  
   'Create the Default Surface at the Origin  (Output 4)
    If CmpDblGreaterthan(parFixtureLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parFixtureWidth, LINEAR_TOLERANCE) And CmpDblGreaterthan(parFixtureHeight, LINEAR_TOLERANCE) Then
        Dim objPlane As IngrGeom3D.Plane3d
        Dim dLineStrPoints(0 To 14) As Double
        Dim oLineString As IngrGeom3D.LineString3d
        dLineStrPoints(0) = 0
        dLineStrPoints(1) = -0.2 * parFixtureWidth
        dLineStrPoints(2) = parFixtureHeight
        dLineStrPoints(3) = 0
        dLineStrPoints(4) = 0.2 * parFixtureWidth
        dLineStrPoints(5) = parFixtureHeight
        dLineStrPoints(6) = parFixtureLength
        dLineStrPoints(7) = 0.2 * parFixtureWidth
        dLineStrPoints(8) = parFixtureHeight
        dLineStrPoints(9) = parFixtureLength
        dLineStrPoints(10) = -0.2 * parFixtureWidth
        dLineStrPoints(11) = parFixtureHeight
        dLineStrPoints(12) = 0
        dLineStrPoints(13) = -0.2 * parFixtureWidth
        dLineStrPoints(14) = parFixtureHeight
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        
        'Create persistent default surface plane - the plane can mate ---
        Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                                           (m_OutputColl.ResourceManager, oLineString)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
        Set objPlane = Nothing
        Set oLineString = Nothing
        Set oGeomFactory = Nothing
    End If

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub


