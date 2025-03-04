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
'   Creation Date:  Tuesday, May 1 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Seven Outputs:
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
    
    Dim parFixtureHeight As Double
    Dim parFixtureWidth As Double
    Dim parFixtureLength As Double
    Dim parBoxHeight As Double
    Dim parBoxLength As Double
    Dim parBoxWidth As Double
    Dim dCylinderHeight As Double
    Dim oBox As Object
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFixtureHeight = arrayOfInputs(2)     'P1
    parFixtureWidth = arrayOfInputs(3)      'P2
    parFixtureLength = arrayOfInputs(4)     'P3
    parBoxHeight = arrayOfInputs(5)         'P4
    parBoxWidth = arrayOfInputs(6)          'P5
    parBoxLength = arrayOfInputs(7)         'P6
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    dCylinderHeight = 0.5 * parBoxHeight
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
   
    If CmpDblGreaterthan(parFixtureLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parFixtureWidth, LINEAR_TOLERANCE) And CmpDblGreaterthan(parFixtureHeight, LINEAR_TOLERANCE) Then
        Dim oLine As Object
        Dim objCollection As Collection
        Set objCollection = New Collection
        Dim objCurve As Object
        Dim oComplexStr As IngrGeom3D.ComplexString3d

        'Create Box (Output 1)
        oStPoint.Set -0.5 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
        oEnPoint.Set -0.5 * parFixtureLength, -0.5 * parFixtureWidth, 0
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                          oEnPoint.x, oEnPoint.y, oEnPoint.z)

        objCollection.Add oLine
        Set oLine = Nothing
        
        oStPoint.Set -0.5 * parFixtureLength, -0.5 * parFixtureWidth, 0
        oEnPoint.Set 0.5 * parFixtureLength, -0.5 * parFixtureWidth, 0
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                          oEnPoint.x, oEnPoint.y, oEnPoint.z)

        objCollection.Add oLine
        Set oLine = Nothing

        oStPoint.Set 0.5 * parFixtureLength, -0.5 * parFixtureWidth, 0
        oEnPoint.Set 0.5 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                          oEnPoint.x, oEnPoint.y, oEnPoint.z)

        objCollection.Add oLine
        Set oLine = Nothing

        
        oStPoint.Set 0.5 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
        oEnPoint.Set 0.48 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                          oEnPoint.x, oEnPoint.y, oEnPoint.z)

        objCollection.Add oLine
        Set oLine = Nothing

        oStPoint.Set 0.48 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
        oEnPoint.Set 0.48 * parFixtureLength, -0.5 * parFixtureWidth, 0.01 * parFixtureHeight
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                          oEnPoint.x, oEnPoint.y, oEnPoint.z)

        objCollection.Add oLine
        Set oLine = Nothing

        oStPoint.Set 0.48 * parFixtureLength, -0.5 * parFixtureWidth, 0.01 * parFixtureHeight
        oEnPoint.Set -0.48 * parFixtureLength, -0.5 * parFixtureWidth, 0.01 * parFixtureHeight
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                          oEnPoint.x, oEnPoint.y, oEnPoint.z)

        objCollection.Add oLine
        Set oLine = Nothing

        oStPoint.Set -0.48 * parFixtureLength, -0.5 * parFixtureWidth, 0.01 * parFixtureHeight
        oEnPoint.Set -0.48 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                          oEnPoint.x, oEnPoint.y, oEnPoint.z)

        objCollection.Add oLine
        Set oLine = Nothing

        oStPoint.Set -0.48 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
        oEnPoint.Set -0.5 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                          oEnPoint.x, oEnPoint.y, oEnPoint.z)

        objCollection.Add oLine
        Set oLine = Nothing

        oStPoint.Set -0.5 * parFixtureLength, -0.5 * parFixtureWidth, parFixtureHeight
        Dim oAxisVec As AutoMath.DVector
        Set oAxisVec = New DVector
        oAxisVec.Set 0, 1, 0
        Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
        Set objCurve = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, parFixtureWidth, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurve
        Set objCurve = Nothing
        Set oAxisVec = Nothing
        Dim iCount As Integer
        For iCount = 1 To objCollection.Count
            objCollection.Remove 1
            oComplexStr.RemoveCurve True
        Next iCount
        Set oComplexStr = Nothing
        Set objCollection = Nothing

        'Create the CylindricalJoint (Output 2and 3)
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, 0, -0.1 * dCylinderHeight
        'Set the Output
        iOutput = iOutput + 1
        m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.2 * parFixtureLength

        oStPoint.Set 0, 0, -0.1 * dCylinderHeight
        oEnPoint.Set 0, 0, -dCylinderHeight
        'Set the Output
        iOutput = iOutput + 1
        m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.1 * parFixtureLength
    End If
    
    'Create the Box(Output 4)
    If CmpDblGreaterthan(parBoxLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBoxWidth, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBoxHeight, LINEAR_TOLERANCE) Then
        oStPoint.Set -0.48 * parBoxLength, -0.5 * parBoxWidth, 0.5 * parFixtureHeight
        oEnPoint.Set 0.48 * parBoxLength, 0.5 * parBoxWidth, oStPoint.z + parBoxHeight
        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
    End If
    
    'Create the Conduit Port(Output 5)
    Dim oPortPos As New AutoMath.DPosition
    Dim objConduitPort As IJConduitPortOcc
    Dim oPortAxis As New AutoMath.DVector

    oPortPos.Set 0, 0, 0
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
    oPortPos.Set 0, 0, 0
                                
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
    Set oPortAxis = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oPortPos = Nothing
    
    'Create the Default Surface at the Origin  (Output 7)
   If CmpDblGreaterthan(parBoxWidth, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBoxHeight, LINEAR_TOLERANCE) Then
        'Create non-persistent circle to use for creating default surface ---
        Dim objPlane As IngrGeom3D.Plane3d
        Dim objCircle As IngrGeom3D.Circle3d
        
        Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                  0, 0, 0, _
                                 0, 0, -1, _
                                0.1 * parFixtureLength)
    
        'Create persistent default surface plane - the plane can mate ---
        Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                                           (m_OutputColl.ResourceManager, objCircle)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
        Set objPlane = Nothing
        Set objCircle = Nothing
        Set oGeomFactory = Nothing
    End If

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

