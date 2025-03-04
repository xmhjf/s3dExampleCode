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
'   This class module has Six Outputs:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI As Double

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

    Dim oPartFclt  As PartFacelets.IJDPart
    Dim iOutput As Double
    Dim iCount As Integer

    Dim parFixtureHeight As Double
    Dim parFixtureWidth As Double
    Dim parBoxHeight As Double
    Dim parBoxWidth As Double
    Dim oTopBox As Object
    Dim dElecEquipHt As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFixtureHeight = arrayOfInputs(2)     'P1
    parFixtureWidth = arrayOfInputs(3)      'P2
    parBoxHeight = arrayOfInputs(4)         'P3
    parBoxWidth = arrayOfInputs(5)          'P4
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    dElecEquipHt = parFixtureHeight - parBoxHeight

    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
   
    If CmpDblGreaterthan(parBoxWidth, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBoxHeight, LINEAR_TOLERANCE) Then
        Dim dAngle As Double
        Dim dLineStrPoints(0 To 20)  As Double
        Dim oAxisVec As New AutoMath.DVector
        Dim oLineString As IngrGeom3D.LineString3d

        'Create TopBox (Output 1)
        dAngle = PI / 3
        oStPoint.Set 0, 0, 0
        'First point is considered on positive Z-axis at an angle (PI / 6 + dAngle) from positive Y-axis.
        For iCount = 1 To 7
            dLineStrPoints(3 * iCount - 3) = oStPoint.x - (parBoxWidth / 2) * Sin(PI / 6 + iCount * dAngle)
            dLineStrPoints(3 * iCount - 2) = oStPoint.y - (parBoxWidth / 2) * Cos(PI / 6 + iCount * dAngle)
            dLineStrPoints(3 * iCount - 1) = oStPoint.z
        Next iCount
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPoints)
        Set oAxisVec = New AutoMath.DVector
        oAxisVec.Set 0, 0, -1
        Set oTopBox = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parBoxHeight, True)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTopBox
        Set oTopBox = Nothing
        Set oAxisVec = Nothing
        Set oLineString = Nothing
        
        'Create the CylindricalJoint (Output 2)
        oStPoint.Set 0, 0, -parBoxHeight
        oEnPoint.Set 0, 0, -(parBoxHeight + 0.05 * dElecEquipHt)
        'Set the Output
        iOutput = iOutput + 1
        m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.3 * parFixtureWidth
    End If
    
    'Create the Reflector (Output 3)
    If CmpDblGreaterthan(parFixtureHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parFixtureWidth, LINEAR_TOLERANCE) Then
        Dim oPoint1 As New AutoMath.DPosition
        Dim oPoint2 As New AutoMath.DPosition
        Dim oPoint3 As New AutoMath.DPosition
        Dim objArc As Object
        Dim objBayLight As Object

        oPoint1.Set 0.15 * parFixtureWidth, 0, -(parBoxHeight + 0.05 * dElecEquipHt)
        oPoint2.Set 0.45 * parFixtureWidth, 0, -0.6 * dElecEquipHt
        oPoint3.Set 0.5 * parFixtureWidth, 0, -dElecEquipHt
        Set objArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, oPoint1.x, oPoint1.y, oPoint1.z, oPoint2.x, oPoint2.y, oPoint2.z, oPoint3.x, oPoint3.y, oPoint3.z)
            
        Dim orevVector As New AutoMath.DVector
        Dim orevCenterPt As New AutoMath.DPosition
        orevVector.Set 0, 0, 1
        orevCenterPt.Set 0, 0, -(parBoxHeight + 0.1 * dElecEquipHt)
        Set objBayLight = PlaceRevolution(m_OutputColl, objArc, orevVector, orevCenterPt, PI * 2, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBayLight
        Set objBayLight = Nothing
        Set objArc = Nothing
        Set orevVector = Nothing
        Set orevCenterPt = Nothing
        Set oPoint1 = Nothing
        Set oPoint2 = Nothing
        Set oPoint3 = Nothing
     End If
    
    'Create the Conduit Port (Output 4)
    Dim oPortPos As New AutoMath.DPosition
    Dim objConduitPort As IJConduitPortOcc
    Dim oPortAxis As New AutoMath.DVector

    oPortPos.Set 0, 0, 0
    oPortAxis.Set 0, 0, 1
    Set objConduitPort = CreateConduitNozzlePH(oPortPos, oPortAxis, m_OutputColl, oPartFclt, 1)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objConduitPort
    Set objConduitPort = Nothing
    
    'Create the Cable Port (Output 5)
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
    Set oPortPos = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    'Create the Default Surface at the top of the Bay Light (Output 6)
   If CmpDblGreaterthan(parBoxWidth, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBoxHeight, LINEAR_TOLERANCE) Then
        'Create non-persistent circle to use for creating default surface ---
        Dim objPlane As IngrGeom3D.Plane3d
        Dim objCircle As IngrGeom3D.Circle3d
        
        Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                  0, 0, 0, _
                                 0, 0, -1, _
                                  0.5 * parBoxWidth / 2)
    
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


