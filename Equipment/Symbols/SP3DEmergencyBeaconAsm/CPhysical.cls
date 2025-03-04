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
'   File:           CPhysical.cls
'   Author:         MA
'   Creation Date:  Monday, April 30 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Nine Outputs. Inwhich Default Surface is created at the origin
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

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parCylinderDia As Double
    Dim parBeaconLength As Double
    Dim parBeaconHeight As Double
       
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCylinderDia = arrayOfInputs(2)        'P1
    parBeaconLength = arrayOfInputs(3)       'P2
    parBeaconHeight = arrayOfInputs(4)       'P3
    
    'Origin is at DP1/PP1
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl

    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    If CmpDblGreaterthan(parCylinderDia, LINEAR_TOLERANCE) Then
        'Create non-persistent circle to use for creating default surface --- (Output 1)
        Dim objPlane As IngrGeom3D.Plane3d
        Dim objCircle As IngrGeom3D.Circle3d
            
        Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                      0, 0, 0, _
                                      1, 0, 0, _
                                    1.2 * parCylinderDia / 2)
        
        'Create persistent default surface plane - the plane can mate ---
        Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                                               (m_OutputColl.ResourceManager, objCircle)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
        Set objPlane = Nothing
        Set objCircle = Nothing
    End If

    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    If CmpDblGreaterthan(parCylinderDia, LINEAR_TOLERANCE) Then
        If CmpDblGreaterthan(parBeaconLength, LINEAR_TOLERANCE) Then
            'Create Cylinder (Output 2)
            oStPoint.Set 0, 0, 0
            oEnPoint.Set 0.1 * (parBeaconLength - parCylinderDia), 0, 0
            'Set the Output
            iOutput = iOutput + 1
            m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 1.2 * parCylinderDia
    
            'Create Cylinder (Output 3)
            oStPoint.Set 0.1 * (parBeaconLength - parCylinderDia), 0, 0
            oEnPoint.Set 0.4 * (parBeaconLength - parCylinderDia), 0, 0
            'Set the Output
            iOutput = iOutput + 1
            m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.8 * parCylinderDia
            
            'Create Box (Output 4)
            Dim oPos1   As New AutoMath.DPosition
            Dim oPos2   As New AutoMath.DPosition
            Dim oBox    As Object
            oPos1.Set 0.4 * (parBeaconLength - parCylinderDia), -0.25 * parCylinderDia, -0.05 * parBeaconHeight
            oPos2.Set (parBeaconLength - parCylinderDia), 0.25 * parCylinderDia, 0.05 * parBeaconHeight
            Set oBox = PlaceBox(m_OutputColl, oPos1, oPos2)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
            Set oPos1 = Nothing
            Set oPos2 = Nothing
            Set oBox = Nothing
        End If
        
        If CmpDblGreaterthan(parBeaconHeight, LINEAR_TOLERANCE) Then
            'Create Cylinder (Output 5)
            oStPoint.Set (parBeaconLength - 0.5 * parCylinderDia), 0, 0.2 * parBeaconHeight
            oEnPoint.Set (parBeaconLength - 0.5 * parCylinderDia), 0, -0.4 * parBeaconHeight
            'Set the Output
            iOutput = iOutput + 1
            m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parCylinderDia
            
            'Create Cylinder (Output 6)
            oStPoint.Set (parBeaconLength - 0.5 * parCylinderDia), 0, -0.4 * parBeaconHeight
            oEnPoint.Set (parBeaconLength - 0.5 * parCylinderDia), 0, -0.75 * parBeaconHeight
            'Set the Output
            iOutput = iOutput + 1
            m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.8 * parCylinderDia
        End If
    End If
    
    'Create SemiEllipsoid (Output 7)
    Dim oQuarterEllipse As Object
    Dim oSemiEllipsoid As New IngrGeom3D.Revolution3d
    Dim oEllCen As New AutoMath.DPosition
    Dim oEllNormal As New AutoMath.DVector
    Dim oEllMajor As New AutoMath.DPosition
    
    oEllCen.Set (parBeaconLength - 0.5 * parCylinderDia), 0, -0.75 * parBeaconHeight
    oEllNormal.Set 0, 0, -1
    oEllMajor.Set (parBeaconLength - 0.5 * parCylinderDia), 0.4 * parCylinderDia, -0.75 * parBeaconHeight
    
    Set oQuarterEllipse = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, oEllCen.x, oEllCen.y, oEllCen.z, _
                                                            oEllMajor.x, oEllMajor.y, oEllMajor.z, _
                                                            (parBeaconLength - 0.5 * parCylinderDia), 0, -0.8 * parBeaconHeight)
    Set oSemiEllipsoid = oGeomFactory.Revolutions3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                oQuarterEllipse, 0, 0, -1, oEllCen.x, oEllCen.y, oEllCen.z, 2 * PI, True)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oSemiEllipsoid
    Set oSemiEllipsoid = Nothing
    Set oQuarterEllipse = Nothing
    Set oEllCen = Nothing
    Set oEllNormal = Nothing
    Set oEllMajor = Nothing
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oGeomFactory = Nothing
    
    'Create the Conduit (Output 8)
    Dim oPortPos As New AutoMath.DPosition
    Dim objConduitPort As IJConduitPortOcc
    Dim oPortAxis As New AutoMath.DVector

    oPortPos.Set 0, 0, 0
    oPortAxis.Set 1, 0, 0
    Set objConduitPort = CreateConduitNozzlePH(oPortPos, oPortAxis, m_OutputColl, oPartFclt, 1)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objConduitPort
    Set objConduitPort = Nothing
    
    'Create Cable Port (Output 9)
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim iDistribPort As IJDistribPort
    Dim oNozzlePHFactory As NozzlePHFactory
    Set oNozzlePHFactory = New NozzlePHFactory
    Dim objCablePort As IJCableNozzle
    
    Set objCablePort = oNozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 2, _
                                                            m_OutputColl.ResourceManager)
    oPortPos.Set 0, 0, 0
    oPortAxis.Set 1, 0, 0
    Set iLogicalDistPort = objCablePort
    iLogicalDistPort.SetCenterLocation oPortPos
    Set iDistribPort = objCablePort
    iDistribPort.SetDirectionVector oPortAxis

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCablePort
    Set objCablePort = Nothing
    
    Set oNozzlePHFactory = Nothing
    Set iLogicalDistPort = Nothing
    Set iDistribPort = Nothing
    Set oPortPos = Nothing
    Set oPortAxis = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub


