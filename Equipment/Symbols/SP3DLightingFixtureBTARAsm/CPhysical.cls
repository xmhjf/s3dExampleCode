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
'   Creation Date:  Monday, April 30 2007
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
    
    Dim parBallastTankDia As Double
    Dim parBallastTankLength As Double
    Dim parGuardLength As Double
    Dim parReflectorBottomDia As Double
    Dim parReflectorLength1 As Double
    Dim parReflectorLength2 As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBallastTankDia = arrayOfInputs(2)        'P1
    parBallastTankLength = arrayOfInputs(3)     'P2
    parGuardLength = arrayOfInputs(4)           'P3
    parReflectorBottomDia = arrayOfInputs(5)    'P4
    parReflectorLength1 = arrayOfInputs(6)      'P5
    parReflectorLength2 = arrayOfInputs(7)      'P6
    
    Dim dGuardDia As Double
    Dim dGuardRadius As Double
    dGuardDia = parReflectorBottomDia / 3
    dGuardRadius = dGuardDia / 2
    
    'Origin is at DP1/PP1
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl

    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    'Create the Default Surface at the top of the Ballast Housing (Output 1)
    If CmpDblGreaterthan(parBallastTankDia, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBallastTankLength, LINEAR_TOLERANCE) Then
        'Create non-persistent circle to use for creating default surface ---
        Dim objPlane As IngrGeom3D.Plane3d
        Dim objCircle As IngrGeom3D.Circle3d
        
        Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                  0, 0, parBallastTankLength, _
                                 0, 0, -1, _
                                 0.5 * parBallastTankDia / 2)
    
        'Create persistent default surface plane - the plane can mate ---
        Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                                           (m_OutputColl.ResourceManager, objCircle)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
        Set objPlane = Nothing
        Set objCircle = Nothing
    End If
    
    If CmpDblGreaterthan(parBallastTankLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBallastTankDia, LINEAR_TOLERANCE) Then
        'Create Ballast tank Support Cylinder (Output 2)
        oStPoint.Set 0, 0, parBallastTankLength
        oEnPoint.Set 0, 0, 0.85 * oStPoint.z
        'Set the Output
        iOutput = iOutput + 1
        m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parBallastTankDia
        
        'Create the Ballast Tank (Output 3)
        oStPoint.Set 0, 0, parBallastTankLength
        oEnPoint.Set 0, 0, 0
        'Set the Output
        iOutput = iOutput + 1
        m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.9 * parBallastTankDia
    End If
    
    'Create the Reflector (Output 4)
    If CmpDblGreaterthan(parBallastTankDia, LINEAR_TOLERANCE) And CmpDblGreaterthan(parReflectorLength1, LINEAR_TOLERANCE) And _
                    CmpDblGreaterthan(parReflectorLength1, LINEAR_TOLERANCE) And CmpDblGreaterthan(parReflectorBottomDia, LINEAR_TOLERANCE) Then
        Dim oRefOrigin As New AutoMath.DPosition
        Dim oRefNormal As New AutoMath.DVector
        Dim dReflectorAngle As Double
        Dim oTopCircle As Object
        Dim oBotCircle As Object
        Dim objRuledSurface As Object
        
        dReflectorAngle = Atn((parReflectorLength2 - parReflectorLength1) / parReflectorBottomDia)
        oRefOrigin.Set 0, 0, 0
        oRefNormal.Set 0, 0, 1
        Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oRefOrigin.x, oRefOrigin.y, oRefOrigin.z, _
                                                oRefNormal.x, oRefNormal.y, oRefNormal.z, (0.9 * parBallastTankDia) / 2)
        
        oRefOrigin.Set 0, 0, -(parReflectorLength2 - (parReflectorBottomDia / 2) * Tan(dReflectorAngle))
        oRefNormal.Set Sin(dReflectorAngle), 0, Cos(dReflectorAngle)
        oRefNormal.Length = 1
        Set oBotCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oRefOrigin.x, oRefOrigin.y, oRefOrigin.z, oRefNormal.x, oRefNormal.y, oRefNormal.z, (parReflectorBottomDia / 2) / Cos(dReflectorAngle))
        
        Set objRuledSurface = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oBotCircle, oTopCircle, False)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRuledSurface
        Set oRefOrigin = Nothing
        Set oRefNormal = Nothing
        Set oTopCircle = Nothing
        Set oBotCircle = Nothing
        Set objRuledSurface = Nothing
    End If
    
    'Create the Guard (Output 5)
    If CmpDblGreaterthan(dGuardDia, LINEAR_TOLERANCE) And CmpDblGreaterthan(parGuardLength, LINEAR_TOLERANCE) Then
        Dim dGuardCylAng As Double
        Dim dGuardCylAngLoc As Double
        Dim iCount As Integer
        dGuardCylAng = 0
        dGuardCylAngLoc = dGuardRadius - (0.05 * dGuardRadius)
        Dim dGuardlength As Double
        dGuardlength = parGuardLength - (0.5 * (dGuardRadius - (0.05 * dGuardRadius)) * Sin(PI / 2))
        iOutput = iOutput + 1
        For iCount = 1 To 4
            oStPoint.Set dGuardCylAngLoc * Sin(dGuardCylAng), dGuardCylAngLoc * Cos(dGuardCylAng), 0
            oEnPoint.Set oStPoint.x, oStPoint.y, -dGuardlength
            m_oSymGeomHelper.CreateCylinder "Guard_", oStPoint, oEnPoint, 0.05 * dGuardDia
            dGuardCylAng = dGuardCylAng + PI / 2
        Next iCount
        Dim oOrigin As New AutoMath.DPosition
        Dim oMajor As New AutoMath.DVector
        Dim oNormal As New AutoMath.DVector
        oNormal.Set 0, 1, 0
        oMajor.Set 1, 0, 0
        oOrigin.Set 0.5 * (dGuardRadius - (0.05 * dGuardRadius)), 0, -dGuardlength
        m_oSymGeomHelper.CreateTorus "Guard_", oOrigin, oNormal, oMajor, 0.5 * (dGuardRadius - (0.05 * dGuardRadius)), 90, 0.05 * dGuardDia
        oOrigin.Set -0.5 * ((dGuardRadius - (0.05 * dGuardRadius))), 0, -dGuardlength
        oNormal.Set 0, -1, 0
        m_oSymGeomHelper.CreateTorus "Guard_", oOrigin, oNormal, oMajor, -0.5 * (dGuardRadius - (0.05 * dGuardRadius)), 90, 0.05 * dGuardDia
        oNormal.Set -1, 0, 0
        oOrigin.Set 0, 0.5 * (dGuardRadius - (0.05 * dGuardRadius)), -dGuardlength
        m_oSymGeomHelper.CreateTorus "Guard_", oOrigin, oNormal, oMajor, 0.5 * (dGuardRadius - (0.05 * dGuardRadius)), 90, 0.05 * dGuardDia
        oOrigin.Set 0, -0.5 * (dGuardRadius - (0.05 * dGuardRadius)), -dGuardlength
        oNormal.Set 1, 0, 0
        m_oSymGeomHelper.CreateTorus "Guard_", oOrigin, oNormal, oMajor, -0.5 * (dGuardRadius - (0.05 * dGuardRadius)), 90, 0.05 * dGuardDia
        
        oStPoint.Set 0.5 * (dGuardRadius - (0.05 * dGuardRadius)), 0, -(dGuardlength + 0.5 * (dGuardRadius - (0.05 * dGuardRadius)))
        oEnPoint.Set -0.5 * (dGuardRadius - (0.05 * dGuardRadius)), 0, -(dGuardlength + 0.5 * (dGuardRadius - (0.05 * dGuardRadius)))
        m_oSymGeomHelper.CreateCylinder "Guard_", oStPoint, oEnPoint, 0.05 * dGuardDia
        oStPoint.Set 0, 0.5 * (dGuardRadius - (0.05 * dGuardRadius)), -(dGuardlength + 0.5 * (dGuardRadius - (0.05 * dGuardRadius)))
        oEnPoint.Set 0, -0.5 * (dGuardRadius - (0.05 * dGuardRadius)), -(dGuardlength + 0.5 * (dGuardRadius - (0.05 * dGuardRadius)))
        m_oSymGeomHelper.CreateCylinder "Guard_", oStPoint, oEnPoint, 0.05 * dGuardDia

        Set oOrigin = Nothing
        Set oMajor = Nothing
        Set oNormal = Nothing
    End If
    
    'Create the ConduitPort (Output 6)
    Dim oPortPos As New AutoMath.DPosition
    Dim objConduitPort As IJConduitPortOcc
    Dim oPortAxis As New AutoMath.DVector

    oPortPos.Set 0, 0, parBallastTankLength
    oPortAxis.Set 0, 0, 1
    Set objConduitPort = CreateConduitNozzlePH(oPortPos, oPortAxis, m_OutputColl, oPartFclt, 1)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objConduitPort

    'Create the Cable Port (Output 7)
    Dim objCableNozzle As CableNozzle
    Dim iDistribPort As IJDistribPort
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim oNozzlePHFactory As NozzlePHFactory
    Set oNozzlePHFactory = New NozzlePHFactory
    
    Set objCableNozzle = oNozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 2, _
                                                            m_OutputColl.ResourceManager)
                                
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
    Set oGeomFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub


