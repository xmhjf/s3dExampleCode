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
'   Creation Date:  Friday, April 19 2007
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
    
    Dim parHousingDia As Double
    Dim parHousingLength As Double
    Dim parGuardLength As Double
    Dim parGuardDia As Double
    Dim parReflectorBottomDia As Double
    Dim parReflectorLength As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHousingDia = arrayOfInputs(2)            'P1
    parHousingLength = arrayOfInputs(3)         'P2
    parGuardLength = arrayOfInputs(4)           'P3
    parGuardDia = arrayOfInputs(5)              'P4
    parReflectorBottomDia = arrayOfInputs(6)    'P5
    parReflectorLength = arrayOfInputs(7)       'P6
        
    'Origin is at DP1/PP1
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl

    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    'Create the Default Surface at the top of the Ballast Housing (Output 1)
    If CmpDblGreaterthan(parHousingDia, LINEAR_TOLERANCE) And CmpDblGreaterthan(parHousingLength, LINEAR_TOLERANCE) Then
        'Create non-persistent circle to use for creating default surface ---
        Dim objPlane As IngrGeom3D.Plane3d
        Dim objCircle As IngrGeom3D.Circle3d
        
        Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                  0, 0, parHousingLength, _
                                 0, 0, -1, _
                                0.5 * parHousingDia / 2)
    
        'Create persistent default surface plane - the plane can mate ---
        Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                                           (m_OutputColl.ResourceManager, objCircle)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
        Set objPlane = Nothing
        Set objCircle = Nothing
    End If
    
    If CmpDblGreaterthan(parHousingLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parHousingDia, LINEAR_TOLERANCE) Then
        'Create Ballast Housing Support Cylinder (Output 2)
        oStPoint.Set 0, 0, parHousingLength
        oEnPoint.Set 0, 0, 0.85 * oStPoint.z
        'Set the Output
        iOutput = iOutput + 1
        m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parHousingDia
        
        'Create the Ballast Housing (Output 3)
        oStPoint.Set 0, 0, parHousingLength
        oEnPoint.Set 0, 0, 0
        'Set the Output
        iOutput = iOutput + 1
        m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.9 * parHousingDia
    End If
    
    'Create the Reflector (Output 4)
    If CmpDblGreaterthan(parHousingDia, LINEAR_TOLERANCE) And CmpDblGreaterthan(parReflectorLength, LINEAR_TOLERANCE) And _
                    CmpDblGreaterthan(parReflectorBottomDia, LINEAR_TOLERANCE) Then
        Dim objReflector As Object
        Dim oLine As Object
        Dim oArc As Object
        Dim oCollection As New Collection
        Dim dStartAngle As Double
        Dim dLength As Double
        Dim dCenterX As Double
        Dim dCenterY As Double
        Dim dCenterz As Double
        Dim iCount As Integer
        
        oStPoint.Set 0, 0, 0
        oEnPoint.Set -(0.9 * parHousingDia) / 2, 0, 0
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                        oEnPoint.x, oEnPoint.y, oEnPoint.z)
        oCollection.Add oLine
        Set oLine = Nothing
    
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set -parReflectorBottomDia / 2, 0, -parReflectorLength
        dLength = Sqr((oEnPoint.x - oStPoint.x) * (oEnPoint.x - oStPoint.x) + (oEnPoint.y - oStPoint.y) * (oEnPoint.y - oStPoint.y) + _
                                (oEnPoint.z - oStPoint.z) * (oEnPoint.z - oStPoint.z))
        
        dStartAngle = Atn((oEnPoint.z - oStPoint.z) / (oEnPoint.x - oStPoint.x))
        dCenterX = ((oStPoint.x + oEnPoint.x) / 2) - 0.1 * dLength * Sin(dStartAngle)
        dCenterY = (oStPoint.y + oEnPoint.y) / 2
        dCenterz = ((oStPoint.z + oEnPoint.z) / 2) + 0.1 * dLength * Cos(dStartAngle)
        
        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                dCenterX, dCenterY, dCenterz, oEnPoint.x, oEnPoint.y, oEnPoint.z)
        oCollection.Add oArc
        Set oArc = Nothing
        
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set 0, 0, -parReflectorLength
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                        oEnPoint.x, oEnPoint.y, oEnPoint.z)
        oCollection.Add oLine
        Set oLine = Nothing
    
        'Prepare Complex string and revolve it about Z-axis
        Dim oRevolCen As New AutoMath.DPosition
        Dim oRevolAxis As New AutoMath.DVector
        Dim oComplexStr As IngrGeom3D.ComplexString3d
    
        oRevolCen.Set 0, 0, 0
        Set oComplexStr = PlaceTrCString(oRevolCen, oCollection)
        oRevolAxis.Set 0, 0, -1
        Set objReflector = PlaceRevolution(m_OutputColl, oComplexStr, oRevolAxis, oRevolCen, 2 * PI, False)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objReflector
        Set objReflector = Nothing
    
        'Removing the references
        Dim objTmpcurves As IJDObject
        Set objTmpcurves = oComplexStr
        objTmpcurves.Remove
        Set oComplexStr = Nothing
        Set objTmpcurves = Nothing
        Set oRevolCen = Nothing
        Set oRevolAxis = Nothing
        For iCount = 1 To oCollection.Count
            oCollection.Remove 1
        Next iCount
        Set oCollection = Nothing
    End If
    
    'Create the Guard (Output 5)
    If CmpDblGreaterthan(parGuardDia, LINEAR_TOLERANCE) And CmpDblGreaterthan(parGuardLength, LINEAR_TOLERANCE) Then
        Dim dGuardCylAng As Double
        Dim dGuardCylAngLoc As Double
        dGuardCylAng = PI / 2
        dGuardCylAngLoc = parGuardDia / 2 - (0.15 * parGuardDia / 2)
        iOutput = iOutput + 1
        For iCount = 1 To 6
            oStPoint.Set dGuardCylAngLoc * Sin(dGuardCylAng), dGuardCylAngLoc * Cos(dGuardCylAng), 0
            oEnPoint.Set oStPoint.x, oStPoint.y, -parGuardLength
            m_oSymGeomHelper.CreateCylinder "Guard_", oStPoint, oEnPoint, 0.05 * parGuardDia
            dGuardCylAng = dGuardCylAng + PI / 3
        Next iCount
        Dim oOrigin As New AutoMath.DPosition
        Dim oMajor As New AutoMath.DVector
        Dim oNormal As New AutoMath.DVector
        Dim dGuardCylZ As Double
        oNormal.Set 0, 0, -1
        oMajor.Set 1, 0, 0
        For iCount = 1 To 3
            dGuardCylZ = iCount * parGuardLength / 3
            oOrigin.Set 0, 0, -(dGuardCylZ - 0.05 * parGuardDia / 2)
            m_oSymGeomHelper.CreateTorus "Guard_", oOrigin, oNormal, oMajor, parGuardDia / 2 - (0.05 * parGuardDia / 2), 360, 0.05 * parGuardDia
        Next iCount
        Set oOrigin = Nothing
        Set oMajor = Nothing
        Set oNormal = Nothing
    End If
    
    'Create the Conduit (Output 6)
    Dim oPortPos As New AutoMath.DPosition
    Dim objConduitPort As IJConduitPortOcc
    Dim oPortAxis As New AutoMath.DVector

    oPortPos.Set 0, 0, parHousingLength
    oPortAxis.Set 0, 0, 1
    Set objConduitPort = CreateConduitNozzlePH(oPortPos, oPortAxis, m_OutputColl, oPartFclt, 1)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objConduitPort
    Set objConduitPort = Nothing
    
    'Create Cable Port (Output 7)
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim iDistribPort As IJDistribPort
    Dim oNozzlePHFactory As NozzlePHFactory
    Set oNozzlePHFactory = New NozzlePHFactory
    Dim objCablePort As IJCableNozzle
    
    Set objCablePort = oNozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 2, _
                                                            m_OutputColl.ResourceManager)
    oPortPos.Set 0, 0, parHousingLength
                                
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
