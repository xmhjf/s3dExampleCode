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
'   Author:         dkl
'   Creation Date:  Wednesday, August 22 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   22.Aug.2007     dkl    CR 123851, Created the symbol.
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
    Dim iOutput     As Integer
    
    Dim parHousingDia As Double
    Dim parHousingLength As Double
    Dim parGuardLength As Double
    Dim parGuardDia As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHousingDia = arrayOfInputs(2)
    parHousingLength = arrayOfInputs(3)
    parGuardLength = arrayOfInputs(4)
    parGuardDia = arrayOfInputs(5)
            
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl

    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
  
    'Create the Default Surface at the top of the Ballast Housing (Output 1)
    If CmpDblGreaterthan(parHousingDia, LINEAR_TOLERANCE) And _
                            CmpDblGreaterthan(parHousingLength, LINEAR_TOLERANCE) Then
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
    
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
        
    'Create Ballast Housing Cylinder (Output 2)
    If CmpDblGreaterthan(parHousingLength, LINEAR_TOLERANCE) _
                    And CmpDblGreaterthan(parHousingDia, LINEAR_TOLERANCE) Then
        Dim oEnPoint As AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        oStPoint.Set 0, 0, parHousingLength
        oEnPoint.Set 0, 0, 0
        'Set the Output
        iOutput = iOutput + 1
        m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, _
                                                                         parHousingDia
    End If
    Set oEnPoint = Nothing
    
    'Create the Guard (Output 4)
    If CmpDblGreaterthan(parGuardDia, LINEAR_TOLERANCE) And _
                               CmpDblGreaterthan(parGuardLength, LINEAR_TOLERANCE) Then
        Dim oEllipticalArc As IngrGeom3D.EllipticalArc3d
        Dim oGuardBottomRibs As IngrGeom3D.IJElements
        Dim oCircleForBottomRibs As IngrGeom3D.Circle3d
        Dim oGuardRibLine As IngrGeom3D.Line3d
        Dim oTraceString As IngrGeom3D.ComplexString3d
        Dim oCurves As Collection
        Dim stNorm() As Double
        Dim ednorm() As Double
        Dim dGuardCylAng As Double
        Dim dGuardCylAngLoc As Double
        Dim oOrigin As AutoMath.DPosition
        Set oOrigin = New AutoMath.DPosition
        oOrigin.Set 0, 0, -parGuardLength
        Dim oMajor As AutoMath.DVector
        Set oMajor = New AutoMath.DVector
        oMajor.Set 0, 0, -1
        dGuardCylAng = PI / 2
        dGuardCylAngLoc = parGuardDia / 2 - (0.15 * parGuardDia / 2)
        Dim iCount As Integer
       
        For iCount = 1 To 6
            'Create Trace curves for Guard.
            Set oEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                Nothing, 0, 0, -parGuardLength * 0.8, _
                Cos(dGuardCylAng), -Sin(dGuardCylAng), 0, _
                 0, 0, -parGuardLength * 0.2, _
                (parGuardDia * 0.5 - 0.075 * parGuardDia) / (parGuardLength * 0.2), 0, PI / 2)
            Set oCurves = New Collection
            oCurves.Add oEllipticalArc
            
            Set oGuardRibLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                dGuardCylAngLoc * Sin(dGuardCylAng), dGuardCylAngLoc * Cos(dGuardCylAng), -parGuardLength * 0.8, _
                dGuardCylAngLoc * Sin(dGuardCylAng), dGuardCylAngLoc * Cos(dGuardCylAng), 0)
            oCurves.Add oGuardRibLine
             
            oStPoint.Set 0, 0, -parGuardLength
            Set oTraceString = PlaceTrCString(oStPoint, oCurves)
            
            'Create Cross-Section to be traced.
            Set oCircleForBottomRibs = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                    0, 0, -parGuardLength, _
                    Sin(dGuardCylAng), Cos(dGuardCylAng), 0, _
                    0.025 * parGuardDia)
            'Form a rib of the Guard.
            Set oGuardBottomRibs = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                        oTraceString, oCircleForBottomRibs, 0, 0, stNorm, ednorm, False)
            
            Dim objSurface As Object
            For Each objSurface In oGuardBottomRibs
                iOutput = iOutput + 1
                m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSurface
            Next
            dGuardCylAng = dGuardCylAng + PI / 3
            Set oCurves = Nothing
        Next iCount
        Set objSurface = Nothing
        Set oGuardBottomRibs = Nothing
        Set oTraceString = Nothing
        Set oStPoint = Nothing
        Set oGuardRibLine = Nothing
        Set oEllipticalArc = Nothing
        Set oCircleForBottomRibs = Nothing
        Set oGeomFactory = Nothing
        
        'Create the rings on the Guard.
        Dim oNormal As AutoMath.DVector
        Set oNormal = New AutoMath.DVector
        Dim dGuardCylZ As Double
        oNormal.Set 0, 0, -1
        oMajor.Set 1, 0, 0
        For iCount = 1 To 3
            dGuardCylZ = iCount * (parGuardLength * 0.8) / 3
            oOrigin.Set 0, 0, -(dGuardCylZ - 0.05 * parGuardDia / 2)
            m_oSymGeomHelper.CreateTorus "Guard" & iCount, oOrigin, oNormal, oMajor, _
                        parGuardDia / 2 - (0.05 * parGuardDia / 2), 360, 0.05 * parGuardDia
        Next iCount
        Set oOrigin = Nothing
        Set oMajor = Nothing
        Set oNormal = Nothing
        
    End If
    
    'Create the Conduit (Output 5)
    Dim oPortPos As AutoMath.DPosition
    Set oPortPos = New AutoMath.DPosition
    Dim objConduitPort As IJConduitPortOcc
    Dim oPortAxis As AutoMath.DVector
    Set oPortAxis = New AutoMath.DVector

    oPortPos.Set 0, 0, parHousingLength
    oPortAxis.Set 0, 0, 1
    Set objConduitPort = CreateConduitNozzlePH(oPortPos, oPortAxis, m_OutputColl, oPartFclt, 1)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objConduitPort
    Set objConduitPort = Nothing
    
    'Create Cable Port (Output 6)
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
