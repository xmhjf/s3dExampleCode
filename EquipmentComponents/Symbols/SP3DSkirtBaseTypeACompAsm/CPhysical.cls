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
'   Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         VRK
'   Creation Date:  Monday, May 7,2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Five Outputs:
'       In Which Circular Foundationport is  created at the origin
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   19.09.2007      RUK     TR-127065  Problems with skirt with base equipment component symbol.
'                               -Added new attribute  parNoOfBolts and according to that supports are created
'                               -Replaced the attribute parVesselDiamter With parSkirtInsideDiameter
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

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt  As PartFacelets.IJDPart
    Dim iOutput As Double

    Dim parSkirtInsideDiameter As Double
    Dim parSkirtThickness As Double
    Dim parSkirtHeight As Double
    Dim parA, parB, parC, parD, parH, parT, parE As Double
    Dim parNoOfBolts As Double
    Dim parBoltSize As Double
    Dim dSupportHt As Double
    Dim dVesselRadius As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parSkirtInsideDiameter = arrayOfInputs(2)
    parSkirtThickness = arrayOfInputs(3)
    parSkirtHeight = arrayOfInputs(4)
    parA = arrayOfInputs(5)
    parB = arrayOfInputs(6)
    parC = arrayOfInputs(7)
    parD = arrayOfInputs(8)
    parE = arrayOfInputs(9)
    parH = arrayOfInputs(10)
    parT = arrayOfInputs(11)
    parBoltSize = arrayOfInputs(12)
    parNoOfBolts = arrayOfInputs(13)
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_outputColl
    dSupportHt = parH - parT
    dVesselRadius = 0.5 * parSkirtInsideDiameter
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
      
    'Create the Default Surface at the Origin (Output 1)
    'Create non-persistent circle to use for creating default surface ---
    If CmpDblGreaterthan(dVesselRadius, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSkirtThickness, LINEAR_TOLERANCE) Then
         If CmpDblGreaterthan((parC + parA), LINEAR_TOLERANCE) Then
            Dim objPlane As IngrGeom3D.Plane3d
            Dim objCircle As IngrGeom3D.Circle3d
            Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                     0, 0, 0, _
                                    0, 0, 1, _
                                          dVesselRadius + parSkirtThickness + parC + parA)
            
            'Create persistent default surface plane - the plane can mate ---
            Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                                              (m_outputColl.ResourceManager, objCircle)
            'Set the output
            iOutput = iOutput + 1
            m_outputColl.AddOutput arrayOfOutputs(iOutput), objPlane
            Set objPlane = Nothing
            Set objCircle = Nothing
        
            'Insert your code for CircularBase(Output 2)
            Dim oBase As Object
            Dim oAxisVec As New AutoMath.DVector
            Dim oCenter As New AutoMath.DPosition
            Dim dLineStrPoints(0 To 14) As Double
            Dim oLineString As IngrGeom3D.LineString3d
            
            If CmpDblGreaterthan(parB, LINEAR_TOLERANCE) Then
                dLineStrPoints(0) = dVesselRadius + parSkirtThickness + parC + parA - parB
                dLineStrPoints(1) = 0
                dLineStrPoints(2) = 0
                dLineStrPoints(3) = dVesselRadius + parSkirtThickness + parC + parA - parB
                dLineStrPoints(4) = 0
                dLineStrPoints(5) = parT
                dLineStrPoints(6) = dVesselRadius + parSkirtThickness + parC + parA
                dLineStrPoints(7) = 0
                dLineStrPoints(8) = parT
                dLineStrPoints(9) = dVesselRadius + parSkirtThickness + parC + parA
                dLineStrPoints(10) = 0
                dLineStrPoints(11) = 0
                dLineStrPoints(12) = dVesselRadius + parSkirtThickness + parC + parA - parB
                dLineStrPoints(13) = 0
                dLineStrPoints(14) = 0
                Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
                oAxisVec.Set 0, 0, 1
                oCenter.Set 0, 0, 0
                Set oBase = PlaceRevolution(m_outputColl, oLineString, oAxisVec, oCenter, 2 * PI, False)
                'Set the output
                iOutput = iOutput + 1
                m_outputColl.AddOutput "CircularBase", oBase
                Set oBase = Nothing
                Set oLineString = Nothing
             End If
         End If
    
            'Insert your code for VesselSkirt(Output 3)
            Dim oVesselSkirt As Object
            dLineStrPoints(0) = dVesselRadius
            dLineStrPoints(1) = 0
            dLineStrPoints(2) = parT
            dLineStrPoints(3) = dVesselRadius
            dLineStrPoints(4) = 0
            dLineStrPoints(5) = parSkirtHeight
            dLineStrPoints(6) = dVesselRadius + parSkirtThickness
            dLineStrPoints(7) = 0
            dLineStrPoints(8) = parSkirtHeight
            dLineStrPoints(9) = dVesselRadius + parSkirtThickness
            dLineStrPoints(10) = 0
            dLineStrPoints(11) = parT
            dLineStrPoints(12) = dVesselRadius
            dLineStrPoints(13) = 0
            dLineStrPoints(14) = parT
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
            oAxisVec.Set 0, 0, 1
            oCenter.Set 0, 0, 0
            Set oVesselSkirt = PlaceRevolution(m_outputColl, oLineString, oAxisVec, oCenter, 2 * PI, False)
            'Set the output
            iOutput = iOutput + 1
            m_outputColl.AddOutput "Skirt", oVesselSkirt
            Set oVesselSkirt = Nothing
            Set oLineString = Nothing
            Set oCenter = Nothing
        End If
   
    'Insert your code for Supports(Output 4)
    If CmpDblGreaterthan(dVesselRadius, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSkirtThickness, LINEAR_TOLERANCE) _
         And CmpDblGreaterthan((parC + parA), LINEAR_TOLERANCE) And CmpDblGreaterthan(parD, LINEAR_TOLERANCE) And CmpDblGreaterthan(parT, LINEAR_TOLERANCE) _
                     And CmpDblGreaterthan(dSupportHt, LINEAR_TOLERANCE) Then
        'Create the Supports
        Dim dSupportLinePoints(0 To 17) As Double
        Dim objSupport  As Object
        Dim dSupptoAxisAng As Double
        
        dSupptoAxisAng = (parD + parE / 2) / (dVesselRadius + parSkirtThickness + parC + parA)
        
        dSupportLinePoints(0) = dVesselRadius + parSkirtThickness
        dSupportLinePoints(1) = -parE / 2
        dSupportLinePoints(2) = parT
        dSupportLinePoints(3) = dVesselRadius + parSkirtThickness
        dSupportLinePoints(4) = -parE / 2
        dSupportLinePoints(5) = parT + dSupportHt
        dSupportLinePoints(6) = dVesselRadius + parSkirtThickness + 0.3 * parC
        dSupportLinePoints(7) = -parE / 2
        dSupportLinePoints(8) = parT + dSupportHt
        dSupportLinePoints(9) = dVesselRadius + parSkirtThickness + parC + parA
        dSupportLinePoints(10) = -parE / 2
        dSupportLinePoints(11) = parT + 0.2 * dSupportHt
        dSupportLinePoints(12) = dVesselRadius + parSkirtThickness + parC + parA
        dSupportLinePoints(13) = -parE / 2
        dSupportLinePoints(14) = parT
        dSupportLinePoints(15) = dVesselRadius + parSkirtThickness
        dSupportLinePoints(16) = -parE / 2
        dSupportLinePoints(17) = parT
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 6, _
                                        dSupportLinePoints)
        oAxisVec.Set 0, 1, 0
        Dim iCount As Integer
        Dim dAngle As Double
        Dim dRotAngle As Double
        
        Dim oRotVec As AutoMath.DVector
        Dim oTransMat As AutoMath.DT4x4
        Dim oTransVec As AutoMath.DVector
        
        Set oTransMat = New DT4x4
        Set oRotVec = New DVector
        Set oTransVec = New DVector
        
        oRotVec.Set 0, 0, 1
        iOutput = iOutput + 1
        For iCount = 1 To parNoOfBolts
            Set objSupport = PlaceProjection(m_outputColl, oLineString, oAxisVec, _
                                                          parE, True)
            oTransVec.Set 0, (parD + parE / 2), 0
            
            oTransMat.LoadIdentity
            oTransMat.Translate oTransVec
            objSupport.Transform oTransMat
            
            oTransMat.LoadIdentity
            oTransMat.Rotate dRotAngle, oRotVec
            objSupport.Transform oTransMat
            
            'Set the Output
            m_outputColl.AddOutput "Support_", objSupport
            Set objSupport = Nothing
            
            Set objSupport = PlaceProjection(m_outputColl, oLineString, oAxisVec, _
                                                          parE, True)
            oTransVec.Set 0, -(parD + parE / 2), 0
            
            oTransMat.LoadIdentity
            oTransMat.Translate oTransVec
            objSupport.Transform oTransMat
            
            oTransMat.LoadIdentity
            oTransMat.Rotate dRotAngle, oRotVec
            objSupport.Transform oTransMat
            
            'Set the Output
            m_outputColl.AddOutput "Support_", objSupport
            Set objSupport = Nothing
            
            dRotAngle = dRotAngle + (2 * PI / parNoOfBolts)
        Next iCount
        Set oTransMat = Nothing
        Set oTransVec = Nothing
        Set oRotVec = Nothing
    End If
    Set oAxisVec = Nothing
    Set oGeomFactory = Nothing

    'Insert your code for FoundationPort (output 5)
    '==========================================
    'Construction of Equipment Foundation Port
    '==========================================
    'Place Circular FoundationPort
    Dim objFoundationPort As IJEqpFoundationPort
    Dim oNozzlePHFactory As NozzlePHFactory
    Set oNozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
    
    'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = 0
    
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0
    
    dZaxis(0) = 0
    dZaxis(1) = 0
    dZaxis(2) = -1
    
    Set objFoundationPort = oNozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_outputColl.ResourceManager)

    Call objFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
                            
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput "FoundationPort", objFoundationPort
    Set objFoundationPort = Nothing
    Set oNozzlePHFactory = Nothing
   
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub


