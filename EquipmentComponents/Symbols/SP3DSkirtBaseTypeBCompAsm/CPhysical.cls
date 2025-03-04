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
'   Creation Date:  Friday, May 4 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Five Outputs: In Which Circular Foundationport is  created at the origin
'
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

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parSkirtInsideDiameter As Double
    Dim parA As Double
    Dim parB As Double
    Dim parC As Double
    Dim parD As Double
    Dim parE As Double
    Dim parF As Double
    Dim parH As Double
    Dim parT1 As Double
    Dim parT2 As Double
    Dim parSkirtHeight As Double
    Dim parSkirtThickness As Double
    Dim parBoltSize As Double
    Dim parNoOfBolts As Double
    Dim dVesselRadius As Double
        
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parSkirtInsideDiameter = arrayOfInputs(2)
    parA = arrayOfInputs(3)
    parB = arrayOfInputs(4)
    parC = arrayOfInputs(5)
    parD = arrayOfInputs(6)
    parE = arrayOfInputs(7)
    parF = arrayOfInputs(8)
    parH = arrayOfInputs(9)
    parT1 = arrayOfInputs(10)
    parT2 = arrayOfInputs(11)
    parSkirtHeight = arrayOfInputs(12)
    parSkirtThickness = arrayOfInputs(13)
    parBoltSize = arrayOfInputs(14)
    parNoOfBolts = arrayOfInputs(15)
    
    'Origin is at DP1/PP1
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_outputColl
    dVesselRadius = 0.5 * parSkirtInsideDiameter

    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Create the Default Surface at the Origin (Output 1)
    'Create non-persistent circle to use for creating default surface ---
    Dim objPlane As IngrGeom3D.Plane3d
    Dim objCircle As IngrGeom3D.Circle3d

    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                             0, 0, 0, _
                               0, 0, 1, _
                               (dVesselRadius + parSkirtThickness + parC + parA))

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
    Set oLineString = New LineString3d
    
    dLineStrPoints(0) = dVesselRadius + parSkirtThickness + parC + parA - parB
    dLineStrPoints(1) = 0
    dLineStrPoints(2) = 0
    
    dLineStrPoints(3) = dVesselRadius + parSkirtThickness + parC + parA - parB
    dLineStrPoints(4) = 0
    dLineStrPoints(5) = parT1
    
    dLineStrPoints(6) = dVesselRadius + parSkirtThickness + parC + parA
    dLineStrPoints(7) = 0
    dLineStrPoints(8) = parT1
    
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
    
    'Insert your code for VesselSkirt(Output 3)
    Dim oVesselSkirt As Object
    Set oLineString = New LineString3d
    
    dLineStrPoints(0) = dVesselRadius
    dLineStrPoints(1) = 0
    dLineStrPoints(2) = parT1
    
    dLineStrPoints(3) = dVesselRadius
    dLineStrPoints(4) = 0
    dLineStrPoints(5) = parSkirtHeight
    
    dLineStrPoints(6) = dVesselRadius + parSkirtThickness
    dLineStrPoints(7) = 0
    dLineStrPoints(8) = parSkirtHeight
    
    dLineStrPoints(9) = dVesselRadius + parSkirtThickness
    dLineStrPoints(10) = 0
    dLineStrPoints(11) = parT1
    
    dLineStrPoints(12) = dVesselRadius
    dLineStrPoints(13) = 0
    dLineStrPoints(14) = parT1
    
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
    oAxisVec.Set 0, 0, 1
    oCenter.Set 0, 0, 0
    Set oVesselSkirt = PlaceRevolution(m_outputColl, oLineString, oAxisVec, oCenter, 2 * PI, False)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput "VesselSkirt", oVesselSkirt
    Set oVesselSkirt = Nothing
    Set oLineString = Nothing
    Set oAxisVec = Nothing
    Set oCenter = Nothing
    
    'Create the Supports
    Dim iCount As Integer
    Dim oTransMat As AutoMath.DT4x4
    Dim oRotvec As AutoMath.DVector
    Dim dRotAngle As Double
    
    'Create the Vertcal Supports
    If CmpDblGreaterthan((parC + parA), LINEAR_TOLERANCE) Then
       If CmpDblGreaterthan(parE, LINEAR_TOLERANCE) And _
                                CmpDblGreaterthan(parH, LINEAR_TOLERANCE) Then
            Dim objVerSupport As Object
            Set oLineString = New LineString3d
            
            dLineStrPoints(0) = dVesselRadius + parSkirtThickness
            dLineStrPoints(1) = -parE / 2
            dLineStrPoints(2) = parT1
            
            dLineStrPoints(3) = dVesselRadius + parSkirtThickness + parC + parA
            dLineStrPoints(4) = -parE / 2
            dLineStrPoints(5) = parT1
            
            dLineStrPoints(6) = dVesselRadius + parSkirtThickness + parC + parA
            dLineStrPoints(7) = -parE / 2
            dLineStrPoints(8) = parH - parT2
            
            dLineStrPoints(9) = dVesselRadius + parSkirtThickness
            dLineStrPoints(10) = -parE / 2
            dLineStrPoints(11) = parH - parT2
            
            dLineStrPoints(12) = dVesselRadius + parSkirtThickness
            dLineStrPoints(13) = -parE / 2
            dLineStrPoints(14) = parT1
            
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, _
                                            dLineStrPoints)
            oAxisVec.Set 0, 1, 0
            Dim dAngle As Double
            Dim dSupptoAxisAng As Double
            
            dSupptoAxisAng = (parD + parE / 2) / (dVesselRadius + parSkirtThickness + parC + parA)
    
            Set oTransMat = New DT4x4
            Set oRotvec = New DVector
            oRotvec.Set 0, 0, 1
            Dim oTransVec As AutoMath.DVector
            Set oTransVec = New DVector
            
            iOutput = iOutput + 1
            For iCount = 1 To parNoOfBolts
                Set objVerSupport = PlaceProjection(m_outputColl, oLineString, oAxisVec, _
                                                              parE, True)
                oTransVec.Set 0, (parD + parE / 2), 0
                
                oTransMat.LoadIdentity
                oTransMat.Translate oTransVec
                objVerSupport.Transform oTransMat
                
                oTransMat.LoadIdentity
                oTransMat.Rotate dRotAngle, oRotvec
                objVerSupport.Transform oTransMat
                
                'Set the Output
                m_outputColl.AddOutput "Support_", objVerSupport
                Set objVerSupport = Nothing
                
                Set objVerSupport = PlaceProjection(m_outputColl, oLineString, oAxisVec, _
                                                              parE, True)
                oTransVec.Set 0, -(parD + parE / 2), 0
                
                oTransMat.LoadIdentity
                oTransMat.Translate oTransVec
                objVerSupport.Transform oTransMat
                
                oTransMat.LoadIdentity
                oTransMat.Rotate dRotAngle, oRotvec
                objVerSupport.Transform oTransMat
                
                'Set the Output
                m_outputColl.AddOutput "Support_", objVerSupport
                Set objVerSupport = Nothing
                
                dRotAngle = dRotAngle + (2 * PI / parNoOfBolts)
            Next iCount
            Set oTransMat = Nothing
            Set oTransVec = Nothing
            Set oRotvec = Nothing
        End If
            
        'Create Horizontal Supports
        If CmpDblGreaterthan(parF, LINEAR_TOLERANCE) And CmpDblGreaterthan(parT2, LINEAR_TOLERANCE) Then
            Dim objHorSupport As Object
            Set oLineString = New LineString3d
            
            dLineStrPoints(0) = dVesselRadius + parSkirtThickness
            dLineStrPoints(1) = -parF / 2
            dLineStrPoints(2) = parH - parT2
            
            dLineStrPoints(3) = dVesselRadius + parSkirtThickness
            dLineStrPoints(4) = -parF / 2
            dLineStrPoints(5) = parH
            
            dLineStrPoints(6) = dVesselRadius + parSkirtThickness
            dLineStrPoints(7) = parF / 2
            dLineStrPoints(8) = parH
            
            dLineStrPoints(9) = dVesselRadius + parSkirtThickness
            dLineStrPoints(10) = parF / 2
            dLineStrPoints(11) = parH - parT2
            
            dLineStrPoints(12) = dVesselRadius + parSkirtThickness
            dLineStrPoints(13) = -parF / 2
            dLineStrPoints(14) = parH - parT2
            
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, _
                                            dLineStrPoints)
            oAxisVec.Set 1, 0, 0
            dRotAngle = 0
            Set oTransMat = New DT4x4
            Set oRotvec = New DVector
            oRotvec.Set 0, 0, 1
            For iCount = 1 To parNoOfBolts
                Set objHorSupport = PlaceProjection(m_outputColl, oLineString, oAxisVec, _
                                                    parC + parA, True)
                oTransMat.LoadIdentity
                oTransMat.Rotate dRotAngle, oRotvec
                objHorSupport.Transform oTransMat
                'Set the output
                m_outputColl.AddOutput "Support_", objHorSupport
                Set objHorSupport = Nothing
                dRotAngle = dRotAngle + (2 * PI / parNoOfBolts)
            Next iCount
            Set oRotvec = Nothing
            Set oTransMat = Nothing
        End If
    End If
    Set oTransMat = Nothing
    Set oRotvec = Nothing
    
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
    Set oGeomFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub




