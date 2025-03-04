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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Thursday, Jan 24 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   07.MAY.2008     MA     CR-141557    Implemented part data basis for the value 13.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjLeftCone  As Object
    Dim ObjRightCone As Object
    
    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
'    parFace1toCenter = arrayOfInputs(4)
'    parFace2toCenter = arrayOfInputs(5)
    
    iOutput = 0
 
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                sptOffset2, depth2
 
    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
 
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
    ' This symbol is based on the following Part data basis values that govern its geometry
    '   (Part data Basis value -10): FacetoCenter (Default case)
    '   (Part data Basis value -13): Face1toCenter and Face2toCenter
 
    Select Case lPartDataBasis
    
    Case Is <= 1, 10
        parFacetoCenter = arrayOfInputs(2)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
    
    Case 13
        parFace1toCenter = arrayOfInputs(4)
        parFace2toCenter = arrayOfInputs(5)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
 
    Case Else
        GoTo ErrorLabel:
        
    End Select
 
    'Insert your code for output 1(Left Side Cone)
    Dim oBaseCenPoint As AutoMath.DPosition
    Dim oTopCenPoint As AutoMath.DPosition
    Set oBaseCenPoint = New AutoMath.DPosition
    Set oTopCenPoint = New AutoMath.DPosition
    oBaseCenPoint.Set -dFace1toCenter + flangeThick, 0, 0
    oTopCenPoint.Set 0, 0, 0
    Set ObjLeftCone = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, pipeDiam / 2, 0, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftCone
    Set ObjLeftCone = Nothing
    Set oBaseCenPoint = Nothing
    Set oTopCenPoint = Nothing
    
    'Insert your code for output 2(Right Side Cone)
    Dim oRightConeBasePt As AutoMath.DPosition
    Dim oRightConeTopPt  As AutoMath.DPosition
    Set oRightConeBasePt = New AutoMath.DPosition
    Set oRightConeTopPt = New AutoMath.DPosition
    oRightConeBasePt.Set 0, dFace2toCenter - flangeThick2, 0
    oRightConeTopPt.Set 0, 0, 0
    Set ObjRightCone = PlaceCone(m_OutputColl, oRightConeBasePt, oRightConeTopPt, pipeDiam2 / 2, 0, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCone
    Set ObjRightCone = Nothing
    Set oRightConeBasePt = Nothing
    Set oRightConeTopPt = Nothing
    
    'Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -dFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    'Place Nozzle 2
    oPlacePoint.Set 0, dFace2toCenter + sptOffset2 - depth2, 0
    oDir.Set 0, 1, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
 
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
