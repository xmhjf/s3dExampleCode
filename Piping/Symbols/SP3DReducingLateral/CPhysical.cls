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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          KKC
'   Creation Date:  Tuesday, January 23 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   45 deg Reducing lateral, NO.W30 from victaulic general catalog, pg 84
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   23.JAN.2007     KKC       Created :CR-112137 Provide additional symbols required for Victaulic General Catalog
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI                  As Double
Private Sub Class_Initialize()

    PI = 4 * Atn(1)

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
    Dim oCenterPos       As AutoMath.DPosition
    
    Set oCenterPos = New AutoMath.DPosition
    oCenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim objTruncCone      As Object
    
    Dim oConeCenterBasePt    As AutoMath.DPosition
    Dim oConeCenterTopPt     As AutoMath.DPosition
    Dim dConeBaseDia         As Double
    Dim dConeTopDia          As Double
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0
    Set oConeCenterBasePt = New AutoMath.DPosition
    Set oConeCenterTopPt = New AutoMath.DPosition
    
' Place truncated cone
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oConeCenterBasePt.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    oConeCenterTopPt.Set oCenterPos.x + parFace2toCenter - flangeThick, _
                        oCenterPos.y, oCenterPos.z
    dConeTopDia = pipeDiam
    dConeBaseDia = pipeDiam
    Set objTruncCone = PlaceCone(m_OutputColl, oConeCenterBasePt, oConeCenterTopPt, dConeBaseDia / 2, dConeTopDia / 2)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTruncCone

' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set oCenterPos.x - parFace1toCenter - sptOffset + depth, _
                    oCenterPos.y, oCenterPos.z
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength( _
                        1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, parFace1toCenter)
    
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set oCenterPos.x + parFace2toCenter + sptOffset - depth, _
                    oCenterPos.y, oCenterPos.z
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    parFace3toCenter = parFace2toCenter
    oPlacePoint.Set oCenterPos.x + (parFace3toCenter + sptOffset - depth) * Cos(PI / 4), _
                    oCenterPos.y + (parFace3toCenter + sptOffset - depth) * Sin(PI / 4), _
                    oCenterPos.z
    oDir.Set Cos(PI / 4), Sin(PI / 4), 0

    Set objNozzle = CreateNozzleWithLength( _
                        3, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, parFace3toCenter)
                        
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set objTruncCone = Nothing
    Set oConeCenterBasePt = Nothing
    Set oConeCenterTopPt = Nothing
    Set oPlacePoint = Nothing
    Set oCenterPos = Nothing
    Set oDir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

