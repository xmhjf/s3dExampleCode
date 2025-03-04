VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Sunday, Jul 18 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     Eccentric Reducing Tee symbol taken from BPE BIO-TECH catalog. It consists of eight outputs
'     Three Nozzles (Physical) and others are insulation aspect outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   30.Jul.2008     KKC      CR- 146404 Enhance Eccentric Reducing Tee symbol for seat-to-seat dimension per JIS G 5527
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
    Dim pipeDiam1        As Double
    Dim flangeThick     As Double
    Dim cptoffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parSeat1toCenter As Double
    Dim parSeat2toCenter As Double
    Dim parInsulationThickness As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
    Set oPartFclt = arrayOfInputs(1)
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
' Inputs
    If lPartDataBasis <= 1 Then
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        parInsulationThickness = arrayOfInputs(4)
        
    ElseIf lPartDataBasis = 25 Then
         parSeat1toCenter = arrayOfInputs(5)
         parSeat2toCenter = arrayOfInputs(6)
         parInsulationThickness = arrayOfInputs(4)
         RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, _
                                    flangeDiam, cptoffset, depth
         dFace1toCenter = parSeat1toCenter + depth - cptoffset
         RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, _
                                    flangeDiam, cptoffset, depth
         dFace2toCenter = parSeat2toCenter + depth - cptoffset
    End If
    iOutput = 0
    
'   Origin is taken at center of port1 and port2
 
'THE Body of the Tee is part of the Port1

' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, cptoffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -dFace1toCenter - cptoffset + depth, 0, 0

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    Dim Port1Length As Double
    Port1Length = dFace1toCenter * 2 - flangeThick
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set dFace1toCenter + cptoffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth

    oPlacePoint.Set 0, dFace2toCenter + cptoffset - depth, (pipeDiam1 - pipeDiam) / 2
    oDir.Set 0, 1, 0

    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, dFace2toCenter)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
End Sub

