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
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          KKC
'   Creation Date:  Monday, September 25,2006
'   Description:
'       The symbol Reference is F145 taken from PCD Page 312.
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   09.Feb.2007    svsmylav   CR-106542: Added implementation for Part Data Basis values.
'   31.Mar.2008     BH   Added and modified implementation for parameter 'Angle'
'   12.May.2008    VRK  CR-139383 and DM-141182 Enhance the true wye symbol
'                               to support an optional angle parameter
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptoffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Integer

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double
    Dim parFace3toCenter As Double
    Dim parAngle As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
'    parFace3toCenter = arrayOfInputs(5)
    
    If UBound(arrayOfInputs) >= 6 Then
        parAngle = arrayOfInputs(6)
    Else
        parAngle = PI / 2 'where PI = 4 * Atn(1)
    End If
    
    iOutput = 0
 
  'Checking for the PartDataBasis
    Dim dFace3toCenter As Double 'Use this variable to locate port 3
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
        
    'Check if PartDataBasis is either 'Default'
    If (lPartDataBasis <= 1 Or lPartDataBasis = 150) Then
        'Or 'Wye based on two face-to-center dimensions'
        dFace3toCenter = parFace2toCenter
    ElseIf (lPartDataBasis = 155) Then
        'Wye based on three face-to-center dimensions'
        parFace3toCenter = arrayOfInputs(5)
        dFace3toCenter = parFace3toCenter
    End If

'The Body of the Y is part of Port1

' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -(parFace1toCenter - (depth - cptoffset)), 0, 0
    Dim Port1Length As Double
    Port1Length = parFace1toCenter
    oDir.Set -1, 0, 0
   
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
  
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    oPlacePoint.Set (parFace2toCenter + cptoffset - depth) * Cos(parAngle / 2), -(parFace2toCenter + cptoffset - depth) * Sin(parAngle / 2), 0
    Dim Port2Length As Double
    oDir.Set Cos(parAngle / 2), -Sin(parAngle / 2), 0
    Port2Length = parFace2toCenter
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port2Length)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
 
' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    oPlacePoint.Set (dFace3toCenter + cptoffset - depth) * Cos(parAngle / 2), _
        (dFace3toCenter + cptoffset - depth) * Sin(parAngle / 2), 0
    oDir.Set Cos(parAngle / 2), Sin(parAngle / 2), 0
    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, dFace3toCenter)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oPipeComponent = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
End Sub
