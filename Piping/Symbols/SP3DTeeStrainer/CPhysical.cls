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
'   Author:         RRK
'   Creation Date:  Monday 29, Jan 2007
'   Description:
'   This implements the physical aspect of the symbol created for T Strainer to support multiple dimensional bases
'   as per CR-113397
'   The symbol supports the following Part data basis value
'   1.)Part data basis value 85:-'T Strainer, specified by face-to-center dimension',
'      the T strainer symbol will use the Face-to-Center dimension for port 1, port 2,
'      and port 3
'   2.)Part data basis value 90:-T Strainer, specified by Face-to-Face Dimension with Face-to-Center Dimension for Branch
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    
    Dim parFacetoCenter As Double
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
'    parFacetoFace = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    iOutput = 0

    If lPartDataBasis <= 1 Or lPartDataBasis = 85 Then
        parFacetoCenter = arrayOfInputs(2)
    ElseIf lPartDataBasis = 90 Then
        parFacetoCenter = arrayOfInputs(2)
        parFacetoFace = arrayOfInputs(3)
    Else: GoTo ErrorLabel
    End If
    

' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim dNozzleLength As Double
    
    If (lPartDataBasis = 90) Then
        oPlacePoint.Set (-parFacetoFace / 2) - sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
        dNozzleLength = parFacetoFace
        If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
        
    ElseIf (lPartDataBasis <= 1 Or lPartDataBasis = 85) Then
        oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
        dNozzleLength = parFacetoCenter
        If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    End If

    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If (lPartDataBasis = 90) Then
        oPlacePoint.Set (parFacetoFace / 2) + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        
    ElseIf (lPartDataBasis <= 1 Or lPartDataBasis = 85) Then
        oPlacePoint.Set parFacetoCenter + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
        dNozzleLength = parFacetoCenter
        If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
        Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)
    End If
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oPlacePoint.Set 0, parFacetoCenter, 0
    oDir.Set 0, 1, 0

    dNozzleLength = parFacetoCenter
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick

    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

