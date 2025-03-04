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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          SSP
'   Creation Date:  Monday, Dec 20 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMOP1,
'   PDMS Category Name=IFRB). The symbol consists of 9 oupputs: 8 physical,
'   and 1 insulation aspect outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  1.NOV.2007      RRK  CR-123952 Updated the symbol to support part data basis values of 15 and 20
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim pipeDiam3        As Double
    Dim pipeDiam4        As Double
    Dim flangeThick     As Double
    Dim flangeDiam      As Double
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim sptOffset3       As Double
    Dim sptOffset4       As Double
    Dim depth1           As Double
    Dim depth2           As Double
    Dim depth3           As Double
    Dim depth4           As Double
    
    Dim iOutput     As Double
    Dim ObjDivergingCone As Object
    Dim ObjConvergingCone As Object
    Dim ObjFlange1 As Object
    Dim ObjFlange2 As Object
    
    Dim parFacetoFace As Double
    Dim parFace1toCenter As Double
    Dim parOrificeFlangeClearance As Double
    Dim parPortRotation1 As Double
    Dim parPortRotation2 As Double
    Dim parInsulationThickness As Double
    Dim NozzleLength As Double
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)            'D2
    parFace1toCenter = arrayOfInputs(3)       'D3
    parOrificeFlangeClearance = arrayOfInputs(4)               'D4
    parPortRotation1 = arrayOfInputs(5)                     'D5
    parPortRotation2 = arrayOfInputs(6)                   'D6
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0
' The orgin is considered to be along the flow axis and at parFace1toCenter plus parClearanc/2 from nozzle 1
' Insert your code for output 1(DivergingCone)
    Dim coneBaseRad As Double
    Dim coneTopRad As Double
    Dim flgneck1 As Double
    Dim flgthk1 As Double
    Dim flgod1 As Double
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Dim blnLengthThrHubInclFaceProj As Boolean
    'Assigning the boolean parameter(which will be passed to RetrieveParameters function)
    'an appropriate value based on part data basis
    If lPartDataBasis <= 1 Or lPartDataBasis = 20 Then 'Length-through-hub dimension includes flange face projection
        blnLengthThrHubInclFaceProj = True
    ElseIf lPartDataBasis = 15 Then 'Length-through-hub dimension excludes flange face projection
        blnLengthThrHubInclFaceProj = False
    End If

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, sptOffset1, depth1, blnLengthThrHubInclFaceProj
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam, sptOffset3, depth3, blnLengthThrHubInclFaceProj
    
'    Assumptions:
'    flgthk1 = pipeDiam3 * 1.4
'    flgneck1=flgthk1*1.5
'    coneTopRad = pipeDiam1/2
'    coneBaseRad=1.6 * pipeDiam1/2 are taken
    flgthk1 = pipeDiam3 * 1.4
    flgneck1 = flgthk1 * 1.5

    stPoint.Set -parOrificeFlangeClearance / 2 - flgthk1, 0, 0
    enPoint.Set -parOrificeFlangeClearance / 2 - flgthk1 - flgneck1, 0, 0

    coneBaseRad = 1.6 * pipeDiam1 / 2
    coneTopRad = pipeDiam1 / 2
    
    Set ObjDivergingCone = PlaceCone(m_OutputColl, stPoint, enPoint, coneBaseRad, coneTopRad, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDivergingCone
    Set ObjDivergingCone = Nothing
    
 ' Insert your code for output 2(ConvergingCone)
    Dim flgneck2 As Double
    Dim flgthk2 As Double
    Dim flgod2 As Double
    'Assumptions:
'    flgthk2 = pipeDiam4* 1.4
'    flgneck1=flgthk2*1.5
'    coneTopRad = pipeDiam2/2
'    coneBaseRad=1.6 * pipeDiam2/2 are taken

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset2, depth2, blnLengthThrHubInclFaceProj
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, flangeDiam, sptOffset4, depth4, blnLengthThrHubInclFaceProj
    
    flgthk2 = pipeDiam4 * 1.4
    flgneck2 = flgthk2 * 1.5

    stPoint.Set parOrificeFlangeClearance / 2 + flgthk2, 0, 0
    enPoint.Set parOrificeFlangeClearance / 2 + flgthk2 + flgneck2, 0, 0

    coneBaseRad = 1.6 * pipeDiam2 / 2
    coneTopRad = pipeDiam2 / 2
    
    Set ObjConvergingCone = PlaceCone(m_OutputColl, stPoint, enPoint, coneBaseRad, coneTopRad, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConvergingCone
    Set ObjConvergingCone = Nothing
    
 ' Insert your code for output 3(Flange1)
 ' Assumptions:
'  flgod1 = pipeDiam1 * 2 is taken

    flgod1 = pipeDiam1 * 2
    
    stPoint.Set -parOrificeFlangeClearance / 2, 0, 0
    enPoint.Set -parOrificeFlangeClearance / 2 - flgthk1, 0, 0

    Set ObjFlange1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, flgod1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlange1
    Set ObjFlange1 = Nothing
    
 ' Insert your code for output 4(Flange2)
 ' Assumptions:
'  flgod2 = pipeDiam2 * 2 is taken

    flgod2 = pipeDiam2 * 2
    
    stPoint.Set parOrificeFlangeClearance / 2, 0, 0
    enPoint.Set parOrificeFlangeClearance / 2 + flgthk2, 0, 0
    
    Set ObjFlange2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, flgod2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlange2
    Set ObjFlange2 = Nothing
    
' Place Nozzle 1

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parOrificeFlangeClearance / 2 - parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    
    NozzleLength = parFace1toCenter - flgthk1 - flgneck1

    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                                                    oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    oPlacePoint.Set parFacetoFace - parFace1toCenter - parOrificeFlangeClearance / 2 + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    
    NozzleLength = parFacetoFace - parFace1toCenter - parOrificeFlangeClearance - flgthk2 - flgneck2

    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                                                                    oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing


'Place Nozzle 3
    Dim xCord As Double
    Dim yCord As Double
    Dim zCord As Double

    xCord = -parOrificeFlangeClearance / 2 - flgthk1 / 2
    yCord = (flgod1 / 2 + sptOffset3 - depth3) * Cos(parPortRotation1)
    zCord = (flgod1 / 2 + sptOffset3 - depth3) * Sin(parPortRotation1)

    oPlacePoint.Set xCord, yCord, zCord
    oDir.Set 0, Cos(parPortRotation1), Sin(parPortRotation1)

    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

'Place Nozzle 4
    xCord = parOrificeFlangeClearance / 2 + flgthk2 / 2
    yCord = (flgod2 / 2 + sptOffset4 - depth4) * Cos(parPortRotation2)
    zCord = (flgod2 / 2 + sptOffset4 - depth4) * Sin(parPortRotation2)

    oPlacePoint.Set xCord, yCord, zCord
    oDir.Set 0, Cos(parPortRotation2), Sin(parPortRotation2)

    Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
            
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

