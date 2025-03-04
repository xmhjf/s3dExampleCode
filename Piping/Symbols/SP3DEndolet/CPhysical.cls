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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          Babu Govindarajan
'   Creation Date:  Monday, July 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The Details were taken from PDS Piping Component Data - Reference Guide at Page no D-81
'   SN=F179 and MC=FOLHC. Symbol consist of Physical and Insulation aspects. Physical aspect is
'   made up of 1.Diverging Tapering Cone, 2.Tapering Cone.
'   Insulation Aspect consist of Insulation for Diverging cone, Converging cone and for Nozzle2
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   23.Jul.2003     Mallepally Srinivas         TR 44416 :Failure when constructing SP3DEndolet symbol.
'                                                                   DivergConeLen becomes negative for 4in socket welded(421) endpreperation.
'                                                                   Put a check for such condition and put ConvergConeLen to flangeThick2, so as to get positive value for DivergConeLen and the cone gets placed.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

''''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt           As PartFacelets.IJDPart
    Dim pipeDiam1           As Double
    Dim pipeDiam2           As Double
    Dim flangeThick1        As Double
    Dim sptOffset1          As Double
    Dim flangeDiam1         As Double
    Dim depth1              As Double
    Dim flangeThick2        As Double
    Dim flangeDiam2         As Double
    Dim sptOffset2          As Double
    Dim depth2              As Double
    Dim MajorBodyDiameter   As Double
    Dim BodyDiameter2       As Double
    Dim iOutput             As Double
    Dim ObjDivergCone       As Object
    Dim ObjConvergCone      As Object
    
    Dim parFacetoFace       As Double
    Dim parInsulationThickness As Double
    Dim DivergConeLen       As Double
    Dim ConvergConeLen      As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0

    
  ' Insert your code for output 1(Diverging Cone)
    Dim baseCenterPt As AutoMath.DPosition
    Dim topCenterPt As AutoMath.DPosition
    Set baseCenterPt = New AutoMath.DPosition
    Set topCenterPt = New AutoMath.DPosition
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    
  'In case port 2 is Bevel end (butt welded); i.e. flangeThick2 = 0
    If flangeThick2 = 0 Then
        ConvergConeLen = parFacetoFace * 2 / 3
        MajorBodyDiameter = pipeDiam2 * 1.65
        BodyDiameter2 = pipeDiam2
    Else
        ConvergConeLen = flangeThick2 * 1.1
        MajorBodyDiameter = flangeDiam2 * 1.25
        BodyDiameter2 = flangeDiam2
    End If
    DivergConeLen = parFacetoFace - ConvergConeLen
    
    'TR 44416 :Failure when constructing SP3DEndolet symbol.
    'DivergConeLen becomes negative for 4in socket welded(421) endpreperation.
    ' Put a check for such condition and put ConvergConeLen to flangeThick2, so as to get positive value for DivergConeLen and the cone gets placed.
    If DivergConeLen <= 0 Then
        If parFacetoFace > flangeThick2 Then
            ConvergConeLen = flangeThick2
            DivergConeLen = parFacetoFace - ConvergConeLen
        Else
            ConvergConeLen = flangeThick2 * 0.5
            DivergConeLen = parFacetoFace - ConvergConeLen
        End If
    End If
        
    baseCenterPt.Set 0, 0, 0
    topCenterPt.Set baseCenterPt.x + DivergConeLen, 0, 0
    'The enlarged portion Diameter of Endolet is 1.5 times of Pipe-2 Diameter, as given in PDS catalog
    Set ObjDivergCone = PlaceCone(m_OutputColl, baseCenterPt, topCenterPt, pipeDiam1 / 2, _
                                                MajorBodyDiameter / 2, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDivergCone
    Set ObjDivergCone = Nothing
    
 ' Insert your code for output 2(Converging Cone)
    baseCenterPt.Set topCenterPt.x, 0, 0
    topCenterPt.Set parFacetoFace, 0, 0
    'The enlarged portion Diameter of Endolet is 1.5 times of Pipe-2 Diameter, as given in PDS catalog
    Set ObjConvergCone = PlaceCone(m_OutputColl, baseCenterPt, topCenterPt, MajorBodyDiameter / 2, _
                                                BodyDiameter2 / 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConvergCone
    Set ObjConvergCone = Nothing
    Set baseCenterPt = Nothing
    Set topCenterPt = Nothing

' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set 0, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parFacetoFace + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
