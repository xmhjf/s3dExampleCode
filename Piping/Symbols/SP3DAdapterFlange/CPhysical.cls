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
'   Copyright (c) 2006-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   ProgID:         SP3DAdapterFlange.CAdapterFlange
'   Author:         kkk
'   Creation Date:  Friday, Feb 03 2006
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  1.NOV.2007      RRK  CR-123952 Updated the symbol to support part data basis values of 15 and 20
'   01.Aug.2008     MP     CR-146407  Enhance the Adapter symbol to use Seat-to-Seat dimension per JIS G 5527
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Sub Class_Initialize()
'''
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt        As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim flangeThick1     As Double
    Dim cptOffset1       As Double
    Dim flangeDiam1      As Double
    Dim depth1           As Double
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim cptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    Dim dblAdapterDia    As Double
    Dim CenterPos        As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput          As Double
    Dim ObjInsulatedBody As Object
    Dim parFacetoFace    As Double
    Dim parSeattoSeat As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    iOutput = 0

' Place Nozzle 1
    Dim oDir              As AutoMath.DVector
    Dim objNozzle         As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint       As New AutoMath.DPosition
    Dim objAdapter        As IJDObject
    Dim stPoint           As New AutoMath.DPosition
    Dim endPoint          As New AutoMath.DPosition
    Dim dblGreaterPipeDia As Double
    
    Set oDir = New AutoMath.DVector
    oDir.Set -1, 0, 0
    
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Dim blnLengthThrHubInclFaceProj As Boolean
    
    'Assigning the boolean parameter(which will be passed to RetrieveParameters function)
    'an appropriate value based on part data basis
    If lPartDataBasis <= 1 Or lPartDataBasis = 20 Then 'Length-through-hub dimension includes flange face projection
        blnLengthThrHubInclFaceProj = True
        parFacetoFace = arrayOfInputs(2)
        parInsulationThickness = arrayOfInputs(3)
    ElseIf lPartDataBasis = 15 Then 'Length-through-hub dimension excludes flange face projection
        blnLengthThrHubInclFaceProj = False
        parFacetoFace = arrayOfInputs(2)
        parInsulationThickness = arrayOfInputs(3)
    ElseIf lPartDataBasis = 25 Then
        blnLengthThrHubInclFaceProj = True
        parSeattoSeat = arrayOfInputs(4)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, cptOffset1, depth1, blnLengthThrHubInclFaceProj
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, cptOffset2, depth2, blnLengthThrHubInclFaceProj
        parFacetoFace = parSeattoSeat + depth1 - cptOffset1 + depth2 - cptOffset2
    End If
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, cptOffset1, depth1, blnLengthThrHubInclFaceProj
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, cptOffset2, depth2, blnLengthThrHubInclFaceProj
    
    If CmpDblGreaterthan(pipeDiam1, pipeDiam2) Then
        dblGreaterPipeDia = pipeDiam1
    Else
        dblGreaterPipeDia = pipeDiam2
    End If
    
    If CmpDblGreaterthan(flangeDiam1, flangeDiam2) Then
        dblAdapterDia = flangeDiam1
    Else
      dblAdapterDia = flangeDiam2
    End If
    
    If CmpDblEqual(flangeDiam1, 0) Or CmpDblEqual(flangeDiam2, 0) Then
        dblAdapterDia = dblGreaterPipeDia
    End If
    
    
    stPoint.Set -parFacetoFace / 2, 0, 0
    endPoint.Set parFacetoFace / 2, 0, 0
   
    Set objAdapter = PlaceCylinder(m_OutputColl, stPoint, endPoint, dblAdapterDia, True)
        
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objAdapter
    Set objAdapter = Nothing
        
    oPlacePoint.Set -parFacetoFace / 2 + (depth1 - cptOffset1), 0, 0

    Set objNozzle = CreateNozzleJustaCircle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parFacetoFace / 2 - (depth2 - cptOffset2), 0, 0
    
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzleJustaCircle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
