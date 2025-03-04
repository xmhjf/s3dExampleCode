VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          KKC
'   Creation Date:  Monday, September 25,2006
'   Description:
'       The symbol Reference is F145 taken from PCD Page 312.
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Feb.2007    svsmylav   CR-106542: Added implementation for Part Data Basis values.
'   31.Mar.2008     BH   Added and modified implementation for parameter 'Angle'
'   12.May.2008    VRK  CR-139383 and DM-141182 Enhance the true wye symbol
'                               to support an optional angle parameter
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
End Sub
Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

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
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
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

' Insert your code for output 1(Insulated Body)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim dInsulationDiameter As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    stPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness), 0, 0
    enPoint.Set 0, 0, 0
    dInsulationDiameter = pipeDiam + parInsulationThickness * 2

'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDiameter
    
' Insert your code for output 2(Insulated Port1)
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness), 0, 0
    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDiameter
    
' Insert your code for output 3(Insulated Branch Port2)
    Dim BranchPoint   As New AutoMath.DPosition
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    BranchPoint.Set (parFace2toCenter - flangeThick - parInsulationThickness) * Cos(parAngle / 2), _
            -(parFace2toCenter - flangeThick - parInsulationThickness) * Sin(parAngle / 2), 0

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), CenterPos, BranchPoint, dInsulationDiameter
    

' Insert your code for output 4(Insulated Port2)
    stPoint.Set parFace2toCenter * Cos(parAngle / 2), -parFace2toCenter * Sin(parAngle / 2), 0
    enPoint.Set (parFace2toCenter - flangeThick - parInsulationThickness) * Cos(parAngle / 2), _
            -(parFace2toCenter - flangeThick - parInsulationThickness) * Sin(parAngle / 2), 0
    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2
    

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDiameter
    
' Insert your code for output 5(Insulated Branch Port3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    BranchPoint.Set (dFace3toCenter - flangeThick - parInsulationThickness) * Cos(parAngle / 2), _
            (dFace3toCenter - flangeThick - parInsulationThickness) * Sin(parAngle / 2), 0

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), CenterPos, BranchPoint, dInsulationDiameter
    
' Insert your code for output 6 (Insulated Port3)
    stPoint.Set dFace3toCenter * Cos(parAngle / 2), dFace3toCenter * Sin(parAngle / 2), 0
    enPoint.Set (dFace3toCenter - flangeThick - parInsulationThickness) * Cos(parAngle / 2), (dFace3toCenter - flangeThick - parInsulationThickness) * Sin(parAngle / 2), 0
    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDiameter
    Set CenterPos = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set BranchPoint = Nothing
    Set oPipeComponent = Nothing
    
Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
End Sub
