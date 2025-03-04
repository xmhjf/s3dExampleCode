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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RUK
'   Creation Date:  Thursday, August 08, 2008
'   Description:
'   Source: E-138 section in Design document
'   The following Part data basis cases are addressed in this symbol
'    Needle valve, linear, Type 1 (1052)
'    Needle valve, linear, Type 2 (1053)
'    Needle valve, Angle(1054)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   07.Aug.2008     RUK     CR-146837  Provide needle valve symbol
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:"    'Used for error messages
Private PI As Double
Private Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam1 As Double
    Dim flangeDiam1 As Double
    Dim flangeThick1 As Double
    Dim cptOffset1 As Double
    Dim depth1 As Double

    Dim pipeDiam2 As Double
    Dim flangeThick2 As Double
    Dim flangeDiam2 As Double
    Dim cptOffset2 As Double
    Dim depth2 As Double

    Dim iOutput As Long

    Dim parFacetoFace As Double
    Dim parFacetoCenter As Double
    Dim parOffsetFrmValCen As Double
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parHandwheelAngle As Double
    Dim parInsulationThickness As Double

    'Inputs
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFacetoFace = arrayOfInputs(2)
    'parFacetoCenter = arrayOfInputs(3)
    parOffsetFrmValCen = arrayOfInputs(4)
    parOperatorHeight = arrayOfInputs(5)
    parOperatorDiameter = arrayOfInputs(6)
    parHandwheelAngle = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)

    iOutput = 0

    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    'Declaration of Common objects
    Dim objInsValveBody As Object
    Dim oCenter As AutoMath.DPosition
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint As AutoMath.DPosition
    
    Set oCenter = New DPosition
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oDir = New AutoMath.DVector
    Set oPlacePoint = New DPosition

    'Declaration of Common variables
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dBodyHeight As Double
    Dim dBodyDia As Double
    Dim dOperatorDia As Double

    'Retrive the nozzle parameters
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, cptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, cptOffset2, depth2

    'Calculate the Face 1 to Center and Face 2 to Center based on the partdatabasis
    Select Case lPartDataBasis
    Case Is <= 1, 1052, 1053
        parFacetoFace = arrayOfInputs(2)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
    Case 1054
        parFacetoCenter = arrayOfInputs(3)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
    Case Else
        GoTo ErrorLabel
    End Select

    'Calculate maximum valve of pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2
    Dim dValvebodyDia As Double
    dValvebodyDia = IIf(CmpDblGreaterthan(BodyOD(pipeDiam1, flangeDiam1), BodyOD(pipeDiam2, flangeDiam2)), _
                        BodyOD(pipeDiam1, flangeDiam1), BodyOD(pipeDiam2, flangeDiam2)) + _
                        2 * parInsulationThickness
    
    'Create the premitives based on part databasis
    Select Case lPartDataBasis
    Case Is <= 1, 1052
        'Calculate the Body Height
        If CmpDblGreaterthan(parOperatorHeight, LINEAR_TOLERANCE) Then
            dBodyHeight = 0.6 * parOperatorHeight
        Else
            dBodyHeight = (dFace1toCenter + dFace2toCenter)
        End If
        dBodyHeight = dBodyHeight + parInsulationThickness
        
        'Calculate the Body Diameter
        If CmpDblGreaterthan(dFace1toCenter + dFace2toCenter - depth1 - depth2, LINEAR_TOLERANCE) Then
            dBodyDia = 0.7 * (dFace1toCenter + dFace2toCenter - depth1 - depth2)
        Else
            dBodyDia = 0.001
        End If
        dBodyDia = dBodyDia + 2 * parInsulationThickness

        'Valve Body
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, dBodyHeight, 0
        Set objInsValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBodyDia, True)

        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsValveBody
        
        'Create the Insulation for Valve Header
        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set dFace2toCenter, 0, 0
        Set objInsValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dValvebodyDia, True)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsValveBody
        Set objInsValveBody = Nothing
    
    Case 1053
        'Calculate the Body Height
        If CmpDblGreaterthan(parOperatorHeight, LINEAR_TOLERANCE) Then
            dBodyHeight = 0.4 * parOperatorHeight
        Else
            dBodyHeight = (dFace1toCenter + dFace2toCenter)
        End If
        dBodyHeight = dBodyHeight '+ parInsulationThickness

        'Calculate the Body Diameter
        If CmpDblGreaterthan(dFace1toCenter + dFace2toCenter - depth1 - depth2, LINEAR_TOLERANCE) Then
            dBodyDia = 0.9 * (dFace1toCenter + dFace2toCenter - depth1 - depth2)
        Else
            dBodyDia = 0.001
        End If
        dBodyDia = dBodyDia + 2 * parInsulationThickness
        
        'Calculate the Operator Diameter
        If CmpDblGreaterthan(parOperatorDiameter, LINEAR_TOLERANCE) Then
            dOperatorDia = parOperatorDiameter
        Else
            dOperatorDia = 0.9 * (dFace1toCenter + dFace2toCenter)
        End If

        dOperatorDia = dOperatorDia + 2 * parInsulationThickness

        'Valve Body 1
        oCenter.Set 0, 0, 0
        Set objInsValveBody = PlaceSphere(m_OutputColl, oCenter, 0.8 * dBodyDia / 2)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsValveBody
        Set objInsValveBody = Nothing

        'Valve Body 2
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, 0.8 * dBodyHeight, 0
        Set objInsValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBodyDia, True)

        'Set the Output
        m_OutputColl.AddOutput "InsulatedBody3", objInsValveBody
        Set objInsValveBody = Nothing

        'Valve Body 3
        oStPoint.Set 0, 0.8 * dBodyHeight, 0
        oEnPoint.Set 0, dBodyHeight, 0
        Set objInsValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dOperatorDia, True)

        'Set the Output
        m_OutputColl.AddOutput "InsulatedBody4", objInsValveBody
        Set objInsValveBody = Nothing
        
        'Create the Insulation for Valve Header
        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set dFace2toCenter, 0, 0
        Set objInsValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dValvebodyDia, True)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsValveBody
        Set objInsValveBody = Nothing
        
    Case 1054
        'Calculate the Body Height
        If CmpDblGreaterthan(parOperatorHeight, LINEAR_TOLERANCE) Then
            dBodyHeight = 0.6 * parOperatorHeight
        Else
            dBodyHeight = (dFace1toCenter + dFace2toCenter)
        End If
        dBodyHeight = dBodyHeight + parInsulationThickness
        
        'Calculate the Body Diameter
        If CmpDblGreaterthan(dFace1toCenter + dFace2toCenter - depth1 - depth2, LINEAR_TOLERANCE) Then
            dBodyDia = 0.7 * (dFace1toCenter + dFace2toCenter - depth1 - depth2)
        Else
            dBodyDia = 0.001
        End If
        dBodyDia = dBodyDia + 2 * parInsulationThickness
    
        'Valve Body 1
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, dBodyHeight, 0
        Set objInsValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBodyDia, True)

        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsValveBody
        
        'Create the Insulation for Valve Body 2
        dValvebodyDia = BodyOD(pipeDiam1, flangeDiam1) + 2 * parInsulationThickness
        oStPoint.Set -dFace1toCenter, 0, 0
        oEnPoint.Set 0, 0, 0
        Set objInsValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dValvebodyDia, True)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsValveBody
        Set objInsValveBody = Nothing
    
        'Create the Insulation for Valve Body 3
        dValvebodyDia = BodyOD(pipeDiam2, flangeDiam2) + 2 * parInsulationThickness
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, -dFace2toCenter, 0
        Set objInsValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dValvebodyDia, True)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "InsulatedBody3", objInsValveBody
        Set objInsValveBody = Nothing
    
    End Select
        
    Set objInsValveBody = Nothing
    Set oCenter = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
        
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
