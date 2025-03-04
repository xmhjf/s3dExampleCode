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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
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
Private Const MODULE = "Physical:"    'Used for error messages
Private Const E_NOOPERATORDEFINED = &H80040400    'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401    'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl

Private Sub Class_Initialize()
    '''
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

    'Declaration of Common objects
    Dim objValveBody As Object
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

    'Create the premitives based on part databasis
    Select Case lPartDataBasis
    Case Is <= 1, 1052
        'Calculate the Body Height
        If CmpDblGreaterthan(parOperatorHeight, LINEAR_TOLERANCE) Then
            dBodyHeight = 0.6 * parOperatorHeight
        Else
            dBodyHeight = (dFace1toCenter + dFace2toCenter)
        End If

        'Calculate the Body Diameter
        If CmpDblGreaterthan(dFace1toCenter + dFace2toCenter - depth1 - depth2, LINEAR_TOLERANCE) Then
            dBodyDia = 0.7 * (dFace1toCenter + dFace2toCenter - depth1 - depth2)
        Else
            dBodyDia = 0.001
        End If
        
        'Valve Body
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, dBodyHeight, 0
        Set objValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBodyDia, True)

        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveBody

    Case 1053
        'Calculate the Body Height
        If CmpDblGreaterthan(parOperatorHeight, LINEAR_TOLERANCE) Then
            dBodyHeight = 0.4 * parOperatorHeight
        Else
            dBodyHeight = (dFace1toCenter + dFace2toCenter)
        End If

        'Calculate the Body Diameter
        If CmpDblGreaterthan(dFace1toCenter + dFace2toCenter - depth1 - depth2, LINEAR_TOLERANCE) Then
            dBodyDia = 0.9 * (dFace1toCenter + dFace2toCenter - depth1 - depth2)
        Else
            dBodyDia = 0.001
        End If
        

        'Calculate the Operator Diameter
        If CmpDblGreaterthan(parOperatorDiameter, LINEAR_TOLERANCE) Then
            dOperatorDia = parOperatorDiameter
        Else
            dOperatorDia = 0.9 * (dFace1toCenter + dFace2toCenter)
        End If

        'Valve Body 1
        oCenter.Set 0, 0, 0
        Set objValveBody = PlaceSphere(m_OutputColl, oCenter, 0.8 * dBodyDia / 2)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveBody
        Set objValveBody = Nothing

        'Valve Body 2
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, 0.8 * dBodyHeight, 0
        Set objValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBodyDia, True)

        'Set the Output
        m_OutputColl.AddOutput "ValveBody2", objValveBody
        Set objValveBody = Nothing

        'Valve Body 3
        oStPoint.Set 0, 0.8 * dBodyHeight, 0
        oEnPoint.Set 0, dBodyHeight, 0
        Set objValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dOperatorDia, True)

        'Set the Output
        m_OutputColl.AddOutput "ValveBody3", objValveBody
        Set objValveBody = Nothing

    Case 1054
        'Calculate the Body Height
        If CmpDblGreaterthan(parOperatorHeight, LINEAR_TOLERANCE) Then
            dBodyHeight = 0.6 * parOperatorHeight
        Else
            dBodyHeight = (dFace1toCenter + dFace2toCenter)
        End If

        'Calculate the Body Diameter
        If CmpDblGreaterthan(dFace1toCenter + dFace2toCenter - depth1 - depth2, LINEAR_TOLERANCE) Then
            dBodyDia = 0.7 * (dFace1toCenter + dFace2toCenter - depth1 - depth2)
        Else
            dBodyDia = 0.001
        End If
        
        'Valve Body
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, dBodyHeight, 0
        Set objValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBodyDia, True)

        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveBody

    End Select

    'Place Nozzle 1
    oDir.Set -1, 0, 0
    oPlacePoint.Set -(dFace1toCenter + cptOffset1 - depth1), 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, _
                                           dFace1toCenter)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    'Place Nozzle 2
    If lPartDataBasis = 1052 Or lPartDataBasis = 1053 Then
        oPlacePoint.Set dFace2toCenter + cptOffset2 - depth2, 0, 0
        oDir.Set 1, 0, 0
    ElseIf lPartDataBasis = 1054 Then
        oPlacePoint.Set 0, -(dFace2toCenter + cptOffset2 - depth2), 0
        oDir.Set 0, -1, 0
    End If
        
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, _
                                           dFace2toCenter)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    Set objValveBody = Nothing
    Set oCenter = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing

    '****************************************************************************************************
    'Operator
    '****************************************************************************************************
    'Insert your code for Valve Operator
    Dim oSymbolHelper As IJSymbolGeometryHelper
    Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
    oSymbolHelper.OutputCollection = m_OutputColl

    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector

    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector

    oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
    oDirY.Set 0, 1, 0
    oDirZ.Set -Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)

    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc As IJPartOcc

    If Not oPipeComponent Is Nothing Then
        'Note: oPipeComponent.GetValveOperatorPart method call may fail in following two cases:
        '1. If valve operator catalog part number is not defined (to check for E_NOOPERATORDEFINED) or
        '2. If valve operator part is not found (to check for E_NOOPERATORPARTFOUND)
        'symbol should raise error for case 2.
        On Error GoTo HandleOperatorCodeLabel
        Set oOperatorPart = oPipeComponent.GetValveOperatorPart
        If Not oOperatorPart Is Nothing Then
            Dim OpOrigin As IJDPosition
            Set OpOrigin = New DPosition
            OpOrigin.Set 0, 0, 0
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, _
                                                                OpOrigin, oDirX, oDirY, oDirZ)
            Set OpOrigin = Nothing
            'Update Occurrence Attributes of Operator.
            Dim oIJDInputsArg As IJDInputsArg
            Dim oIJDEditJDArgument As IJDEditJDArgument

            Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)

            'Update the Occurence Operator Attributes
            'Offset From Valve Centerline
            UpdateOperatorAttributes oOperatorOcc, "IJUAOffsetFrmValCen", "OffsetFrmValCen", _
                                     parOffsetFrmValCen, oIJDEditJDArgument

            'Operator Height
            If CmpDblGreaterthan(parOperatorHeight, LINEAR_TOLERANCE) Then
                UpdateOperatorAttributes oOperatorOcc, "IJUAValveOperator", "OperatorHeight", _
                                         parOperatorHeight, oIJDEditJDArgument
            End If

'            'Handwheel Offset from Valve Centerline
'            If CmpDblGreaterthan(parOperatorDiameter, LINEAR_TOLERANCE) Then
'                UpdateOperatorAttributes oOperatorOcc, "IJUAValveOperator", "OperatorDiameter", _
'                                         parOperatorDiameter, oIJDEditJDArgument
'            End If
            Set oIJDInputsArg = oOperatorOcc
            oIJDInputsArg.Update
            Set oIJDEditJDArgument = Nothing
            Set oIJDInputsArg = Nothing

        End If
    End If
    
    Set oDirX = Nothing
    Set oDirY = Nothing
    Set oDirZ = Nothing

HandleOperatorCodeLabel:
    '   The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
    If Err.Number = E_NOOPERATORPARTFOUND Then          'Case 2
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If

    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub UpdateOperatorAttributes(oOccurrence As Object, strInterfaceName As String, strAttribNameOnInterface As String, _
                                     vInputParameterValue As Variant, oIJDEditJDArgument As IJDEditJDArgument)

    Const METHOD = "UpdateOperatorAttributes"
    On Error GoTo ErrorLabel

    Dim oIJDAttributes As IJDAttributes
    Dim oIJDAttr As IJDAttribute
    Dim oAttribCol As IJDAttributesCol
    '    Dim vNewValue As Variant
    Dim lInputIndex As Long

    Dim oSymbol As IJDSymbol
    Dim oInputs As IJDInputs
    Dim oInput As IJDInput
    Dim lCount As Long
    Dim lCount1 As Long

    ''To find the index number of the Attibute to be updated in the operator code.
    Set oSymbol = oOccurrence
    Set oInputs = oSymbol.IJDSymbolDefinition(1).IJDInputs
    lCount = oInputs.Count
    For lCount1 = 1 To lCount
        Set oInput = oInputs.Item(lCount1)
        If StrComp(oInput.name, strAttribNameOnInterface, vbTextCompare) = 0 Then
            lInputIndex = oInput.index
            Exit For
        End If
    Next lCount1

    ''If lInputIndex  is ZERO then operator doesnt have the attribute to be updated as it's InputParameter,
    ''so that no need for futher excecution.
    If lInputIndex = 0 Then Exit Sub

    ' Set value of the user attribute
    Set oIJDAttributes = oOccurrence
    Set oAttribCol = oIJDAttributes.CollectionOfAttributes(strInterfaceName)
    Set oIJDAttr = oAttribCol.Item(strAttribNameOnInterface)
    '    vNewValue = dInputParameter
    oIJDAttr.Value = vInputParameterValue

    ' Provide this value as input argument to the operator symbol
    UpdateSymbolArgument CDbl(vInputParameterValue), lInputIndex, oIJDEditJDArgument

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, MODULE & METHOD

End Sub

Private Sub UpdateSymbolArgument(dblValue As Double, lOperatorAttributeArgumentIndex As Long, _
                                 oIJDEditJDArgument As IJDEditJDArgument)

    Const METHOD = "UpdateSymbolArgument"
    On Error GoTo ErrorLabel

    Dim oPC As IJDParameterContent
    Dim oArgument As IJDArgument

    Set oPC = New DParameterContent
    oPC.Type = igValue
    oPC.UomType = 0
    oPC.uomValue = dblValue

    Set oArgument = New DArgument
    oArgument.index = lOperatorAttributeArgumentIndex
    oArgument.Entity = oPC

    oIJDEditJDArgument.SetArg oArgument

    Set oArgument = Nothing
    Set oPC = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, MODULE & METHOD

End Sub

Private Function GetIJDEditJDArgument(oOperatorOcc As IJPartOcc) As IJDEditJDArgument

    Const METHOD = "GetIJDEditJDArgument"
    On Error GoTo ErrorLabel

    Dim oIJDEditJDArgument As IJDEditJDArgument
    Dim oIEnumJDArgument As IEnumJDArgument
    Dim oSymbol As IJDSymbol
    Dim oIJDValuesArg As IJDValuesArg

    Set oSymbol = oOperatorOcc
    Set oIJDValuesArg = oSymbol.IJDValuesArg
    Set oIEnumJDArgument = oIJDValuesArg.GetValues(igINPUT_ARGUMENTS_SET)
    Set oIJDEditJDArgument = oIEnumJDArgument
    Set GetIJDEditJDArgument = oIJDEditJDArgument

    Exit Function

ErrorLabel:
    Err.Raise Err.Number, MODULE & METHOD
End Function


