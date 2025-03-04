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
'   Author:         MA
'   Creation Date:  Monday, Oct 15 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is developed for CR-127627
'   Source: Saunders Aseptic Diaphragm Valves, ZDT Zero Deadleg ‘T’ Pattern
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private Const E_NOOPERATORDEFINED = &H80040400 'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl

Private Sub Class_Initialize()
'''
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
    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parEndLength As Double
    Dim parBranchLength As Double
    Dim parPortVOffset As Double
    Dim parOpCentertoEnd As Double
    Dim parValCentertoOpCenter As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle As Double
    Dim parPositionerOffset As Double
    Dim parWidth As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parFacetoCenter = arrayOfInputs(3)
'    parFace1toCenter = arrayOfInputs(4)
'    parFace2toCenter = arrayOfInputs(5)
    parEndLength = arrayOfInputs(6)
    parBranchLength = arrayOfInputs(7)
    parPortVOffset = arrayOfInputs(8)
    parOpCentertoEnd = arrayOfInputs(9)
    parValCentertoOpCenter = arrayOfInputs(10)
    parInsulationThickness = arrayOfInputs(11)
    parHandwheelAngle = arrayOfInputs(12)
    parPositionerOffset = arrayOfInputs(13)
    parWidth = arrayOfInputs(14)
    
    Dim objCyl As Object
    Dim objBox As Object
    Dim dCyldia As Double
    
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                        sptOffset, depth

 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -5): FacetoFace
'   (Part data Basis value -10): FacetoCenter
'   (Part data Basis value -13): Face1toCenter and Face2toCenter
    
    Select Case lPartDataBasis
    
        Case Is <= 1, 5
            parFacetoFace = arrayOfInputs(2)
            dFace1toCenter = parFacetoFace / 2
            dFace2toCenter = parFacetoFace / 2
    
        Case 10
            parFacetoCenter = arrayOfInputs(3)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
        
        Case 13
            parFace1toCenter = arrayOfInputs(4)
            parFace2toCenter = arrayOfInputs(5)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
    
        Case Else
            GoTo ErrorLabel:
        
    End Select
    
    'Insert your code for output 1(Cylinder 1)
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set -dFace1toCenter, 0, 0
    oEnPoint.Set dFace2toCenter, 0, 0
    dCyldia = pipeDiam

    Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
    Set objCyl = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    'Insert your code for output 2(Box 1)
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set -(dFace1toCenter - parEndLength), 2.5 * parPortVOffset, _
                                    -parValCentertoOpCenter
    oEnPoint.Set (dFace2toCenter - parEndLength), -parValCentertoOpCenter, _
                                    0.8 * parValCentertoOpCenter
                                    
    Set objBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBox
    Set objBox = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    'Insert your code for output 3(Box 2)
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set -(dFace1toCenter - parEndLength), 2.5 * parPortVOffset, _
                                    0.8 * parValCentertoOpCenter
    oEnPoint.Set (dFace2toCenter - parEndLength), _
                            -2 * parPortVOffset, _
                    (parValCentertoOpCenter + parOpCentertoEnd - parBranchLength)

    Set objBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBox
    Set objBox = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    'Insert your code for output 4(Cylinder 2)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                        sptOffset, depth
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set 0, parPortVOffset, _
                        (parValCentertoOpCenter + parOpCentertoEnd - parBranchLength)
    oEnPoint.Set 0, parPortVOffset, (parValCentertoOpCenter + parOpCentertoEnd)
    dCyldia = pipeDiam

    Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
    Set objCyl = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    'Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                        sptOffset, depth
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -dFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0

    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    'Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                        sptOffset, depth
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set dFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    'Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                        sptOffset, depth
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set 0, parPortVOffset, _
                    (parValCentertoOpCenter + parOpCentertoEnd) + sptOffset - depth
    oDir.Set 0, 0, 1
    
    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
   
' Insert your code for output 8 (Valve Operator)
    Dim oSymbolHelper As IJSymbolGeometryHelper
    Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
    oSymbolHelper.OutputCollection = m_OutputColl

    On Error Resume Next
    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector

    oDirX.Set Sin(parHandwheelAngle), 0, -Cos(parHandwheelAngle)
    oDirY.Set 0, 1, 0
    oDirZ.Set Cos(parHandwheelAngle), Sin(parHandwheelAngle), 0
    
    Set oPipeComponent = oPartFclt
    On Error GoTo ErrorLabel
    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc   As IJPartOcc
    If Not oPipeComponent Is Nothing Then
'       Note: oPipeComponent.GetValveOperatorPart method call may fail in following two cases:
'        1. If valve operator catalog part number is not defined (to check for E_NOOPERATORDEFINED) or
'        2. If valve operator part is not found (to check for E_NOOPERATORPARTFOUND)
'        symbol should raise error for case 2.
        On Error GoTo HandleOperatorCodeLabel
        Set oOperatorPart = oPipeComponent.GetValveOperatorPart
        If Not oOperatorPart Is Nothing Then
            Dim OpOrigin As IJDPosition
            Set OpOrigin = New DPosition
            OpOrigin.Set 0, 0, parValCentertoOpCenter
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
        End If
    End If
HandleOperatorCodeLabel:
'   The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
        If Err.Number = E_NOOPERATORPARTFOUND Then    'Case 2
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If

    '' Update Occurrence Attributes of Operator.
    Dim oIJDInputsArg As IJDInputsArg
    Dim oIJDEditJDArgument As IJDEditJDArgument

    Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)

    UpdateOperatorAttributes oOperatorOcc, "IJUAInstrumentPositioner", "PositionerOffset", 2.5 * parPortVOffset, oIJDEditJDArgument
    
    If lPartDataBasis <= 1 Or lPartDataBasis = 5 Then
    UpdateOperatorAttributes oOperatorOcc, "IJUAWidth", "Width", 2 * (dFace1toCenter - parEndLength), oIJDEditJDArgument
    ElseIf lPartDataBasis = 10 Then
    UpdateOperatorAttributes oOperatorOcc, "IJUAWidth", "Width", 2 * (dFace1toCenter - parEndLength), oIJDEditJDArgument
    ElseIf lPartDataBasis = 13 Then
    UpdateOperatorAttributes oOperatorOcc, "IJUAWidth", "Width", (dFace1toCenter + dFace2toCenter) - 2 * parEndLength, oIJDEditJDArgument
    End If

    Set oIJDInputsArg = oOperatorOcc
    oIJDInputsArg.Update

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
   ReportUnanticipatedError MODULE, METHOD

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
   ReportUnanticipatedError MODULE, METHOD

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
   ReportUnanticipatedError MODULE, METHOD

End Function



