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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.

'   ProgID          :  SP3DDiaphragmValveBody.DiaValBody
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Friday, 26 Oct 2007
'   Description     :  Saunders Standard Forged Diaphragm Valve
'   Change History  :
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   26-OCT-2007      PK        CR-127157:Created the symbol.
'   05-Feb-2008     VRK        CR:135745:Enhance Diaphragm Valve on face-to-face basis for JIS
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Physical:"    'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim sptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double
    Dim iOutput As Double

    Dim parFacetoFace As Double
    Dim parValveHeight As Double
    Dim parValveHeight1 As Double
    Dim parProjectionLength As Double
    Dim parHandwheelAngle As Double
    Dim parInsulationThickness As Double

    Dim oCentPos As AutoMath.DPosition
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set oCentPos = New AutoMath.DPosition
    
    iOutput = 0
    oCentPos.Set 0, 0, 0

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parValveHeight = arrayOfInputs(3)
'    parValveHeight1 = arrayOfInputs(4)
'    parProjectionLength = arrayOfInputs(5)

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                       sptOffset, depth

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis

    Select Case lPartDataBasis

    Case Is <= 1
        parFacetoFace = arrayOfInputs(2)
        parValveHeight = arrayOfInputs(3)
        parValveHeight1 = arrayOfInputs(4)
        parProjectionLength = arrayOfInputs(5)
        
        'Insert your code for the Cylindrical Body
        Dim oCylindricalBody As Object
        oStPoint.Set oCentPos.x - parFacetoFace / 2, oCentPos.y, oCentPos.z
        oEnPoint.Set oCentPos.x + parFacetoFace / 2, oCentPos.y, oCentPos.z
        Set oCylindricalBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

        'Set Output1(Cylindrical Body)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCylindricalBody
        Set oCylindricalBody = Nothing

        'Insert your code for body
        Dim oBody As Object
        Dim oTrapezium As IngrGeom3D.ComplexString3d
        Dim oAxis As AutoMath.DVector
        Dim lines As Collection
        Dim oLine As Object
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Dim dValue1 As Double
        Dim dValue2 As Double
        Dim dValue3 As Double

        Set lines = New Collection
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oAxis = New AutoMath.DVector

        dValue1 = parFacetoFace - 2 * parProjectionLength
        dValue2 = parValveHeight - parValveHeight1
        dValue3 = parValveHeight1 - 0.2 * parValveHeight

        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * (dValue1), _
                                                         oCentPos.y - (dValue2), oCentPos.z + 0.3 * (dValue1), oCentPos.x - 0.5 * _
                                                                                                               (dValue1), oCentPos.y - (dValue2), oCentPos.z - 0.3 * (dValue1))
        lines.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * (dValue1), _
                                                         oCentPos.y - (dValue2), oCentPos.z - 0.3 * (dValue1), oCentPos.x - 0.5 * _
                                                                                                               (dValue1), oCentPos.y + (dValue3), oCentPos.z - 0.4 * (dValue1))
        lines.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * (dValue1), _
                                                         oCentPos.y + (dValue3), oCentPos.z - 0.4 * (dValue1), oCentPos.x - 0.5 * _
                                                                                                               (dValue1), oCentPos.y + (dValue3), oCentPos.z + 0.4 * (dValue1))
        lines.Add oLine
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * (dValue1), _
                                                         oCentPos.y + (dValue3), oCentPos.z + 0.4 * (dValue1), oCentPos.x - 0.5 * _
                                                                                                               (dValue1), oCentPos.y - (dValue2), oCentPos.z + 0.3 * (dValue1))
        lines.Add oLine
        oStPoint.Set oCentPos.x - 0.5 * (dValue1), oCentPos.y - (dValue2), oCentPos.z + 0.3 * _
                                                                           (dValue1)

        Set oTrapezium = PlaceTrCString(oStPoint, lines)

        oAxis.Set 1, 0, 0
        Set oBody = PlaceProjection(m_OutputColl, oTrapezium, oAxis, dValue1, True)

        'Set Output2(Body)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBody
        Set oBody = Nothing
        Set oGeomFactory = Nothing
        Set oAxis = Nothing
        Set oTrapezium = Nothing
        Set oLine = Nothing
        Set lines = Nothing

        'Insert your code for Operator Base
        Dim oOperatorBase As Object
        oStPoint.Set oCentPos.x - (dValue1) / 2, oCentPos.y + (dValue3), oCentPos.z + (dValue1) / 2
        oEnPoint.Set oCentPos.x + (dValue1) / 2, oCentPos.y + parValveHeight1, oCentPos.z - _
                                                                               (dValue1) / 2

        Set oOperatorBase = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set Output3(Operator Base)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oOperatorBase
        Set oOperatorBase = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        
    Case 5
    
        parFacetoFace = arrayOfInputs(2)
        
        'Cone1(Output 1)
        Dim objCone As Object
        oStPoint.Set oCentPos.x - (0.5 * parFacetoFace - flangeThick), oCentPos.y, oCentPos.z
        oEnPoint.Set oCentPos.x, oCentPos.y, oCentPos.z
        Set objCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam / 2, 0, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
        Set objCone = Nothing
        
        'Cone2(Output 2)
        oStPoint.Set oCentPos.x + (0.5 * parFacetoFace - flangeThick), oCentPos.y, oCentPos.z
        oEnPoint.Set oCentPos.x, oCentPos.y, oCentPos.z
        Set objCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam / 2, 0, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
        Set objCone = Nothing
        
        'Insert your code for output 3(Sphere)
        Dim ObjValveBody As Object
        Set ObjValveBody = PlaceSphere(m_OutputColl, oCentPos, parFacetoFace / 8)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
        Set ObjValveBody = Nothing
        
    Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.
        
    End Select
    
        Set oStPoint = Nothing
        Set oEnPoint = Nothing


        'Place Nozzle1
        Dim oPosition As AutoMath.DPosition
        Dim oVec As AutoMath.DVector
        Dim oNozzle As GSCADNozzleEntities.IJDNozzle

        Set oPosition = New AutoMath.DPosition
        Set oVec = New AutoMath.DVector

        oVec.Set -1, 0, 0
        oPosition.Set oCentPos.x - parFacetoFace / 2 + sptOffset - depth, oCentPos.y, oCentPos.z

        Set oNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oVec, oPosition)

        'Set Output4(Nozzle1)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle

        Set oVec = Nothing
        Set oPosition = Nothing
        Set oNozzle = Nothing

        'Place Nozzle2
        Set oVec = New AutoMath.DVector
        Set oPosition = New AutoMath.DPosition

        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

        oVec.Set 1, 0, 0
        oPosition.Set oCentPos.x + parFacetoFace / 2 - sptOffset + depth, oCentPos.y, oCentPos.z
        Set oNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oVec, oPosition)

        'Set Output5(Nozzle2)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle

        Set oVec = Nothing
        Set oPosition = Nothing
        Set oNozzle = Nothing

        'Insert your code for valve operator
        Dim oDirX As IJDVector
        Dim oDirY As IJDVector
        Dim oDirZ As IJDVector
        Dim oSymbolHelper As IJSymbolGeometryHelper
        Dim oOperatorPart As IJDPart
        Dim oOperatorOcc As IJPartOcc
        Dim OpOrigin As IJDPosition

        Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
        oSymbolHelper.OutputCollection = m_OutputColl

        On Error Resume Next

        Set oDirX = New AutoMath.DVector
        Set oDirY = New AutoMath.DVector
        Set oDirZ = New AutoMath.DVector

        oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
        oDirY.Set 0, 1, 0
        oDirZ.Set Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)

        On Error GoTo ErrorLabel

        Dim oIJDInputsArg As IJDInputsArg
        Dim oIJDEditJDArgument As IJDEditJDArgument

        Set OpOrigin = New AutoMath.DPosition

        OpOrigin.Set oCentPos.x, oCentPos.y, oCentPos.z
        Set oPipeComponent = oPartFclt
        If Not oPipeComponent Is Nothing Then
            Set oOperatorPart = oPipeComponent.GetValveOperatorPart
            If Not oOperatorPart Is Nothing Then
                Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", _
                                                                   oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
                If lPartDataBasis <= 1 Then
                    Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)
                    UpdateOperatorAttributes oOperatorOcc, "IJUAWidth", "Width", _
                                              (dValue1), oIJDEditJDArgument
                    UpdateOperatorAttributes oOperatorOcc, "IJUAInstrumentPositioner", _
                                         "PositionerOffset", parValveHeight1, oIJDEditJDArgument
                End If
                Set oIJDInputsArg = oOperatorOcc
                oIJDInputsArg.Update
            End If
        End If

        Set oDirX = Nothing
        Set oDirY = Nothing
        Set oDirZ = Nothing
        Set oCentPos = Nothing
        Set oSymbolHelper = Nothing
        Set oPipeComponent = Nothing
        Set oOperatorPart = Nothing
        Set oOperatorOcc = Nothing
        Set oIJDInputsArg = Nothing
        Set oIJDEditJDArgument = Nothing
        Set OpOrigin = Nothing
        Exit Sub

ErrorLabel:
        ReportUnanticipatedError MODULE, METHOD
        Resume Next

    End Sub

Private Sub UpdateOperatorAttributes(oOccurrence As Object, strInterfaceName As String, strAttribNameOnInterface As String, _
                                     vInputParameterValue As Variant, oIJDEditJDArgument As IJDEditJDArgument)

    Const METHOD = "UpdateOperatorAttributes"
    On Error GoTo ErrorLabel

    Dim oIJDAttributes As IJDAttributes
    Dim oIJDAttr As IJDAttribute
    Dim oAttribCol As IJDAttributesCol
    Dim lInputIndex As Long
    Dim oSymbol As IJDSymbol
    Dim oInputs As IJDInputs
    Dim oInput As IJDInput
    Dim lCount As Long
    Dim lCount1 As Long
    'Dim vNewValue       As Variant

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

    'Set value of the user attribute
    Set oIJDAttributes = oOccurrence
    Set oAttribCol = oIJDAttributes.CollectionOfAttributes(strInterfaceName)
    Set oIJDAttr = oAttribCol.Item(strAttribNameOnInterface)
    'vNewValue = dInputParameter
    oIJDAttr.Value = vInputParameterValue

    'Provide this value as input argument to the operator symbol
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


