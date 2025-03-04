Attribute VB_Name = "CanRuleAttrMgmtUtil"
'   July 17, 2009       GG   TR #166187 Attribute management code needs to validate user keyed in values and update related values
'   July 24, 2009       GG   TR #166187 Added support for multi-editing
Option Explicit
Const MODULE = "CanRuleAttrMgmtUtil"
Public Sub RuleSetReadOnly(vRuleValue As Variant, oID As IJAttributeDescriptor, oOD As IJAttributeDescriptor)
    Const METHOD = "RuleSetReadOnly"
    On Error GoTo ErrorHandler
    If IsNull(vRuleValue) Then
        oID.AttrState = AttributeDescriptor_ReadOnly
        oOD.AttrState = AttributeDescriptor_ReadOnly
        Exit Sub
    End If
    Dim ruleValue As Long
    ruleValue = vRuleValue
    Select Case ruleValue
        Case DiameterRule_ByOD:
            oID.AttrState = AttributeDescriptor_ReadOnly
            oOD.AttrState = AttributeDescriptor_ReadOnly
        Case DiameterRule_ByID:
            oOD.AttrState = AttributeDescriptor_ReadOnly
            oID.AttrState = AttributeDescriptor_ReadOnly
        Case DiameterRule_User:
            oID.AttrState = AttributeDescriptor_ReadOnly
            removeReadOnly oOD
    End Select
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub LMethodSetReadOnly(vMethod As Variant, oLength As IJAttributeDescriptor, oFactor As IJAttributeDescriptor)
    Const METHOD = "LMethodSetReadOnly"
    On Error GoTo ErrorHandler
    If IsNull(vMethod) Then
        oFactor.AttrState = AttributeDescriptor_ReadOnly
        oLength.AttrState = AttributeDescriptor_ReadOnly
        Exit Sub
    End If
    Dim LMethod As Long
    LMethod = vMethod
    Select Case (LMethod Mod 2)
        Case 0: oFactor.AttrState = AttributeDescriptor_ReadOnly
                removeReadOnly oLength
        Case 1: oLength.AttrState = AttributeDescriptor_ReadOnly
                removeReadOnly oFactor
    End Select
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub ConeMethodSetReadOnly(vConeMethod As Variant, oLength As IJAttributeDescriptor, oSlope As IJAttributeDescriptor, oAngle As IJAttributeDescriptor)
    Const METHOD = "ConeMethodSetReadOnly"
    On Error GoTo ErrorHandler
    If IsNull(vConeMethod) Then
        oAngle.AttrState = AttributeDescriptor_ReadOnly
        oLength.AttrState = AttributeDescriptor_ReadOnly
        oSlope.AttrState = AttributeDescriptor_ReadOnly
        Exit Sub
    End If
    Dim coneMethod As Long
    coneMethod = vConeMethod
    Select Case (coneMethod)
        Case ConeMethod_Slope:  oAngle.AttrState = AttributeDescriptor_ReadOnly
                                oLength.AttrState = AttributeDescriptor_ReadOnly
                                removeReadOnly oSlope
                                
        Case ConeMethod_Angle:  oSlope.AttrState = AttributeDescriptor_ReadOnly
                                oLength.AttrState = AttributeDescriptor_ReadOnly
                                removeReadOnly oAngle
                                
        Case ConeMethod_Length: oSlope.AttrState = AttributeDescriptor_ReadOnly
                                oAngle.AttrState = AttributeDescriptor_ReadOnly
                                removeReadOnly oLength
    End Select
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub removeReadOnly(oAttribute As IJAttributeDescriptor)
    ' AttributeDescriptor_ReadOnly = 15 which is the last four bits. We want to force those to 0
    ' while every thing else stays a 1 if it already is a 1 so Not AttributeDescriptor_ReadOnly sets those
    ' last four bits to 0 and everything else is to 1 so the and operations perseves all ones except the last four.
    If Not oAttribute Is Nothing Then
        oAttribute.AttrState = (oAttribute.AttrState And (Not AttributeDescriptor_ReadOnly))
    End If
End Sub

' Sets all attributes to read only when attribute management is working on a well defined can
' The return value is true for a well defined can rule and true for a custom can rule
Public Function ifWellDefinedSetReadOnly(ByRef strDefinition As String, oAttributes As Collection) As Boolean
    Const METHOD = "ifWellDefinedSetReadOnly"
    On Error GoTo ErrorHandler
    
    ifWellDefinedSetReadOnly = False
    
    Dim oAttr As IJAttributeDescriptor
    If Not isCustomCan(strDefinition) Then
        For Each oAttr In oAttributes
            oAttr.AttrState = AttributeDescriptor_ReadOnly
        Next oAttr
        
        ifWellDefinedSetReadOnly = True
    End If
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Function isCustomCan(ByRef strDefinition As String) As Boolean
    
    Const METHOD = "isCustomCan"
    On Error GoTo ErrorHandler
    isCustomCan = False

    Dim dValue As Double
    
    dValue = GetUserAttrDefault(strDefinition, "IJUASMCanRuleDiameter", "CanOutsideDiameter")
    isCustomCan = Abs(dValue) < 0.000001
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Function GetUserAttrDefault(strDefinition As String, strInterface As String, strAttribute As String) As Variant
    Const METHOD = "GetUserAttrDefault"
    On Error GoTo ErrorHandler
    GetUserAttrDefault = Empty
    
    Dim oConnection         As IJDPOM
    Dim oItem               As Object
    
    Set oConnection = GetCatalogDBConnection()

    Set oItem = oConnection.GetObject(strDefinition)
    
    Dim oAttributeMetadata  As IJDAttributeMetaData
    Dim oItemAttributes     As IJDAttributes
    Dim oAttributesCol      As IJDAttributesCol
    Dim oAttribute          As IJDAttribute
    Dim vInterfaceID        As Variant
    
    Set oAttributeMetadata = oConnection
    
    vInterfaceID = oAttributeMetadata.IID(strInterface)
    
    If Not oItem Is Nothing Then
        Set oItemAttributes = oItem
        Set oAttributesCol = oItemAttributes.CollectionOfAttributes(vInterfaceID)
        Set oAttribute = oAttributesCol.Item(strAttribute)
        GetUserAttrDefault = oAttribute.Value
    Else
        Debug.Assert "failed to resolve " & strDefinition
    End If
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function
Public Function ValidateAttribute(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal oAttrDescr As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
    Const METHOD = "ValidateAttribute"
    On Error GoTo ErrorHandler
    Dim oLocalizer As IJLocalizer
    Set oLocalizer = New IMSLocalizer.Localizer
    oLocalizer.Initialize App.Path & "\" & App.EXEName
    
    ValidateAttribute = vbNullString
    Dim strTemp As String
    strTemp = oAttrDescr.attrName
    If oAttrDescr.attrName = attrL2CompMethod _
       Or oAttrDescr.attrName = attrL3CompMethod _
       Or oAttrDescr.attrName = attrConeLengthMethod _
       Or oAttrDescr.attrName = attrCone1LengthMethod _
       Or oAttrDescr.attrName = attrCone2LengthMethod Then
         If CInt(varNewAttrValue) < 1 Then
            ValidateAttribute = oLocalizer.GetString(IDS_SELECT_METHOD, "You must select a method")
        End If
    ElseIf oAttrDescr.attrName = attrL2Length _
       Or oAttrDescr.attrName = attrL3Length _
       Or oAttrDescr.attrName = attrConeLength _
       Or oAttrDescr.attrName = attrCone1Length _
       Or oAttrDescr.attrName = attrCone2Length _
       Or oAttrDescr.attrName = attrMinimumExtensionDistance _
       Or oAttrDescr.attrName = attrRoundoffDistance Then
        If CDbl(varNewAttrValue) < 0 Then
            ValidateAttribute = oLocalizer.GetString(IDS_DISTANCE_MUSTBE_POSITIVE, "Distance must be a positive number")
        End If
    ElseIf oAttrDescr.attrName = attrL2Factor _
       Or oAttrDescr.attrName = attrL3Factor Then
        If CDbl(varNewAttrValue) < dTol Or CDbl(varNewAttrValue) > 100 Then
            ValidateAttribute = oLocalizer.GetString(IDS_FACTOR_RANGE, "Factor range is 0.000001 to 100")
        End If
    ElseIf oAttrDescr.attrName = attrConeSlope _
       Or oAttrDescr.attrName = attrCone1Slope _
       Or oAttrDescr.attrName = attrCone2Slope _
       Or oAttrDescr.attrName = attrChamferSlope Then
        If CDbl(varNewAttrValue) < 0 Then
            ValidateAttribute = oLocalizer.GetString(IDS_SLOPE_MUSTBE_POSITIVE, "Slope must be a positive number")
        End If
    ElseIf oAttrDescr.attrName = attrConeAngle _
       Or oAttrDescr.attrName = attrCone1Angle _
       Or oAttrDescr.attrName = attrCone2Angle Then
        If CDbl(varNewAttrValue) < 0 Or CDbl(varNewAttrValue) > PI / 2 Then
            ValidateAttribute = oLocalizer.GetString(IDS_ANGLE_RANGE, "Angle range is 0 to 90 degrees")
        End If
   End If
   Set oLocalizer = Nothing

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Function UpdateAttributeValues(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal oAttrDescr As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
    Const METHOD = "UpdateAttributeValues"
    On Error GoTo ErrorHandler

    Dim oTempAttrDescr As IJAttributeDescriptor
    Dim oCanOD As IJAttributeDescriptor
    Dim oCanID As IJAttributeDescriptor
    Dim oCanThickness As IJAttributeDescriptor
    UpdateAttributeValues = vbNullString
    Dim vCanOD As Variant
    Dim vCanThickness As Variant
    For Each oTempAttrDescr In CollAllDisplayedValues
        If oTempAttrDescr.attrName = attrOuterDiameter Then
            Set oCanOD = oTempAttrDescr
            vCanOD = oTempAttrDescr.AttrValue
        ElseIf oTempAttrDescr.attrName = attrInnerDiameter Then
            Set oCanID = oTempAttrDescr
        ElseIf oTempAttrDescr.attrName = attrCanThickness Then
            Set oCanThickness = oTempAttrDescr
            vCanThickness = oTempAttrDescr.AttrValue
        End If
    Next oTempAttrDescr
    ' update values
    Dim oLocalizer As IJLocalizer
    Set oLocalizer = New IMSLocalizer.Localizer
    oLocalizer.Initialize App.Path & "\" & App.EXEName
   
    If oAttrDescr.attrName = attrCanThickness Then
        'only check for user-defined
        If (oCanOD.AttrState <> AttributeDescriptor_ReadOnly) And vCanOD < 2 * CDbl(varNewAttrValue) Then
            UpdateAttributeValues = oLocalizer.GetString(IDS_OD_GREATER_THAN_2THICKNESS, "Can outside diameter should not be less than 2 times the Can thickness")
            Exit Function
        End If
        oCanID.AttrValue = vCanOD - 2 * CDbl(varNewAttrValue)
    ElseIf oAttrDescr.attrName = attrOuterDiameter Then
        'only check for user-defined
        If (oCanOD.AttrState <> AttributeDescriptor_ReadOnly) And CDbl(varNewAttrValue) < 2 * vCanThickness Then
            UpdateAttributeValues = oLocalizer.GetString(IDS_OD_GREATER_THAN_2THICKNESS, "Can outside diameter should not be less than 2 times the Can thickness")
            Exit Function
        End If
        oCanID.AttrValue = CDbl(varNewAttrValue) - 2 * vCanThickness
    End If
    Set oLocalizer = Nothing
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Function UpdateAttributeReadOnly(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal oAttrDescr As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
    Const METHOD = "UpdateAttributeReadOnly"
    On Error GoTo ErrorHandler

    Dim oTempAttrDescr As IJAttributeDescriptor
    
    Dim oCanOD As IJAttributeDescriptor
    Dim oCanID As IJAttributeDescriptor
    
    Dim oL2Factor As IJAttributeDescriptor
    Dim oL2Length As IJAttributeDescriptor
    
    Dim oL3Factor As IJAttributeDescriptor
    Dim oL3Length As IJAttributeDescriptor
    
    Dim oConeSlope  As IJAttributeDescriptor
    Dim oConeLength As IJAttributeDescriptor
    Dim oConeAngle  As IJAttributeDescriptor
    
    Dim oCone1Slope  As IJAttributeDescriptor
    Dim oCone1Length As IJAttributeDescriptor
    Dim oCone1Angle  As IJAttributeDescriptor
    
    Dim oCone2Slope  As IJAttributeDescriptor
    Dim oCone2Length As IJAttributeDescriptor
    Dim oCone2Angle  As IJAttributeDescriptor
    
    UpdateAttributeReadOnly = vbNullString
    
    For Each oTempAttrDescr In CollAllDisplayedValues
        Select Case oTempAttrDescr.attrName
            Case attrOuterDiameter:
                Set oCanOD = oTempAttrDescr
            Case attrInnerDiameter:
                Set oCanID = oTempAttrDescr
            Case attrL2Factor:
                Set oL2Factor = oTempAttrDescr
            Case attrL2Length:
                Set oL2Length = oTempAttrDescr
            Case attrL3Factor:
                Set oL3Factor = oTempAttrDescr
            Case attrL3Length:
                Set oL3Length = oTempAttrDescr
            Case attrConeSlope:
                Set oConeSlope = oTempAttrDescr
            Case attrConeAngle:
                Set oConeAngle = oTempAttrDescr
            Case attrConeLength:
                Set oConeLength = oTempAttrDescr
            Case attrCone1Slope:
                Set oCone1Slope = oTempAttrDescr
            Case attrCone1Angle:
                Set oCone1Angle = oTempAttrDescr
            Case attrCone1Length:
                Set oCone1Length = oTempAttrDescr
            Case attrCone2Slope:
                Set oCone2Slope = oTempAttrDescr
            Case attrCone2Angle:
                Set oCone2Angle = oTempAttrDescr
            Case attrCone2Length:
                Set oCone2Length = oTempAttrDescr
        End Select
    Next oTempAttrDescr
    
    Select Case oAttrDescr.attrName
        Case attrDiameterRule:
            RuleSetReadOnly varNewAttrValue, oCanID, oCanOD
        Case attrConeLengthMethod:
            ConeMethodSetReadOnly varNewAttrValue, oConeLength, oConeSlope, oConeAngle
        Case attrCone1LengthMethod:
            ConeMethodSetReadOnly varNewAttrValue, oCone1Length, oCone1Slope, oCone1Angle
        Case attrCone2LengthMethod:
            ConeMethodSetReadOnly varNewAttrValue, oCone2Length, oCone2Slope, oCone2Angle
        Case attrL2CompMethod:
            LMethodSetReadOnly varNewAttrValue, oL2Length, oL2Factor
        Case attrL3CompMethod:
            LMethodSetReadOnly varNewAttrValue, oL3Length, oL3Factor
    End Select
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Function OnAttributeChange(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
    Const METHOD = "OnAttributeChange"
    On Error GoTo ErrorHandler
    
    'Validate attribute value
    OnAttributeChange = ValidateAttribute(pIJDAttrs, CollAllDisplayedValues, pAttrToChange, varNewAttrValue)
    If Len(OnAttributeChange) > 0 Then
        Exit Function
    End If
    
    ' material, grade, and thickness are all from drop down list now
    
    ' update values
    If pAttrToChange.attrName = attrCanThickness _
    Or pAttrToChange.attrName = attrOuterDiameter Then
        OnAttributeChange = UpdateAttributeValues(pIJDAttrs, CollAllDisplayedValues, pAttrToChange, varNewAttrValue)
        If Len(OnAttributeChange) > 0 Then
            Exit Function
        End If
    End If
    
    ' update readonly flags
    OnAttributeChange = UpdateAttributeReadOnly(pIJDAttrs, CollAllDisplayedValues, pAttrToChange, varNewAttrValue)
    Exit Function

ErrorHandler:
    OnAttributeChange = METHOD & ": " & Err.Description
End Function


