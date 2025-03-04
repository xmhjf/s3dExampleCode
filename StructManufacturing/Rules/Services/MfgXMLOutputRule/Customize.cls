VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''***********************************************************************************
''  Copyright (C) 2013, Intergraph Corporation.  All rights reserved.
''
''  Project     : StrMfgXMLOutput
''  File        : Customize.cls
''
''  Description :
''
''  Author      : KKamph
''
''  History     :
''                18 April 13       KKamph      - Creation
''                19 August 13      BAffenzeller    - CR-CP-228267
''***********************************************************************************
Option Explicit

Implements IJDMfgOutputCustomRule
Private SymPath As String

Private Const MODULE = "Customize"
Private Const dDistTolerance = 1
Private Const dAngleTolerance = 1
Private Const TKApplicationContext = "ApplicationContext"
Private Const strAssemblyPathDelimiter = "\"
Private Const IID_IJDProductionRouting As String = "{B447C9B4-FB74-11D1-8A49-00A0C9065DF6}"
Private Const IID_IJPlnProductionRouting As String = "{E9B1D9F6-C687-40BA-85E2-311534F4782A}"
Private Const IID_IJMfgOuptput As String = "{08E0F9B1-41F0-4279-8EBA-AE37DD7389CD}"
Private Const strProductionRoutingColName As String = "ProductionRouting"
Private Const strRoutingActionColName As String = "RoutingAction"
Private Const strNestingDataColName As String = "MfgPartNestData_ORIG"

Private m_strFormat As String
Private m_oControlXML As DOMDocument
Private m_oXML As IXMLDOMElement

Private Sub IJDMfgOutputCustomRule_CustomAnnotation(ByVal oDispMfgPart As Object, ByVal oMfgGeom2d As GSCADMfgRulesDefinitions.IJMfgGeom2d, ByVal bstrFormat As String, ByVal oGeom2DNode As Object)
    Const METHOD = "IJDMfgOutputCustomRule_CustomPartInfo"
    On Error GoTo ErrorHandler

Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IJDMfgOutputCustomRule_CustomBevelInfo(ByVal oDispMfgPart As Object, ByVal oMfgGeom2d As GSCADMfgRulesDefinitions.IJMfgGeom2d, ByVal bstrFormat As String, ByVal oGeom2DNode As Object)
    Const METHOD = "IJDMfgOutputCustomRule_CustomPartInfo"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IJDMfgOutputCustomRule_CustomGeometryInfo(ByVal oDispMfgPart As Object, ByVal oMfgGeom2d As GSCADMfgRulesDefinitions.IJMfgGeom2d, ByVal bstrFormat As String, ByVal oGeom2DNode As Object)
    Const METHOD = "IJDMfgOutputCustomRule_CustomPartInfo"
    On Error GoTo ErrorHandler

Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IJDMfgOutputCustomRule_CustomLabel(ByVal oDispMfgPart As Object, ByVal bstrXSectionType As String, ByVal oLabelNode As Object)
    Const METHOD = "IJDMfgOutputCustomRule_CustomPartInfo"
    On Error GoTo ErrorHandler
        
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IJDMfgOutputCustomRule_CustomPartInfo(ByVal oDispMfgPart As Object, ByVal bstrFormat As String, ByVal oDoc As Object)
    Const METHOD = "IJDMfgOutputCustomRule_CustomPartInfo"
    On Error GoTo ErrorHandler
    
    Dim oXMLElement As IXMLDOMElement
    Set oXMLElement = oDoc
    Set m_oXML = oXMLElement 'SMS_PROD_INFO Node
    
    Dim oControlXML As DOMDocument
    Set oControlXML = GetControlXML(bstrFormat)
    Dim bCreateProductionAction As Boolean: bCreateProductionAction = True
    
    If Not oControlXML Is Nothing Then
        Dim oPropertyNode As IXMLDOMNode
        'SMS_PROD_INFO||SMS_PART_INFO||SMS_PROFILE_SECTION_INFO is not implemented
        For Each oPropertyNode In oControlXML.selectNodes("//SMS_OUTPUT[@TYPE ='*']/PROPERTY")
            Dim strAttributeValue As String
            Dim strAttributeValues() As String
            Dim strAttributeName As String
            strAttributeName = oPropertyNode.selectSingleNode("./@NAME").nodeValue
            Dim strActionName As String
            strActionName = oPropertyNode.selectSingleNode("./@ACTION").nodeValue
            
            Dim strProcessName As String
            If Not oPropertyNode.selectSingleNode("./@PROCESS") Is Nothing Then
            strProcessName = oPropertyNode.selectSingleNode("./@PROCESS").nodeValue
            End If
            
            Dim strFormat As String
            If Not oPropertyNode.selectSingleNode("./@FORMAT") Is Nothing Then
            strFormat = oPropertyNode.selectSingleNode("./@FORMAT").nodeValue
            End If
            
            Dim strQuery As String
            Dim strTemp() As String
            strTemp = Split(strAttributeName, "||")
            Dim lCount As Long
            
            Dim strXPathQuery As String
            If strActionName = "SMS_SMS_ATTRIBUTE" Then
                strXPathQuery = "//"
            Else
                strXPathQuery = "./"
            End If
            
            Dim strNodeQuery As String
            If oXMLElement.nodeName = strTemp(LBound(strTemp)) Then
                strNodeQuery = "."
            Else
                strNodeQuery = strXPathQuery + strTemp(LBound(strTemp))
            End If
            
            For lCount = LBound(strTemp) + 1 To UBound(strTemp) - 1
                strXPathQuery = strXPathQuery + strTemp(lCount)
                strXPathQuery = strXPathQuery + "/"
            Next lCount
            
            Dim oNodeList As IXMLDOMNodeList
            Set oNodeList = oXMLElement.selectNodes(strNodeQuery)
            
            Dim nNodeCount As Integer
            For nNodeCount = 0 To oNodeList.Length - 1

            
                Dim oNodeElement As IXMLDOMElement
                Set oNodeElement = oNodeList.Item(nNodeCount)

            
                Dim oObject As Object
                'Dim strPartAttributeValue As String
                
                If oNodeElement.nodeName = "SMS_PART_INFO" Then
                    Dim strPartGuid As String: strPartGuid = "" 'MfgPart
                    Dim strModelPartGuid As String: strModelPartGuid = "" 'DetailObject
                    Dim oPartGuidNode As IXMLDOMNode
                    Dim oModelPartGuidNode As IXMLDOMNode
                    
                    On Error Resume Next
                    Set oPartGuidNode = oNodeElement.selectSingleNode("./@PART_GUID")
                    Set oModelPartGuidNode = oNodeElement.selectSingleNode("./@MODEL_PART_GUID")

                    On Error GoTo ErrorHandler
                    If Not oPartGuidNode Is Nothing Then
                        strPartGuid = oPartGuidNode.nodeValue
                    End If
                    If Not oModelPartGuidNode Is Nothing Then
                        strModelPartGuid = oModelPartGuidNode.nodeValue
                    End If
                    
                    If strPartGuid = CleanGUID(GetOID(oDispMfgPart)) Then
                        Set oObject = oDispMfgPart
                    ElseIf strPartGuid <> "" Then
                        Set oObject = GetObjectFromOid("{" + strPartGuid + "}")
                    ElseIf strModelPartGuid <> "" Then
                        Set oObject = GetObjectFromOid("{" + strModelPartGuid + "}")
                    End If
                Else
                    'Input oDispMfgPart is default
                    Set oObject = oDispMfgPart
                End If
                
                Dim strTempAttributeName As String: strTempAttributeName = ""
            Dim strTempXPathQuery As String: strTempXPathQuery = ""

            Dim nCount As Long
            If strActionName = "PRODUCTION_ROUTING" Then
            
                'WorkCenter
                    strTempXPathQuery = strXPathQuery + strTemp(UBound(strTemp)) + "/@" + "WORKCENTER"
                                        
                    For lCount = LBound(strTemp) + 1 To UBound(strTemp)
                        strTempAttributeName = strTempAttributeName + strTemp(lCount) + "||"
                    Next lCount
                    strTempAttributeName = strTempAttributeName + "WORKCENTER"
                    
                strFormat = "IJPlnProductionRouting::WorkCenter"
                
                    strAttributeValue = GetAttributeValue(oObject, oNodeElement, strTempAttributeName, strActionName, strProcessName, strFormat)
                    CreateNode oNodeElement, strTempAttributeName, False
                    SetNodeValue oNodeElement, strTempXPathQuery, strAttributeValue, strActionName

                'StageCode
                    strTempXPathQuery = strXPathQuery + strTemp(UBound(strTemp)) + "/@" + "STAGE_CODE"
                    For lCount = LBound(strTemp) + 1 To UBound(strTemp)
                        strTempAttributeName = strTempAttributeName + strTemp(lCount) + "||"
                    Next lCount
                    strTempAttributeName = strTempAttributeName + "STAGE_CODE"
                strFormat = "IJPlnProductionRouting::StageCode"
                 
                    strAttributeValue = GetAttributeValue(oObject, oNodeElement, strTempAttributeName, strActionName, strProcessName, strFormat)
                    SetNodeValue oNodeElement, strTempXPathQuery, strAttributeValue, strActionName
    
                    
            ElseIf strActionName = "PRODUCTION_ROUTING_ACTION" Then
               
                'TYPE
                    strTempXPathQuery = strXPathQuery + strTemp(UBound(strTemp)) + "/@" + "TYPE"
                    For lCount = LBound(strTemp) + 1 To UBound(strTemp)
                        strTempAttributeName = strTempAttributeName + strTemp(lCount) + "||"
                    Next lCount
                    strTempAttributeName = strTempAttributeName + "TYPE"
                strFormat = "IJPlnRoutingAction::Action"
                
                    strAttributeValues = GetAttributeValues(oObject, strTempAttributeName, strActionName, strProcessName, strFormat)
                If UBound(strAttributeValues) > 0 Then
                    For nCount = LBound(strAttributeValues) To UBound(strAttributeValues)
                            CreateNode oNodeElement, strTempAttributeName, True
                    Next nCount
                End If
                
                    SetNodeValues oNodeElement, strTempXPathQuery, strAttributeValues
                Erase strAttributeValues
                
                'CODE
                    strTempXPathQuery = strXPathQuery + strTemp(UBound(strTemp)) + "/@" + "CODE"
                    For lCount = LBound(strTemp) + 1 To UBound(strTemp)
                        strTempAttributeName = strTempAttributeName + strTemp(lCount) + "||"
                    Next lCount
                    strTempAttributeName = strTempAttributeName + "CODE"
                strFormat = "IJPlnRoutingAction::Code"
                
                    strAttributeValues = GetAttributeValues(oObject, strTempAttributeName, strActionName, strProcessName, strFormat)
                    SetNodeValues oNodeElement, strTempXPathQuery, strAttributeValues
                Erase strAttributeValues
                
                'MACHINE
                    strTempXPathQuery = strXPathQuery + strTemp(UBound(strTemp)) + "/@" + "MACHINE"
                    For lCount = LBound(strTemp) + 1 To UBound(strTemp)
                        strTempAttributeName = strTempAttributeName + strTemp(lCount) + "||"
                    Next lCount
                    strTempAttributeName = strTempAttributeName + "MACHINE"
                strFormat = "IJPlnRoutingAction::Machine"

                    strAttributeValues = GetAttributeValues(oObject, strTempAttributeName, strActionName, strProcessName, strFormat)
                    SetNodeValues oNodeElement, strTempXPathQuery, strAttributeValues
                Erase strAttributeValues
                
                'SEQUENCE
                    strTempXPathQuery = strXPathQuery + strTemp(UBound(strTemp)) + "/@" + "SEQUENCE"
                    For lCount = LBound(strTemp) + 1 To UBound(strTemp)
                        strTempAttributeName = strTempAttributeName + strTemp(lCount) + "||"
                    Next lCount
                    strTempAttributeName = strTempAttributeName + "SEQUENCE"
                strFormat = ""
                
                    strAttributeValues = GetAttributeValues(oObject, strTempAttributeName, strActionName, strProcessName, strFormat)
                    SetNodeValues oNodeElement, strTempXPathQuery, strAttributeValues
                Erase strAttributeValues
                
                    
            ElseIf strActionName = "ROUTING_ACTION_TYPE" Or strActionName = "ROUTING_ACTION_CODE" Then
                    strTempXPathQuery = strXPathQuery + "@" + strTemp(UBound(strTemp))
                    For lCount = LBound(strTemp) + 1 To UBound(strTemp) - 1
                        strTempAttributeName = strTempAttributeName + strTemp(lCount) + "||"
                    Next lCount
                    strTempAttributeName = strTempAttributeName + strTemp(UBound(strTemp))
                    
                    strAttributeValues = GetAttributeValues(oObject, strTempAttributeName, strActionName, strProcessName, strFormat)
                If UBound(strAttributeValues) > 0 Then
                    For nCount = LBound(strAttributeValues) To UBound(strAttributeValues)
                            CreateNode oNodeElement, strTempAttributeName, bCreateProductionAction
                    Next nCount
                    bCreateProductionAction = False
                End If
                    SetNodeValues oNodeElement, strTempXPathQuery, strAttributeValues
                Erase strAttributeValues
                
            ElseIf strActionName = "NESTING_DATA" Then
                
                    strTempXPathQuery = strXPathQuery + "@" + strTemp(UBound(strTemp))

                    For lCount = LBound(strTemp) + 1 To UBound(strTemp) - 1
                        strTempAttributeName = strTempAttributeName + strTemp(lCount) + "||"
                    Next lCount
                    strTempAttributeName = strTempAttributeName + strTemp(UBound(strTemp))

                    strAttributeValue = GetAttributeValue(oObject, oNodeElement, strTempAttributeName, strActionName, strProcessName, strFormat)
                    CreateNode oNodeElement, strTempAttributeName, False
                    SetNodeValue oNodeElement, strTempXPathQuery, strAttributeValue, strActionName
                    
                ElseIf strActionName = "SMS_SMS_ATTRIBUTE" Then
                    strTempXPathQuery = "./" + "@" + strTemp(UBound(strTemp))
                    strAttributeValue = GetAttributeValue(oObject, oNodeElement, strAttributeName, strActionName, strProcessName, strFormat)
                    SetNodeValue oNodeElement, strTempXPathQuery, strAttributeValue, strActionName
                    
                Else
                    strTempXPathQuery = strXPathQuery + "@" + strTemp(UBound(strTemp))
                    strAttributeValue = GetAttributeValue(oObject, oNodeElement, strAttributeName, strActionName, strProcessName, strFormat)
                    SetNodeValue oNodeElement, strTempXPathQuery, strAttributeValue, strActionName
            End If
            

                Set oObject = Nothing
            Next nNodeCount
            
            Erase strTemp

        Next oPropertyNode
    End If
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function GetAttributeValue(oDispMfgPart As Object, ByVal oXMLElement As IXMLDOMElement, strAttributeName As String, strAction As String, strProcess As String, strFormat As String) As String
    Const METHOD = "GetAttributeValue"
    On Error GoTo ErrorHandler
    Dim strBlockName As String
    Dim strBlockBoardSide As String
    Dim strAssemblyName As String
    
    If strAttributeName = "SMS_PROD_INFO||BLOCK_BOARDSIDE" Or _
       strAttributeName = "SMS_PROD_INFO||BLOCK_NAME" Or _
       strAttributeName = "SMS_PROD_INFO||ASSEMBLY_NAME" Then
            GetAttributeValue = GetBlockInfo(oDispMfgPart, strAttributeName, strAction, strProcess, strFormat, strBlockName, strAssemblyName, strBlockBoardSide)
    ElseIf strAttributeName = "SMS_PART_INFO||MIN_BOX_LENGTH" Or _
           strAttributeName = "SMS_PART_INFO||MIN_BOX_WIDTH" Then
                GetAttributeValue = GetMinBoundingBoxInfo(oDispMfgPart, strAttributeName, strAction, strProcess, strFormat, strBlockName, strAssemblyName, strBlockBoardSide)
    ElseIf strAttributeName = "SMS_PART_INFO||SEQUENCE" Then
        GetAttributeValue = GetSequence(oDispMfgPart)
    ElseIf strAttributeName = "SMS_COMMON_PART_INFO||COMMON_PART_NAME" Then
            GetAttributeValue = GetPartModelName(oDispMfgPart, strAttributeName, strAction, strProcess, strFormat, strBlockName, strAssemblyName, strBlockBoardSide)
'    ElseIf strAttributeName = "SMS_PROD_INFO||ASSEMBLY_LEVEL" Then
'        GetAttributeValue = GetAssemblyLevel(oDispMfgPart)
    ElseIf InStr(strAttributeName, "ASSEMBLY") > 0 Or InStr(strAttributeName, "BLOCK") > 0 Then
        GetAttributeValue = GetGenericAssemblyAttribute(oDispMfgPart, strProcess, strFormat) '"Not Implemented"
    ElseIf strAction = "PRODUCTION_ROUTING" Then
        GetAttributeValue = GetProductionRoutingInfo(oDispMfgPart, strProcess, strFormat)
    ElseIf strAction = "NESTING_DATA" Then
        GetAttributeValue = GetNestingInfo(oDispMfgPart, oXMLElement, strProcess, strFormat)
    ElseIf strAction = "SMS_SMS_ATTRIBUTE" Then
        GetAttributeValue = GetProjectInfo(oDispMfgPart, strProcess, strFormat)
    Else
        GetAttributeValue = "Not Implemented"
    End If
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetAttributeValues(oDispMfgPart As Object, strAttributeName As String, strAction As String, strProcess As String, strFormat As String) As String()
    Const METHOD = "GetAttributeValues"
    On Error GoTo ErrorHandler
    Dim strBlockName As String
    Dim strBlockBoardSide As String
    Dim strAssemblyName As String
    
    If strAction = "PRODUCTION_ROUTING_ACTION" Or strAction = "ROUTING_ACTION_TYPE" Or strAction = "ROUTING_ACTION_CODE" Then
        GetAttributeValues = GetProductionRoutingActionInfos(oDispMfgPart, strAttributeName, strAction, strProcess, strFormat)
    Else
        Dim strEmptyValues() As String
        ReDim strEmptyValues(0 To 0)

        GetAttributeValues = strEmptyValues
        Erase strEmptyValues
        
    End If
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetControlXML(strFormat As String) As DOMDocument
    Const METHOD = "GetControlXML"
    On Error GoTo ErrorHandler
    
    If (strFormat = m_strFormat And Not m_oControlXML Is Nothing) Then
        Set GetControlXML = m_oControlXML
        Exit Function
    Else
        Dim oOutputHelper As IJMfgOutputHelper
        Set oOutputHelper = New MfgCatalogQueryHelper
        strFormat = oOutputHelper.GetOutputRuleValue(strFormat, "CustomProcess")
        m_strFormat = strFormat
        Dim lPipeLoc As String
        lPipeLoc = InStr(strFormat, "|")
        If (lPipeLoc > 0) Then
            Dim strControlXML As String
            strControlXML = GetSymbolSharePath() + "\" + Right(strFormat, Len(strFormat) - lPipeLoc)
            Set m_oControlXML = New DOMDocument
            m_oControlXML.Load strControlXML
            Set GetControlXML = m_oControlXML
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetGenericAssemblyAttribute(dispMfgPart As Object, strProcess As String, strFormat As String) As String
    Const METHOD = "GetGenericAssemblyAttribute"
    On Error GoTo ErrorHandler
    
    Dim strInterface As String
    Dim strAttribute As String
    strInterface = Left(strFormat, InStr(strFormat, "::") - 1)
    strAttribute = Right(strFormat, Len(strFormat) - Len(strInterface) - 2)
    Dim oAssemblyChild As IJAssemblyChild
    
    Dim bAssemblyName As Boolean
    bAssemblyName = False
    Dim oStrMfgEntityHelper As IJMfgEntityHelper
    Set oStrMfgEntityHelper = New MfgEntityHelper
    GetGenericAssemblyAttribute = "Unknown"       'Default value

    If TypeOf dispMfgPart Is IJMfgPlatePart Then
        Dim pPlatePart As IJPlatePart
        dispMfgPart.GetDetailedPart pPlatePart
        Set oAssemblyChild = pPlatePart
    ElseIf TypeOf dispMfgPart Is IJMfgProfilePart Then
        Dim oProfilePart As Object
        Set oProfilePart = dispMfgPart.GetDetailedPart
        Set oAssemblyChild = oProfilePart
    Else
        Set oAssemblyChild = dispMfgPart
    End If
    
    Dim oPartAsAssemblyChild As IJAssemblyChild
    Set oPartAsAssemblyChild = oAssemblyChild

    Do While (Not oAssemblyChild Is Nothing)
        Dim oParent As Object
        Set oParent = oAssemblyChild.Parent
        
        If (Not oParent Is Nothing) Then
            Dim oConfigRoot As IJConfigProjectRoot
            'Set oConfigRoot = oParent
            If (Not TypeOf oParent Is IJConfigProjectRoot) Then
                Dim oAssemblyBase  As IJAssemblyBase
                Set oAssemblyBase = oParent
                Dim lAssemblyType As Long
                lAssemblyType = oAssemblyBase.Type
                
                Dim strTest As String
                strTest = TypeName(oAssemblyBase)
                If IsStringInString(strProcess, "" & lAssemblyType) Then
                    Dim oAttribute As IJDAttribute
                    
                    'Set oAttribute = oStrMfgEntityHelper.GetAttributeFromVirtualInterface(oParent, strInterface, strAttribute)
                    If Not oAttribute Is Nothing Then
                        GetGenericAssemblyAttribute = oAttribute.Value
                    Else
                        Dim oMetaDataHelp      As IJDAttributeMetaData
                        Dim oAttributes        As IJDAttributes
                        Dim oAttrIID           As Variant
                        Dim oIJDAttributesCol  As IJDAttributesCol
                        Dim oIJDInterfaceInfo  As IJDInterfaceInfo
                        Set oAttributes = oParent
                        Set oMetaDataHelp = oParent
        
                        For Each oAttrIID In oAttributes
                            If oAttrIID = oMetaDataHelp.IID(strInterface) Then
                                Set oIJDInterfaceInfo = oMetaDataHelp.InterfaceInfo(oAttrIID)
                                Set oIJDAttributesCol = oAttributes.CollectionOfAttributes(oIJDInterfaceInfo.Type)
                                
                                For Each oAttribute In oIJDAttributesCol
                                    If oAttribute.AttributeInfo.Name = strAttribute Then
                                        If oAttribute.AttributeInfo.CodeListTableName = "" Then
                                            GetGenericAssemblyAttribute = oAttribute.Value
                                        Else
                                            GetGenericAssemblyAttribute = GetShortDescription(oParent, oAttribute.AttributeInfo.CodeListTableName, oAttribute.Value)
                                            Exit Function
                                        End If
                                    End If
                                Next
                                
                            End If
                        Next
                    End If
                    GoTo CleanUp 'Exit Function
                End If
            Else
                GoTo CleanUp
            End If
        End If
        Set oAssemblyChild = oParent
    Loop
ErrorHandler:
CleanUp:
    Set oStrMfgEntityHelper = Nothing
    Set oAttribute = Nothing
    Set oAssemblyChild = Nothing
    Set pPlatePart = Nothing
    Set oProfilePart = Nothing
    Set oPartAsAssemblyChild = Nothing
    Set oParent = Nothing
    Set oConfigRoot = Nothing
    
    Exit Function
End Function

Private Function GetShortDescription(oObject As Object, strTableName As String, valueID As Long) As String
    On Error GoTo ErrorHandler
    Const METHOD = "GetShortDescription"

    Dim oIJDCodeListMetaData As IJDCodeListMetaData
    Set oIJDCodeListMetaData = oObject

    If Not oIJDCodeListMetaData Is Nothing Then
        GetShortDescription = oIJDCodeListMetaData.ShortStringValue(strTableName, valueID)
    End If

    Exit Function
ErrorHandler:
    Resume Next

End Function

Private Function GetBlockInfo(dispMfgPart As Object, strAttributeName As String, strAssemblyNameType As String, strProcess As String, strFormat As String, strBlockName As String, strAssemblyName As String, strBlockBoardSide As String) As String
    Const METHOD = "GetBlockName"
    On Error GoTo ErrorHandler
    
    Dim oAssemblyChild As IJAssemblyChild
    
    Dim bAssemblyName As Boolean
    bAssemblyName = False
    
    strBlockName = "null"       'Default value
    strBlockBoardSide = "C"     'Default value
    strAssemblyName = "null"    'Default value

    If TypeOf dispMfgPart Is IJMfgPlatePart Then
        Dim pPlatePart As IJPlatePart
        dispMfgPart.GetDetailedPart pPlatePart
        Set oAssemblyChild = pPlatePart
    ElseIf TypeOf dispMfgPart Is IJMfgProfilePart Then
        Dim oProfilePart As Object
        Set oProfilePart = dispMfgPart.GetDetailedPart
        Set oAssemblyChild = oProfilePart
    Else
        Exit Function
    End If
    
    Dim oPartAsAssemblyChild As IJAssemblyChild
    Set oPartAsAssemblyChild = oAssemblyChild

    Do While (Not oAssemblyChild Is Nothing)
        Dim oParent As Object
        Set oParent = oAssemblyChild.Parent
        
        If (Not oParent Is Nothing) Then
            Dim oConfigRoot As IJConfigProjectRoot
            'Set oConfigRoot = oParent
            If (Not TypeOf oParent Is IJConfigProjectRoot) Then
                Dim oAssemblyBase  As IJAssemblyBase
                Set oAssemblyBase = oParent
                Dim lAssemblyType As Long
                lAssemblyType = oAssemblyBase.Type
                Dim strTest As String
                strTest = TypeName(oAssemblyBase)
                If ((IsStringInString(strProcess, "" & lAssemblyType) And strAssemblyNameType = "ASSEMBLY_TYPE") Or _
                    (TypeOf oAssemblyBase Is IJBlock And strAssemblyNameType = "INTERFACE" And IsStringInString(strProcess, "IJBlock")) Or _
                    (TypeOf oAssemblyBase Is IJAssembly And strAssemblyNameType = "INTERFACE" And IsStringInString(strProcess, "IJAssembly")) And Not bAssemblyName) Then
                    
                    strAssemblyName = GetFormattedAssemblyType(oPartAsAssemblyChild, oParent, strAttributeName, strAssemblyNameType, strProcess, strFormat, strBlockName, strAssemblyName, strBlockBoardSide)

                    bAssemblyName = True
                End If
                
                If ((IsStringInString(strProcess, "" & lAssemblyType) And strAssemblyNameType = "ASSEMBLY_TYPE") Or _
                    (TypeOf oAssemblyBase Is IJBlock And strAssemblyNameType = "INTERFACE" And IsStringInString(strProcess, "IJBlock")) Or _
                    (TypeOf oAssemblyBase Is IJAssembly And strAssemblyNameType = "INTERFACE" And IsStringInString(strProcess, "IJAssembly"))) Then

                    strBlockName = GetFormattedAssemblyType(oPartAsAssemblyChild, oParent, strAttributeName, strAssemblyNameType, strProcess, strFormat, strBlockName, strAssemblyName, strBlockBoardSide)

                    Dim lBoardSide As Long
                    lBoardSide = oAssemblyBase.BoardSide
    
                    strBlockBoardSide = "C"
                    If (lBoardSide = 1) Then strBlockBoardSide = "C"
                    If (lBoardSide = 2) Then strBlockBoardSide = "S"
                    If (lBoardSide = 3) Then strBlockBoardSide = "P"
                    If (lBoardSide = 4) Then strBlockBoardSide = "B"
                
                    If strAttributeName = "SMS_PROD_INFO||BLOCK_NAME" Then GetBlockInfo = strBlockName
                    If strAttributeName = "SMS_PROD_INFO||BLOCK_BOARDSIDE" Then GetBlockInfo = strBlockBoardSide
                    If strAttributeName = "SMS_PROD_INFO||ASSEMBLY_NAME" Then GetBlockInfo = strAssemblyName
                    Exit Function
                End If
                Set oAssemblyChild = Nothing
                Set oAssemblyChild = oParent
            Else
                GoTo CleanUp
            End If
        End If
        Set oAssemblyChild = oParent
    Loop
ErrorHandler:
CleanUp:
    If strAttributeName = "SMS_PROD_INFO||BLOCK_NAME" Then GetBlockInfo = strBlockName
    If strAttributeName = "SMS_PROD_INFO||BLOCK_BOARDSIDE" Then GetBlockInfo = strBlockBoardSide
    If strAttributeName = "SMS_PROD_INFO||ASSEMBLY_NAME" Then GetBlockInfo = strAssemblyName
    Exit Function
End Function

Private Function GetMinBoundingBoxInfo(dispMfgPart As Object, strAttributeName As String, strAssemblyNameType As String, strProcess As String, strFormat As String, strBlockName As String, strAssemblyName As String, strBlockBoardSide As String) As String
    Const METHOD = "GetMinBoundingBoxInfo"
    On Error GoTo ErrorHandler

    Dim dLength As Double
    Dim dWidth As Double
           
    GetMinBoundingBoxInfo = TruncateDouble(0#, 3, vbTrue, vbFalse, vbFalse)

    If TypeOf dispMfgPart Is IJMfgPlatePart Then
    
        Dim oMfgPlatePart As IJMfgPlatePart
        Set oMfgPlatePart = dispMfgPart
        oMfgPlatePart.BestFitDimensions dLength, dWidth
        
    ElseIf TypeOf dispMfgPart Is IJMfgProfilePart Then
        
        Dim strCellName As String
        Dim oCellNameNode As IXMLDOMNode
        Set oCellNameNode = m_oXML.selectSingleNode("./SMS_PART_INFO/@CELL_NAME")
       
        If Not oCellNameNode Is Nothing Then
            strCellName = oCellNameNode.nodeTypedValue
            
            If strCellName = "SEC" Then
                strCellName = "TOP_FLANGE"
            End If
        
            Dim oMfgProfilePart As IJMfgProfilePart
            Set oMfgProfilePart = dispMfgPart
            
            oMfgProfilePart.BestFitDimensions strCellName, dLength, dWidth
        End If
        
    Else
        Exit Function
    End If
    
    If strAttributeName = "SMS_PART_INFO||MIN_BOX_LENGTH" Then
        GetMinBoundingBoxInfo = TruncateDouble(dLength * 1000, 3, vbTrue, vbFalse, vbFalse)
    ElseIf strAttributeName = "SMS_PART_INFO||MIN_BOX_WIDTH" Then
        GetMinBoundingBoxInfo = TruncateDouble(dWidth * 1000, 3, vbTrue, vbFalse, vbFalse)
    End If
    
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetSequence(dispMfgPart As Object) As String
Const METHOD = "GetSequence"
On Error GoTo ErrorHandler

Dim strSequence As String: strSequence = ""
Dim oCommonParent As IJAssemblySequence
Dim oMfgPlatePart As IJMfgPlatePart
Dim oDetailedPart As IJAssemblyChild

Dim oParent As IJAssemblySequence
Dim oPanelPart As IJMfgPlatePart
Dim oPanelNode As IXMLDOMNode
Dim strPanelGuid As String

Set oPanelNode = m_oXML.selectSingleNode("./SMS_PART_INFO[@PART_TYPE='PANEL']/@PART_GUID")
If Not oPanelNode Is Nothing Then
    strPanelGuid = oPanelNode.nodeTypedValue

    Set oPanelPart = GetObjectFromOid("{" + strPanelGuid + "}")
    
    Dim oPlnIntHelper As IJDPlnIntHelper
    Dim oPanelColl As IJElements
    
    Set oPanelColl = oPanelPart.PanelPartCollection
    Set oPlnIntHelper = New CPlnIntHelper
    Set oCommonParent = oPlnIntHelper.FindCommonParent(oPanelColl, "IJAssembly")
    
    If TypeOf dispMfgPart Is IJMfgPlatePart Then
        Set oMfgPlatePart = dispMfgPart
        
        If Not oMfgPlatePart.PanelMode Then
            oMfgPlatePart.GetDetailedPart oDetailedPart
            strSequence = GetPanelPlateSequence(oDetailedPart, oCommonParent)
        End If
        
    ElseIf TypeOf dispMfgPart Is IJPlatePart Then
        strSequence = GetPanelPlateSequence(dispMfgPart, oCommonParent)
    End If
    
ElseIf TypeOf dispMfgPart Is IJMfgProfilePart Then
    Dim oMfgProfilePart As IJMfgProfilePart
    Set oMfgProfilePart = dispMfgPart
    Set oDetailedPart = oMfgProfilePart.GetDetailedPart
    Set oParent = oDetailedPart.Parent

    If (Not oParent Is Nothing) Then
        strSequence = CStr(oParent.ChildIndex(oDetailedPart))
    End If
    
ElseIf TypeOf dispMfgPart Is IJMfgPlatePart Then
    Set oMfgPlatePart = dispMfgPart
    oMfgPlatePart.GetDetailedPart oDetailedPart
    Set oParent = oDetailedPart.Parent
    
    If (Not oParent Is Nothing) Then
        strSequence = CStr(oParent.ChildIndex(oDetailedPart))
    End If
End If

If strSequence = "-1" Then
    strSequence = ""
End If

GetSequence = strSequence

CleanUp:
    oCommonParent = Nothing
    oMfgPlatePart = Nothing
    oDetailedPart = Nothing
    oParent = Nothing
    oPanelPart = Nothing
    oPanelNode = Nothing
    Exit Function
ErrorHandler:
End Function

Private Function GetPanelPlateSequence(oDetailedPlate As Object, oCommonParent As IJAssemblySequence) As String
Const METHOD = "GetPanelPlateSequence"
On Error GoTo ErrorHandler

Dim strSequence As String: strSequence = ""
Dim oDetailedPart As IJAssemblyChild
Dim oParent As IJAssemblySequence
Dim IsSequenceNull As Boolean: IsSequenceNull = True
Dim oAssemblyClimber As IJAssemblyChild
Dim strSubSequence As String

Set oDetailedPart = oDetailedPlate
Set oParent = oDetailedPart.Parent

Do While Not oCommonParent Is Nothing
    strSubSequence = CStr(oParent.ChildIndex(oDetailedPart))
    If strSubSequence = "-1" Then
        strSubSequence = "0"
    Else
        IsSequenceNull = False
    End If

    If Not strSequence = "" Then
        strSequence = strSubSequence & " - " & strSequence
    Else
        strSequence = strSubSequence
    End If

    If oParent Is oCommonParent Then
        Exit Do
    End If

    Set oAssemblyClimber = oParent
    Set oParent = oAssemblyClimber.Parent

    Set oDetailedPart = oDetailedPart.Parent
Loop

If IsSequenceNull = True Then
    strSequence = ""
End If

GetPanelPlateSequence = strSequence

CleanUp:
    oDetailedPart = Nothing
    oParent = Nothing
    oAssemblyClimber = Nothing
    Exit Function
ErrorHandler:
End Function
Private Function GetFormattedAssemblyType(oPart As Object, oAssemblyChild As Object, strAttributeName As String, strAssemblyNameType As String, strProcess As String, strFormat As String, strBlockName As String, strAssemblyName As String, strBlockBoardSide As String) As String
    Const METHOD = "GetFormattedAssemblyType"
    On Error GoTo ErrorHandler
    
    Dim oPartAsAssemblyChild As IJAssemblyChild
    Set oPartAsAssemblyChild = oPart
    
    Dim strFormattedAssemblyType As String
    strFormattedAssemblyType = "null"

    Dim oPlnIntHelper As IJDPlnIntHelper
    Set oPlnIntHelper = New CPlnIntHelper
    
    If Not strFormat <> "FULL_PATH" And Not TypeOf oPartAsAssemblyChild.Parent Is IJConfigProjectRoot Then
    
        strFormattedAssemblyType = oPlnIntHelper.GetAssemblyPath(oPartAsAssemblyChild.Parent, strAssemblyPathDelimiter)
            
    ElseIf Not strFormat <> "ROOT_PATH" And TypeOf oAssemblyChild Is IJAssemblyBase Then
    
        strFormattedAssemblyType = oPlnIntHelper.GetAssemblyPath(oAssemblyChild, strAssemblyPathDelimiter)
        
    ElseIf Not strFormat <> "SUB_PATH" And _
           Not TypeOf oPartAsAssemblyChild.Parent Is IJConfigProjectRoot And _
           TypeOf oAssemblyChild Is IJAssemblyBase Then
        
        Dim strFullPath As String
        strFullPath = oPlnIntHelper.GetAssemblyPath(oPartAsAssemblyChild.Parent, strAssemblyPathDelimiter)

        Dim strRootPath As String
        strRootPath = oPlnIntHelper.GetAssemblyPath(oAssemblyChild, strAssemblyPathDelimiter)
        strRootPath = strRootPath + strAssemblyPathDelimiter
        
        Dim strAssemblyChildName As String
        strAssemblyChildName = Mid(strFullPath, InStr(strFullPath, strRootPath) + Len(strRootPath))
    
        Dim oNamedItem As IJNamedItem
        Set oNamedItem = oAssemblyChild
        If strAssemblyChildName <> "" Then
            strAssemblyChildName = oNamedItem.Name + strAssemblyPathDelimiter + strAssemblyChildName
        Else
            strAssemblyChildName = oNamedItem.Name
        End If
        
        If strAssemblyChildName <> "" Then
            strFormattedAssemblyType = strAssemblyChildName
        End If
        
    ElseIf Not strFormat <> "" Then
    
        Dim oAssemblyName As IJNamedItem
        Set oAssemblyName = oAssemblyChild
        strFormattedAssemblyType = oAssemblyName.Name
        
    End If
    
    GetFormattedAssemblyType = strFormattedAssemblyType
CleanUp:
    Exit Function
ErrorHandler:
    GetFormattedAssemblyType = strFormattedAssemblyType
End Function

Private Function GetPartModelName(dispMfgPart As Object, strAttributeName As String, strAssemblyNameType As String, strProcess As String, strFormat As String, strBlockName As String, strAssemblyName As String, strBlockBoardSide As String) As String
    Const METHOD = "GetPartModelName"
    On Error GoTo ErrorHandler
    
    Dim strCommonPartName As String
    Dim oCommonPartNameAttribute As IXMLDOMAttribute
    Set oCommonPartNameAttribute = m_oXML.selectSingleNode("/SMS_SMS/*/SMS_PLATE/SMS_PROD_INFO/SMS_PART_INFO/SMS_COMMON_PART_INFO/@COMMON_PART_NAME")
    If oCommonPartNameAttribute Is Nothing Then
        strCommonPartName = "null"
    Else
        strCommonPartName = oCommonPartNameAttribute.Value
    End If
    
    Dim oNamedItem As IJNamedItem
    If TypeOf dispMfgPart Is IJMfgPlatePart Then
        Dim pPlatePart As IJPlatePart
        dispMfgPart.GetDetailedPart pPlatePart
        Set oNamedItem = pPlatePart
    ElseIf TypeOf dispMfgPart Is IJMfgProfilePart Then
        Dim oProfilePart As Object
        Set oProfilePart = dispMfgPart.GetDetailedPart
        Set oNamedItem = oProfilePart
    Else
        Exit Function
    End If
    
    If Not strAssemblyNameType <> "MODEL_PART_NAME" Then
        strCommonPartName = oNamedItem.Name
    Else
        strCommonPartName = "Not Implemented"
    End If
    
ErrorHandler:
CleanUp:
    GetPartModelName = strCommonPartName
End Function

Private Sub PropogateValue(oXMLDoc As Object, strXPathQuery As String, strAttributeName As String, ByVal strActionName As String, ByVal strProcess As String, ByVal strFormat As String)
    Const METHOD = "PropogateValue"
    On Error GoTo ErrorHandler
    
    Dim oXMLElement As IXMLDOMElement
    Set oXMLElement = oXMLDoc


    If InStr(strXPathQuery, "SMS_PART_INFO") > 0 Then
        Dim oPartInfoNodes As IXMLDOMNodeList
        Set oPartInfoNodes = oXMLElement.selectNodes("//SMS_PART_INFO")

        Dim nIndex As Integer
        For nIndex = 1 To oPartInfoNodes.Length - 1
        
            Dim oPartInfoElem As IXMLDOMElement
            Set oPartInfoElem = oPartInfoNodes.Item(nIndex)
            
            If Not oPartInfoElem Is Nothing Then
                'Get Attribute Value
                
                Dim strPartGuid As String: strPartGuid = "" 'MfgPart
                Dim strModelPartGuid As String: strModelPartGuid = "" 'DetailObject
                
                Dim oDispMfgObject As Object
                Dim oDispDetailObject As Object
                
                strPartGuid = oPartInfoElem.selectSingleNode("./@PART_GUID").nodeValue
                strModelPartGuid = oPartInfoElem.selectSingleNode("./@MODEL_PART_GUID").nodeValue
                
                Dim strPartAttributeValue As String
                If strPartGuid <> "" Then
                    Set oDispMfgObject = GetObjectFromOid("{" + strPartGuid + "}")
                    If Not oDispMfgObject Is Nothing Then
                        strPartAttributeValue = GetAttributeValue(oDispMfgObject, oPartInfoElem, strAttributeName, strActionName, strProcess, strFormat)
                        CreateNode oPartInfoElem, strAttributeName, False
                        SetNodeValue oPartInfoElem, strXPathQuery, strPartAttributeValue, strActionName
                    End If
                ElseIf strModelPartGuid <> "" Then
                    Set oDispDetailObject = GetObjectFromOid("{" + strModelPartGuid + "}")
                    If Not oDispMfgObject Is Nothing Then
                        strPartAttributeValue = GetAttributeValue(oDispDetailObject, oXMLElement, strAttributeName, strActionName, strProcess, strFormat)
                        CreateNode oPartInfoElem, strAttributeName, False
                        SetNodeValue oPartInfoElem, strXPathQuery, strPartAttributeValue, strActionName
                    End If
                Else
                    'Do nothing
                End If

            End If

        Next
    End If
    
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub PropogateValues(oXMLDoc As Object, strXPathQuery As String, strAttributeValue() As String)
    Const METHOD = "PropogateValues"
    On Error GoTo ErrorHandler
    
    Dim oXMLElement As IXMLDOMElement
    Set oXMLElement = oXMLDoc
    
    If InStr(strXPathQuery, "SMS_PART_INFO") > 0 Then
    'If Not Left(strXPathQuery, 15) <> "SMS_PART_INFO||" Then
        Dim oPartInfoNodes As IXMLDOMNodeList
        Set oPartInfoNodes = oXMLElement.selectNodes("//SMS_PART_INFO")
        
        Dim iIndex As Integer
        For iIndex = 1 To oPartInfoNodes.Length - 1
            SetNodeValues oXMLElement, Replace(strXPathQuery, "/@", "[" & iIndex & "]" & "/@"), strAttributeValue()
        Next
    End If
    
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function TruncateDouble(ByVal dValue As Double, ByVal iPrecision As Integer, ByVal IncludeLeadingDigit As VbTriState, ByVal UseParensForNegativeNumbers As VbTriState, ByVal GroupDigits As VbTriState) As String
    Const METHOD = "TruncateDouble"
    On Error GoTo ErrorHandler

    Dim strTruncatedDouble As String
    Dim iDeciamlLocation As Integer
    Dim strDouble As String
    
    strDouble = CStr(dValue)
    iDeciamlLocation = InStr(strDouble, ".")
    
    If iDeciamlLocation = 0 Then
        strTruncatedDouble = strDouble
    Else
        strTruncatedDouble = Left(strDouble, iDeciamlLocation + iPrecision)
    End If
    
    TruncateDouble = FormatNumber(CDbl(strTruncatedDouble), iPrecision, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
    
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IsStringInString(strTotalString As String, strSubString As String) As Boolean
    Const METHOD = "IsStringInString"
    On Error GoTo ErrorHandler
    IsStringInString = False
    Dim strSmallStrings() As String
    Dim lCount As Long
    strSmallStrings = Split(strTotalString, ",")
    For lCount = LBound(strSmallStrings) To UBound(strSmallStrings)
        If strSubString = Trim(strSmallStrings(lCount)) Then
            IsStringInString = True
            Exit Function
        End If
    Next lCount

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Private Sub SetNodeValue(oXMLDoc As IXMLDOMElement, strQuery As String, strNodeValue As String, ByVal strActionName As String)
    Const METHOD = "SetNodeValue"
    On Error GoTo ErrorHandler
        
    Dim oDataNode As IXMLDOMNode

    If strActionName = "SMS_SMS_ATTRIBUTE" Then
        Dim oRootDocument As IXMLDOMDocument
        Set oRootDocument = oXMLDoc.ownerDocument
        Set oDataNode = oRootDocument.selectSingleNode(strQuery)
        Set oRootDocument = Nothing
    Else
    Set oDataNode = oXMLDoc.selectSingleNode(strQuery)
    End If
    
    If Not oDataNode Is Nothing Then
        If TypeOf oDataNode Is IXMLDOMAttribute Then
            oDataNode.nodeValue = strNodeValue
        ElseIf TypeOf oDataNode Is IXMLDOMElement Then
            oDataNode.nodeTypedValue = strNodeValue
        End If
    Else
        Dim oDataElement As IXMLDOMElement
        Dim strSubQuery() As String
        strSubQuery = Split(strQuery, "/@")

        If UBound(strSubQuery) - LBound(strSubQuery) + 1 = 2 Then
            Set oDataNode = oXMLDoc.selectSingleNode(strSubQuery(0))
            
            If Not oDataNode Is Nothing Then
                Set oDataElement = oDataNode
                oDataElement.setAttribute strSubQuery(1), strNodeValue
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub SetNodeValues(oXMLDoc As IXMLDOMElement, strQuery As String, strNodeValues() As String)
    Const METHOD = "SetNodeValues"
    On Error GoTo ErrorHandler
    
    If UBound(strNodeValues) = 0 Then Exit Sub
        
    Dim oDataNodeList As IXMLDOMNodeList
    Set oDataNodeList = oXMLDoc.selectNodes(strQuery)
        
    Dim nCount As Integer
    Dim oDataNode As IXMLDOMNode
    If oDataNodeList.Length > 0 Then
        If oDataNodeList.Length <> UBound(strNodeValues) Then
            Exit Sub
        End If

        For nCount = 1 To UBound(strNodeValues)
            Set oDataNode = oDataNodeList.Item(nCount - 1)
            
            If Not oDataNode Is Nothing Then
                If TypeOf oDataNode Is IXMLDOMAttribute Then
                    oDataNode.nodeValue = strNodeValues(nCount)
                ElseIf TypeOf oDataNode Is IXMLDOMElement Then
                    oDataNode.nodeTypedValue = strNodeValues(nCount)
                End If
                Set oDataNode = Nothing
            End If
            
        Next nCount
    Else
        Dim oDataElement As IXMLDOMElement
        Dim strSubQuery() As String
        strSubQuery = Split(strQuery, "/@")

        If UBound(strSubQuery) - LBound(strSubQuery) + 1 = 2 Then
        

            Dim oDataNodeList1 As IXMLDOMNodeList
            Set oDataNodeList1 = oXMLDoc.selectNodes(strSubQuery(0))
            
            If oDataNodeList1.Length <> UBound(strNodeValues) Then
                Exit Sub
            Else
                For nCount = 1 To UBound(strNodeValues)
                    Set oDataNode = oDataNodeList1.Item(nCount - 1)
            
                    If Not oDataNode Is Nothing Then
                        Set oDataElement = oDataNode
                        oDataElement.setAttribute strSubQuery(1), strNodeValues(nCount)
                        Set oDataNode = Nothing
                    End If
                Next nCount
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function GetSymbolSharePath() As String
    Const METHOD = "GetSymbolLocationPath"
    On Error GoTo ErrorHandler
    
    Dim oContext As IJContext
    Dim strContextString As String
    Dim strSymbolShare As String
    
    strContextString = "OLE_SERVER"
    
    'Get IJContext
    Set oContext = GetJContext()
    
    If Not oContext Is Nothing Then
        strSymbolShare = oContext.GetVariable(strContextString)
    End If
    GetSymbolSharePath = strSymbolShare
    Set oContext = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function


'ProductionRouting
'RoutingAction
Private Function GetRelatedObjects(ByVal oObject As Object, ByVal strInterfaceID As String, ByVal strCollectionName As String) As IJElements
Const METHOD = "GetRelatedObjects"
On Error GoTo ErrorHandler

    Dim oAssocRel As IJDAssocRelation
    Dim oTargetObjCol As IJDTargetObjectCol
    
    Dim oResultColl As IJElements
    Set oResultColl = New JObjectCollection
    
    On Error Resume Next
    Set oAssocRel = oObject

    If Not oAssocRel Is Nothing Then
        Set oTargetObjCol = oAssocRel.CollectionRelations(strInterfaceID, strCollectionName)
        
        If oTargetObjCol.Count > 0 Then
            Dim i As Long
            
            For i = 1 To oTargetObjCol.Count
                oResultColl.Add oTargetObjCol.Item(i)
            Next
        End If
    End If
    
    Set GetRelatedObjects = oResultColl

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function



Private Function GetProductionRoutingInfo(ByVal oDispMfgPart As Object, ByVal strProcess As String, ByVal strFormat As String) As String
    Const METHOD = "GetProductionRoutingInfo"
    On Error GoTo ErrorHandler
       
    Dim strInterface As String
    Dim strAttribute As String

    strInterface = Left(strFormat, InStr(strFormat, "::") - 1)
    strAttribute = Right(strFormat, Len(strFormat) - Len(strInterface) - 2)
    
   
    GetProductionRoutingInfo = "Unknown"       'Default value
    
    Dim oDetailedObject As Object

    If TypeOf oDispMfgPart Is IJMfgPlatePart Then
        Dim oMfgPlatePart As IJMfgPlatePart
        Set oMfgPlatePart = oDispMfgPart
        oMfgPlatePart.GetDetailedPart oDetailedObject
        Set oMfgPlatePart = Nothing

    ElseIf TypeOf oDispMfgPart Is IJMfgProfilePart Then
        Dim oMfgProfilePart As IJMfgProfilePart
        Set oMfgProfilePart = oDispMfgPart
        Set oDetailedObject = oMfgProfilePart.GetDetailedPart
        Set oMfgProfilePart = Nothing
    Else
       'It can be DetailObject
       Set oDetailedObject = oDispMfgPart
    End If
    
'Private Const IID_IJDProductionRouting As String = "{B447C9B4-FB74-11D1-8A49-00A0C9065DF6}"
'Private Const IID_IJPlnProductionRouting As String = "{E9B1D9F6-C687-40BA-85E2-311534F4782A}"
'Private Const strProductionRoutingColName As String = "ProductionRouting"
'Private Const strRoutingActionColName As String = "RoutingAction"
    
    Dim oProductionRoutingObject As Object
    Dim oProductionRoutingColl As IJElements
    On Error Resume Next
    Set oProductionRoutingColl = GetRelatedObjects(oDetailedObject, IID_IJDProductionRouting, strProductionRoutingColName)
    
    If oProductionRoutingColl.Count > 0 Then
        Set oProductionRoutingObject = oProductionRoutingColl.Item(1)
    Else
        Exit Function
    End If
  

    Dim oMetaDataHelp      As IJDAttributeMetaData
    Dim oAttributes        As IJDAttributes
    Dim oAttrIID           As Variant
    
    Set oAttributes = oProductionRoutingObject
    Set oMetaDataHelp = oProductionRoutingObject

    For Each oAttrIID In oAttributes

        If oAttrIID = oMetaDataHelp.IID(strInterface) Then
            
            Dim oIJDAttributesCol  As IJDAttributesCol
            Dim oIJDInterfaceInfo  As IJDInterfaceInfo
            
            Set oIJDInterfaceInfo = oMetaDataHelp.InterfaceInfo(oAttrIID)
            Set oIJDAttributesCol = oAttributes.CollectionOfAttributes(oIJDInterfaceInfo.Type)
            
            Dim oAttribute As IJDAttribute
            For Each oAttribute In oIJDAttributesCol
                If oAttribute.AttributeInfo.Name = strAttribute Then
                    If oAttribute.AttributeInfo.CodeListTableName = "" Then
                        GetProductionRoutingInfo = CStr(oAttribute.Value)
                    Else
                        GetProductionRoutingInfo = GetShortDescription(oProductionRoutingObject, oAttribute.AttributeInfo.CodeListTableName, oAttribute.Value)
                    End If
                End If
                Set oAttribute = Nothing
            Next
            
            Set oIJDInterfaceInfo = Nothing
            Set oIJDAttributesCol = Nothing
            
        End If
    Next

ErrorHandler:
CleanUp:
    Set oDetailedObject = Nothing
    Set oProductionRoutingObject = Nothing
    Set oAttributes = Nothing
    Set oMetaDataHelp = Nothing
    
    Exit Function
End Function

Private Function GetProjectInfo(ByVal oDispMfgPart As Object, ByVal strProcess As String, ByVal strFormat As String) As String
    Const METHOD = "GetProductionRoutingInfo"
    On Error GoTo ErrorHandler
       
    Dim strInterface As String
    Dim strAttribute As String

    strInterface = Left(strFormat, InStr(strFormat, "::") - 1)
    strAttribute = Right(strFormat, Len(strFormat) - Len(strInterface) - 2)
    
    
    GetProjectInfo = "Unknown"       'Default value

    
    Dim oPrjectRoot As Object
    Set oPrjectRoot = GetActiveConfigShipClass()
    
    If oPrjectRoot Is Nothing Then
        Exit Function
    End If
  

    Dim oMetaDataHelp      As IJDAttributeMetaData
    Dim oAttributes        As IJDAttributes
    Dim oAttrIID           As Variant
    
    Set oAttributes = oPrjectRoot
    Set oMetaDataHelp = oPrjectRoot

    For Each oAttrIID In oAttributes

        If oAttrIID = oMetaDataHelp.IID(strInterface) Then
            
            Dim oIJDAttributesCol  As IJDAttributesCol
            Dim oIJDInterfaceInfo  As IJDInterfaceInfo
            
            Set oIJDInterfaceInfo = oMetaDataHelp.InterfaceInfo(oAttrIID)
            Set oIJDAttributesCol = oAttributes.CollectionOfAttributes(oIJDInterfaceInfo.Type)
            
            Dim oAttribute As IJDAttribute
            For Each oAttribute In oIJDAttributesCol
                If oAttribute.AttributeInfo.Name = strAttribute Then
                    If oAttribute.AttributeInfo.CodeListTableName = "" Then
                        GetProjectInfo = CStr(oAttribute.Value)
                    Else
                        GetProjectInfo = GetShortDescription(oPrjectRoot, oAttribute.AttributeInfo.CodeListTableName, oAttribute.Value)
                    End If
                End If
                Set oAttribute = Nothing
            Next
            
            Set oIJDInterfaceInfo = Nothing
            Set oIJDAttributesCol = Nothing
            
        End If
    Next

ErrorHandler:
CleanUp:
    Set oPrjectRoot = Nothing
    Set oAttributes = Nothing
    Set oMetaDataHelp = Nothing
    
    Exit Function
End Function

Private Function GetActiveConfigShipClass() As IJProjectRoot
    Const METHOD = "GetActiveConfigShipClass"
    On Error GoTo ErrorHandler

    Dim oMfgEntityHelper As IJMfgEntityHelper
    Set oMfgEntityHelper = New MfgEntityHelper

    Set GetActiveConfigShipClass = oMfgEntityHelper.GetConfigProjectRoot
    
    Set oMfgEntityHelper = Nothing
    
    Exit Function

ErrorHandler:
CleanUp:
    
    
End Function
Private Function GetProductionRoutingActionInfos(ByVal oDispMfgPart As Object, ByVal strAttribute As String, ByVal strAction As String, ByVal strProcess As String, ByVal strFormat As String) As String()
    Const METHOD = "GetProductionRoutingActionInfos"
    On Error GoTo ErrorHandler
      
    Dim strInterface As String
    Dim strInterfaceAttribute As String
    If InStr(strFormat, "::") > 0 Then
        strInterface = Left(strFormat, InStr(strFormat, "::") - 1)
        strInterfaceAttribute = Right(strFormat, Len(strFormat) - Len(strInterface) - 2)
    Else
        strInterface = "IJPlnRoutingAction"
        If InStr(strAttribute, "TYPE") > 0 Then
            strInterfaceAttribute = "Action"
        ElseIf InStr(strAttribute, "CODE") > 0 Then
            strInterfaceAttribute = "Code"
        ElseIf InStr(strAttribute, "MACHINE") > 0 Then
            strInterfaceAttribute = "Machine"
        End If
        
    End If
    

    Dim strEmptyValues() As String
    ReDim strEmptyValues(0 To 0)
    GetProductionRoutingActionInfos = strEmptyValues    'Default value
    Erase strEmptyValues
    
    Dim oDetailedObject As Object

    If TypeOf oDispMfgPart Is IJMfgPlatePart Then
        Dim oMfgPlatePart As IJMfgPlatePart
        Set oMfgPlatePart = oDispMfgPart
        oMfgPlatePart.GetDetailedPart oDetailedObject
        Set oMfgPlatePart = Nothing

    ElseIf TypeOf oDispMfgPart Is IJMfgProfilePart Then
        Dim oMfgProfilePart As IJMfgProfilePart
        Set oMfgProfilePart = oDispMfgPart
        Set oDetailedObject = oMfgProfilePart.GetDetailedPart
        Set oMfgProfilePart = Nothing
    Else
        'Detailed Part
        Set oDetailedObject = oDispMfgPart
        
    End If
    
'Private Const IID_IJDProductionRouting As String = "{B447C9B4-FB74-11D1-8A49-00A0C9065DF6}"
'Private Const IID_IJPlnProductionRouting As String = "{E9B1D9F6-C687-40BA-85E2-311534F4782A}"
'Private Const strProductionRoutingColName As String = "ProductionRouting"
'Private Const strRoutingActionColName As String = "RoutingAction"
    
    Dim oProductionRoutingObject As Object
    Dim oProductionRoutingColl As IJElements
    Dim oProductionRoutingActionColl As IJElements
    Set oProductionRoutingColl = GetRelatedObjects(oDetailedObject, IID_IJDProductionRouting, strProductionRoutingColName)
    
    If oProductionRoutingColl.Count > 0 Then
        Set oProductionRoutingObject = oProductionRoutingColl.Item(1)
        Set oProductionRoutingActionColl = GetRelatedObjects(oProductionRoutingObject, IID_IJPlnProductionRouting, strRoutingActionColName)
    Else
        Exit Function
    End If
    
    
    
    If oProductionRoutingActionColl.Count > 0 Then
  
        Dim nCount As Long
        Dim strTempReturnColl() As String
        ReDim strTempReturenColl(1 To oProductionRoutingActionColl.Count)
        Dim strReturnColl() As String

        Dim nActionCount As Long: nActionCount = 0
        
        For nCount = 1 To oProductionRoutingActionColl.Count
            Dim strProcessValue As String
            
            If strAction = "ROUTING_ACTION_TYPE" Then
                strProcessValue = GetRoutingActionName(oProductionRoutingActionColl.Item(nCount), "Action")
                
            ElseIf strAction = "ROUTING_ACTION_CODE" Then
                strProcessValue = GetRoutingActionName(oProductionRoutingActionColl.Item(nCount), "Code")

            End If
        
            Dim oMetaDataHelp      As IJDAttributeMetaData
            Dim oAttributes        As IJDAttributes
            Dim oAttrIID           As Variant
            
            Set oAttributes = oProductionRoutingActionColl.Item(nCount)
            Set oMetaDataHelp = oProductionRoutingActionColl.Item(nCount)

            If strProcess = "" Or (strProcessValue <> "" And InStr(strProcess, strProcessValue) > 0) Then

                nActionCount = nActionCount + 1
                If InStr(strAttribute, "SEQUENCE") > 0 Then
                    strTempReturenColl(nActionCount) = CStr(nCount)
                Else
                    For Each oAttrIID In oAttributes
            
                        If oAttrIID = oMetaDataHelp.IID(strInterface) Then
                            
                            Dim oIJDAttributesCol  As IJDAttributesCol
                            Dim oIJDInterfaceInfo  As IJDInterfaceInfo
                            
                            Set oIJDInterfaceInfo = oMetaDataHelp.InterfaceInfo(oAttrIID)
                            Set oIJDAttributesCol = oAttributes.CollectionOfAttributes(oIJDInterfaceInfo.Type)
                            
                            Dim oAttribute As IJDAttribute
                            For Each oAttribute In oIJDAttributesCol
                                    
                                If oAttribute.AttributeInfo.Name = strInterfaceAttribute Then
                                    If oAttribute.AttributeInfo.CodeListTableName = "" Then
                                        strTempReturenColl(nActionCount) = CStr(oAttribute.Value)
                                    Else
                                        strTempReturenColl(nActionCount) = GetShortDescription(oProductionRoutingObject, oAttribute.AttributeInfo.CodeListTableName, oAttribute.Value)
                                    End If
                                End If

                                Set oAttribute = Nothing
                                
                            Next
                            
                            Set oIJDInterfaceInfo = Nothing
                            Set oIJDAttributesCol = Nothing
                            
                        End If
                    Next
                End If
            End If
        Next nCount
        

        Dim i As Integer
        ReDim strReturnColl(1 To nActionCount)
        For i = 1 To nActionCount
            strReturnColl(i) = strTempReturenColl(i)
        Next
        
        GetProductionRoutingActionInfos = strReturnColl
        Erase strReturnColl
        Erase strTempReturenColl
    End If
    


ErrorHandler:
CleanUp:
    Set oDetailedObject = Nothing
    Set oProductionRoutingObject = Nothing
    Set oAttributes = Nothing
    Set oMetaDataHelp = Nothing
    
    Exit Function
End Function


Private Function GetNestingInfo(ByVal oDispMfgPart As Object, ByVal oXMLElement As IXMLDOMElement, ByVal strProcess As String, ByVal strFormat As String) As String
    Const METHOD = "GetNestingInfo"
    On Error GoTo ErrorHandler

       
    Dim strInterface As String
    Dim strAttribute As String
    strInterface = Left(strFormat, InStr(strFormat, "::") - 1)
    strAttribute = Right(strFormat, Len(strFormat) - Len(strInterface) - 2)
    
    GetNestingInfo = ""

    
'Private Const IID_IJMfgNestData As String = "{176B4462-2A91-44CE-8D3F-C5B994C283B6}"
'Private Const strNestingDataColName As String = "MfgPartNestData_ORIG"

    On Error Resume Next
    Dim oNestData As Object
    
    Dim oNestDataColl As IJElements
    Set oNestDataColl = GetRelatedObjects(oDispMfgPart, IID_IJMfgOuptput, strNestingDataColName)

    If oNestDataColl.Count = 0 Then
       Exit Function
    End If
  
    Dim nCount As Long

    Dim strNestGUIDInXML As String
    strNestGUIDInXML = oXMLElement.selectSingleNode("./@NEST_GUID").nodeValue
    
    For nCount = 1 To oNestDataColl.Count
        Dim oMetaDataHelp      As IJDAttributeMetaData
        Dim oAttributes        As IJDAttributes
        Dim oAttrIID           As Variant
    
        Set oNestData = oNestDataColl.Item(nCount)
        
        Dim strNestGUID As String
        strNestGUID = CleanGUID(GetOID(oNestData))
        
        If strNestGUID = strNestGUIDInXML Then
        
            Set oAttributes = oNestData
            Set oMetaDataHelp = oNestData
    
            For Each oAttrIID In oAttributes
        
                If oAttrIID = oMetaDataHelp.IID(strInterface) Then
                    
                    Dim oIJDAttributesCol  As IJDAttributesCol
                    Dim oIJDInterfaceInfo  As IJDInterfaceInfo
                    
                    Set oIJDInterfaceInfo = oMetaDataHelp.InterfaceInfo(oAttrIID)
                    Set oIJDAttributesCol = oAttributes.CollectionOfAttributes(oIJDInterfaceInfo.Type)
                    
                    Dim oAttribute As IJDAttribute
                    For Each oAttribute In oIJDAttributesCol
                        If oAttribute.AttributeInfo.Name = strAttribute Then
                            If oAttribute.AttributeInfo.CodeListTableName = "" Then
                                GetNestingInfo = CStr(oAttribute.Value)
                                Exit For
                            Else
                                GetNestingInfo = GetShortDescription(oNestData, oAttribute.AttributeInfo.CodeListTableName, oAttribute.Value)
                                Exit For
                            End If
                        End If
                        Set oAttribute = Nothing
                    Next
                    
                    Set oIJDInterfaceInfo = Nothing
                    Set oIJDAttributesCol = Nothing
                    
                End If
            Next
            
            Set oNestData = Nothing
            Set oAttributes = Nothing
            Set oMetaDataHelp = Nothing
        End If
    Next



ErrorHandler:
CleanUp:
    Set oNestData = Nothing
    Set oAttributes = Nothing
    Set oMetaDataHelp = Nothing
    Set oNestDataColl = Nothing
    
    Exit Function
End Function

'If bCreateNew is true, create new node evenif there is existing one
'If bCreateNew is false, don't create new node if there is existing one
Private Sub CreateNode(ByVal oXMLDoc As IXMLDOMElement, ByVal strAttributeName As String, ByVal bCreateNew As Boolean)
    Const METHOD = "CreateNode"
    On Error GoTo ErrorHandler
    
    Dim strNodes() As String
    strNodes = Split(strAttributeName, "||")
        
    Dim oParentNode As IXMLDOMNode
    Set oParentNode = oXMLDoc
    
    Dim oSMSOutputDocument As IXMLDOMDocument
    Set oSMSOutputDocument = m_oXML.ownerDocument
    
    Dim lCount As Long
    Dim strXPathQuery As String: strXPathQuery = "./"
    For lCount = LBound(strNodes) To UBound(strNodes) - 1
        strXPathQuery = strXPathQuery + strNodes(lCount)
        'Query Path
        On Error Resume Next
        Dim oXMLNodeList As IXMLDOMNodeList
        Set oXMLNodeList = oXMLDoc.selectNodes(strXPathQuery)
   
        Dim oNewNode As IXMLDOMNode
        Dim oCreratedNewNode As IXMLDOMNode
        
        If oXMLNodeList.Length = 0 Then
            Set oCreratedNewNode = oSMSOutputDocument.CreateNode(NODE_ELEMENT, strNodes(lCount), "")
            Set oNewNode = oParentNode.appendChild(oCreratedNewNode)
            
            Set oParentNode = Nothing
            Set oParentNode = oNewNode
            Set oNewNode = Nothing
        Else
            If lCount = UBound(strNodes) - 1 And bCreateNew = True Then
                Set oCreratedNewNode = oSMSOutputDocument.CreateNode(NODE_ELEMENT, strNodes(lCount), "")
                Set oNewNode = oParentNode.appendChild(oCreratedNewNode)
            
                Set oParentNode = Nothing
                Set oParentNode = oNewNode
                Set oNewNode = Nothing
            Else
                Set oParentNode = Nothing
                Set oParentNode = oXMLNodeList.Item(0)
            End If
        End If
        
        Set oXMLNodeList = Nothing
        strXPathQuery = strXPathQuery + "/"
        
    Next lCount
    
    
    Erase strNodes
    Set oParentNode = Nothing
    Set oSMSOutputDocument = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub


Private Function GetRoutingActionName(ByVal oAction As Object, ByVal strAttributeName As String) As String
On Error GoTo ErrorHandler
Const METHOD = "GetRoutingActionName"
    
    Dim oAttributeMetaData As IJDAttributeMetaData
    Dim oAttrHelper As IJDAttributes
    Dim oAttributesCollection As IJDAttributesCol
    Dim oInterfaceInfo As IJDInterfaceInfo
     
    Set oAttributeMetaData = oAction
    Set oAttrHelper = oAction
    Set oInterfaceInfo = oAttributeMetaData.InterfaceInfo(oAttributeMetaData.IID("IJPlnRoutingAction"))
    
    If Not oInterfaceInfo Is Nothing Then
        On Error Resume Next
        Set oAttributesCollection = oAttrHelper.CollectionOfAttributes(oInterfaceInfo.Type)
        Err.Clear
        On Error GoTo ErrorHandler
        If Not oAttributesCollection Is Nothing Then
        
            Dim oAttribute As IJDAttribute
            For Each oAttribute In oAttributesCollection
                If oAttribute.AttributeInfo.Name = strAttributeName Then
                    If oAttribute.AttributeInfo.CodeListTableName = "" Then
                        GetRoutingActionName = CStr(oAttribute.Value)
                    Else
                        GetRoutingActionName = GetShortDescription(oAction, oAttribute.AttributeInfo.CodeListTableName, oAttribute.Value)
                    End If
                End If
                Set oAttribute = Nothing
            Next
        End If
    End If


ErrorHandler:
CleanUp:
    Set oAttributeMetaData = Nothing
    Set oAttrHelper = Nothing
    Set oAttributesCollection = Nothing
    Set oInterfaceInfo = Nothing
    
    Exit Function
End Function


Private Function GetResourceManager() As IJDPOM
    Const METHOD = "GetActiveConnection"
    On Error GoTo ErrorHandler
    
    Dim oCmnAppGenericUtil As IJDCmnAppGenericUtil
    Set oCmnAppGenericUtil = New CmnAppGenericUtil
    
    Dim oPOM As IUnknown
    oCmnAppGenericUtil.GetResourceManager oPOM
    
    Set GetResourceManager = oPOM

    Set oCmnAppGenericUtil = Nothing
    Exit Function
    
ErrorHandler:
    Err.Clear

End Function


Private Function GetOID(oObject As Object) As String
    On Error GoTo ErrorHandler
    Const METHOD = "GetOID"
    
    Dim oPOM As IJDPOM
    Set oPOM = GetResourceManager()

    'Retrive the OID from the Manufacturing Object
    Dim oMoniker As IMoniker
    Set oMoniker = oPOM.GetObjectMoniker(oObject)
    GetOID = oPOM.DbIdentifierFromMoniker(oMoniker)
    
    Set oMoniker = Nothing
    
    Exit Function
    
ErrorHandler:
    Resume Next
End Function

Private Function CleanGUID(sGUID As String) As String
    On Error GoTo ErrorHandler
    Const METHOD = "CleanGUID"
    
    Dim sTemp As String
    
    If InStr(1, sGUID, "{") >= 1 And InStr(1, sGUID, "}") >= 1 Then
        'The string has brackets in it.
        sTemp = Mid(sGUID, InStr(1, sGUID, "{") + 1, Len(sGUID))
        CleanGUID = Mid(sTemp, 1, InStr(1, sTemp, "}") - 1)
    Else
        CleanGUID = sGUID
    End If
    Exit Function
ErrorHandler:
    Resume Next
End Function

Private Function GetObjectFromOid(ByVal strOid As String) As Object
    On Error GoTo ErrorHandler
    Const METHOD = "GetOID"
    
    Dim oPOM As IJDPOM
    Set oPOM = GetResourceManager()

    Dim oMoniker As IMoniker
    Set oMoniker = oPOM.MonikerFromDbIdentifier(strOid)
    
    Set GetObjectFromOid = oPOM.GetObject(oMoniker)
    
    Set oMoniker = Nothing
    
    Exit Function
    
ErrorHandler:
    Resume Next
End Function




