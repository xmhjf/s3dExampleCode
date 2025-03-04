Attribute VB_Name = "MarineLibraryCommon"
Option Explicit

Public Const CUSTOMERID = "SM"
Private Const MODULE = "StructDetail\Include\MarineLibraryCommon.bas"


Public Sub GetSelectorAnswer(oOccurrence As Object, strQuestion As String, _
                                strAnswer As Variant, _
                                Optional SmartObject As Object)

    Dim oSmartItem As IJSmartItem
    Dim oSmartOccurrence As DEFINITIONHELPERSINTFLib.IJSmartOccurrence
    
    Dim oParameterLogic As IJDParameterLogic
    Dim oMemberDescription As IJDMemberDescription
    Dim oSelectorLogic As IJDSelectorLogic
    
    Dim oParentSmartClass As IJSmartClass
    Dim oParentSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    
    Dim oCommonHelper As New DefinitionHlprs.CommonHelper
    
      
    If TypeOf oOccurrence Is IJDMemberDescription Then
        Set oMemberDescription = oOccurrence
        Set oSmartOccurrence = oMemberDescription.CAO
        Set oSmartItem = oSmartOccurrence.ItemObject
        Set oParentSmartClass = oSmartItem.Parent
        
    ElseIf TypeOf oOccurrence Is IJDParameterLogic Then
        Set oParameterLogic = oOccurrence
        Set oSmartItem = oParameterLogic.SmartItem
        Set oSmartOccurrence = oParameterLogic.SmartOccurrence
        Set oParentSmartClass = oSmartItem.Parent
        
    ElseIf TypeOf oOccurrence Is IJSmartOccurrence Then
         Set oSmartOccurrence = oOccurrence
         Set oSmartItem = oSmartOccurrence.ItemObject
         Set oParentSmartClass = oSmartItem.Parent
         
    ElseIf TypeOf oOccurrence Is IJDSelectorLogic Then
         Set oSelectorLogic = oOccurrence
         Set oSmartOccurrence = oSelectorLogic.SmartOccurrence
         strAnswer = oSelectorLogic.Answer(strQuestion)
         Set oParentSmartClass = oSmartOccurrence.RootSelectionObject
         
    ElseIf TypeOf oOccurrence Is IJSmartClass Then
        Set oSmartOccurrence = SmartObject
        Set oParentSmartClass = oOccurrence
    
    End If
    
    Set oParentSymbolDefinition = oParentSmartClass.SelectionRuleDef
    
    Set oCommonHelper = New DefinitionHlprs.CommonHelper
    
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oSymbolDefinition = oParentSmartClass.SelectionRuleDef
    
    Dim pAttributes As IJDAttributes
    Dim pAttrCol As IJDAttributesCol
    Dim pCmd As IJDCommandDescription
    Dim pAtt As IJDAttribute
    Dim pObject As IJDObject
    Dim pCodeListMD As IJDCodeListMetaData
    
    Set pObject = oSmartOccurrence
    Set pCodeListMD = pObject.ResourceManager
    
    Dim interfaceName As String
    
    'Fix for TR-231923 GetSelectorAnswer() in MarineLibraryCommon is giving exception
    'Should work for .NET and COM. Currently only excersied for .NET as the SelectionRuleDef is null for .NET by design
    If oSymbolDefinition Is Nothing Then
        interfaceName = oParentSmartClass.SelectionRuleInterface
    Else
        On Error Resume Next
        Set pCmd = oSymbolDefinition.IJDUserCommands.GetCommand("AnswersIID")
        If Len(pCmd.Source) < 1 Then
            Exit Sub
        End If
        On Error GoTo ErrorHandler:
        interfaceName = pCmd.Source
    End If
    
    Set pAttributes = oSmartOccurrence
    Set pAttrCol = pAttributes.CollectionOfAttributes(interfaceName)
    
    If Not pAttrCol Is Nothing Then
        For Each pAtt In pAttrCol
            If strQuestion = pAtt.AttributeInfo.Name Then
                If pAtt.AttributeInfo.CodeListTableName <> "" Then
                    If pAtt.Value < 65536 Then
                        strAnswer = pAtt.Value
                    ElseIf pAtt.Value >= (65536 * 2) Then ' it's a double code list
                        strAnswer = CDbl(pCodeListMD.ShortStringValue(pAtt.AttributeInfo.CodeListTableName, pAtt.Value))
                    Else ' it is a string code list
                        strAnswer = pCodeListMD.ShortStringValue(pAtt.AttributeInfo.CodeListTableName, pAtt.Value)
                    End If
                Else
                    strAnswer = pAtt.Value
                End If
            Else
                If CStr(strAnswer) = "0" Then
                    On Error Resume Next
                    strAnswer = CStr("")
               End If
            End If
        Next pAtt
    End If
    
     If strAnswer = vbEmpty Or Len(CStr(strAnswer)) = 0 Then
        Dim oGrandParentSmartClass As DEFINITIONHELPERSINTFLib.IJSmartClass
         If TypeOf oParentSmartClass Is DEFINITIONHELPERSINTFLib.IJSmartClass Then
            If TypeOf oParentSmartClass.Parent Is DEFINITIONHELPERSINTFLib.IJSmartClass Then
             Set oGrandParentSmartClass = oParentSmartClass.Parent
             GetSelectorAnswer oGrandParentSmartClass, strQuestion, strAnswer, oSmartOccurrence
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Sub



