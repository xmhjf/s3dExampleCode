Attribute VB_Name = "Utilities"
'********************************************************************
' Copyright (C) 1998-2006 Intergraph Corporation.  All Rights Reserved.
'
' File: Utilities.bas
'
' Author: D.A. Trent
'
' Abstract:
'
'********************************************************************
Option Explicit
'
Private Const Module = "CheckMfctyTeeWeldChamfer.Utilities"
Private Const AnswersCMD As String = "AnswersIID"

Public m_bLogError As Boolean
'

'********************************************************************
' ' Routine: LogError
'
' Description:  default Error Reporter
'********************************************************************
Public Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
     
    m_bLogError = True
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
End Function

'******************************************************************************
' Routine: GetPOM
'
' Abstract:
' This method returns the Persistent Object Manager (POM) for the
' specified database.
'
' Description:
'
' Inputs:
'
' Outputs:
'
'******************************************************************************
Public Function GetPOM(databaseType As String) As IJDPOM
    Const Method As String = "GetPOM"
    On Error GoTo ErrorHandler

    Dim sProgressMessage As String
    sProgressMessage = ""
        
    ' Get the Server Context.
    Dim oPOM As IJDPOM
    Dim oContext As IJContext
    Dim oAccessMiddle As IJDAccessMiddle
    
    sProgressMessage = "Get the Server Context"
    Set oContext = GetJContext()

    ' Get the AccessMiddle object.
    sProgressMessage = "Get the IJDAccessMiddle"
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")

    ' Get the Persistent Object Managers (POM).
    sProgressMessage = "Get the Persistent Object Manager (POM)"
    Set oPOM = oAccessMiddle.GetResourceManagerFromType(databaseType)
   
    ' Return the POM.
    sProgressMessage = "Return the POM"
    Set GetPOM = oPOM
   
    Set oContext = Nothing
    Set oAccessMiddle = Nothing
    Set oPOM = Nothing

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, Method, sProgressMessage).Number

End Function

'******************************************************************************
' Routine: GetParentSmartObject
'
' Abstract:
' Given an SmartItem object
'   Gets the owning Smart Parent object
'   (the SmartObject that has the given Smart Item as an Item)
'
' Description:
'
' Inputs:
'
' Outputs:
'
'******************************************************************************
Public Sub GetParentSmartObject(oItemObject As Object, _
                                oParentObject As Object)
Const Method As String = "GetParentSmartObject"
On Error GoTo ErrorHandler
    
    Dim nCount As Long
    
    Dim oObject As Object
    Dim oNamedItem As IJNamedItem
    Dim oRelationShip As IJDRelationship
    Dim oAssocRelation As IJDAssocRelation
    Dim oRelationShipCol As IJDRelationshipCol
    
    ' Get the given's SmartOrrurence's Parent:
    '   the SmartOccurence that the SmartItem was created by
    Set oParentObject = Nothing
    If Not TypeOf oItemObject Is IJDAssocRelation Then
        Exit Sub
    End If
    
    Set oAssocRelation = oItemObject
    Set oObject = oAssocRelation.CollectionRelations("IJFullObject", _
                                                     "toAssembly")
    If oObject Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oObject Is IJDRelationshipCol Then
        Exit Sub
    End If
    
    Set oRelationShipCol = oObject
    nCount = oRelationShipCol.Count
                      
    If nCount < 1 Then
        Exit Sub
    End If
    
    Set oRelationShip = oRelationShipCol.Item(1)
    Set oParentObject = oRelationShip.Target
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, Method).Number

End Sub

'********************************************************************
' Routine: GetParameterRuleOutputs
'
' Abstract: Retrieve the collection of Parameter Rule Attributes outputs
'           for the given the Smart Occurrence
'********************************************************************
Private Sub GetParameterRuleOutputs(oSmartOcc As IJSmartOccurrence, _
                                    oOutputColl As IJDOutputs, _
                                    oAttrColl As IJDAttributesCol)
Const Method As String = "GetParentSmartObject"
On Error GoTo ErrorHandler
    
    Dim oSmartItem As IJSmartItem
    Dim oSymbolDefinition As IJDSymbolDefinition
    Set oSmartItem = oSmartOcc.ItemObject
    
    On Error Resume Next
    Set oSymbolDefinition = oSmartItem.ParameterRuleDef
    Set oSmartItem = Nothing
    On Error GoTo ErrorHandler
    
    'No parameter rule exists for Smart Item
    If oSymbolDefinition Is Nothing Then Exit Sub
    
    Dim oOutputCollection As IJDOutputCollection
    Dim oSmartOccHelper As IJSmartOccurrenceHelper
    Set oSmartOccHelper = New CSmartOccurrenceCES
    Set oOutputCollection = oSmartOccHelper.ExecuteParameterRule(oSymbolDefinition, _
                                                           oSmartOcc)
    
    'Get the Answer interface
    Dim oCmd As IJDCommandDescription
    Dim oOutput As IJDOutput
    Dim iOutputIndex As Integer
    Dim oOutputControl As IJOutputControl
    Dim bCanModify As Boolean
    
    Set oCmd = oSymbolDefinition.IJDUserCommands.GetCommand(AnswersCMD)
    If Len(oCmd.Source) = 0 Then
        'No parameter rule outputs defined
        Exit Sub
    Else
        Dim oAttributes As IJDAttributes
        Set oAttributes = oSmartOcc
        Set oAttrColl = oAttributes.CollectionOfAttributes(oCmd.Source)
        Set oAttributes = Nothing
    End If
    
    Set oOutputColl = oSymbolDefinition.IJDRepresentations(1).IJDOutputs
    
    Set oSmartOccHelper = Nothing
    Set oOutputCollection = Nothing
    Set oSymbolDefinition = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, Module, Method).Number

End Sub

' ********************************************************************************
' Method: GetSmartOccAnswer
'
' Abstract:
'   Gets Smart Occurrence / Symbol Question/Answer value
'
' Inputs:
'
' Outputs:
' ********************************************************************************
Public Sub GetSmartOccAnswer(oSmartObject As Object, _
                             oSmartClassObject As Object, _
                             sQuestion As String, sAnswer As String)
Const Method As String = "GetParentSmartObject"
On Error GoTo ErrorHandler
     
    Dim vAnswer As Variant
    
    Dim oParameterLogic As IJDParameterLogic
    Dim oMemberDescription As IJDMemberDescription
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartOccurrence As DEFINITIONHELPERSINTFLib.IJSmartOccurrence
    
    Dim oSmartClass As IJSmartClass
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    On Error GoTo ErrorHandler
    sAnswer = ""
    
    If TypeOf oSmartObject Is IJSmartOccurrence Then
        Set oSmartOccurrence = oSmartObject
        Set oSmartItem = oSmartOccurrence.ItemObject

    ElseIf TypeOf oSmartObject Is IJDMemberDescription Then
        Set oMemberDescription = oSmartObject
        Set oSmartOccurrence = oMemberDescription.CAO
        Set oSmartItem = oSmartOccurrence.ItemObject
        
    ElseIf TypeOf oSmartObject Is IJDParameterLogic Then
        Set oParameterLogic = oSmartObject
        Set oSmartItem = oParameterLogic.SmartItem
        Set oSmartOccurrence = oParameterLogic.SmartOccurrence
    
    Else
        Exit Sub
    End If
    
    ' Set the IJSmartClass object to retrieve the Selection Questions/Answers
    If Not oSmartClassObject Is Nothing Then
        If TypeOf oSmartClassObject Is IJSmartClass Then
            Set oSmartClass = oSmartClassObject
        Else
            Exit Sub
        End If
    
    ElseIf oSmartClass Is Nothing Then
        If oSmartItem Is Nothing Then
            Exit Sub
        End If
        Set oSmartClass = oSmartItem.Parent
    End If
    
    Set oSymbolDefinition = oSmartClass.SelectionRuleDef

    Set oCommonHelper = New DefinitionHlprs.CommonHelper
    vAnswer = oCommonHelper.GetAnswer(oSmartOccurrence, oSymbolDefinition, _
                                      sQuestion)
    sAnswer = vAnswer
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, Method).Number
End Sub

' ********************************************************************************
' Method: GetSmartSelector
'
' Abstract:
'   Gets Smart Occurrence / Symbol available Selection Rule SmartClass(es)
'
' Inputs:
'
' Outputs:
' ********************************************************************************
Public Sub GetSmartSelector(oSmartObject As Object, _
                            sSmartSelector As String, _
                            oSelectorCollection As Collection, _
                            Optional bSeachParents As Boolean = False)
Const Method As String = "GetSmartSelector"
On Error GoTo ErrorHandler

    Dim bContinue As Boolean
    Dim bSelectionRule As Boolean
    Dim sSelectionRule As String
    
    Dim oParentObject As Object
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oParameterLogic As IJDParameterLogic
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oMemberDescription As IJDMemberDescription
    
    
    ' Initalize output Collection (if required)
    If oSelectorCollection Is Nothing Then
        Set oSelectorCollection = New Collection
    End If
    
    ' Check Type of SmartObject passed in as input
    ' expect to begin with a SmartOcurrence object
    ' but:
    '   handle cases where we know how to get SmartOcurrence from given object
    If TypeOf oSmartObject Is IJDMemberDescription Then
        Set oMemberDescription = oSmartObject
        Set oSmartOccurrence = oMemberDescription.CAO
        
    ElseIf TypeOf oSmartObject Is IJDParameterLogic Then
        Set oParameterLogic = oSmartObject
        Set oSmartItem = oParameterLogic.SmartItem
        Set oSmartOccurrence = oParameterLogic.SmartOccurrence
         
    ElseIf TypeOf oSmartObject Is IJSmartOccurrence Then
        Set oSmartOccurrence = oSmartObject
    Else
        Exit Sub
    End If
    
    ' verify that we have a valid starting SmartOccurrence
    If oSmartOccurrence Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oSmartOccurrence Is IJSmartOccurrence Then
        Exit Sub
    End If
        

    ' Get the Current SmartItem from the SmartOcurrence
    ' The SmartItem is (normally) the Last object in the SmartOcurrence
    Set oSmartItem = oSmartOccurrence.ItemObject
    
    ' Loop thru the SmartItem Parents
    ' expect the parent objects to be IJSmartClass objects
    ' (the parents will be all the different Selectors used)
    ' continue looping until the parent object is NOT an IJSmartClass object
    Set oParentObject = oSmartItem.Parent
    If TypeOf oParentObject Is IJSmartClass Then
        Set oSmartClass = oParentObject
        Set oParentObject = Nothing
    Else
        Set oSmartClass = Nothing
    End If
    
    bSelectionRule = False
    If Not oSmartClass Is Nothing Then
        
        bContinue = True
        While bContinue
            ' Check if Current SmartClass has a Selection Rule
            ' Check if Current Selection Rule is the one requested
            If Len(Trim(oSmartClass.SelectionRule)) > 0 Then
                sSelectionRule = Trim(oSmartClass.SelectionRule)
                If Len(Trim(sSmartSelector)) < 1 Then
                    ' Call Requesting All Selection Rules
                    oSelectorCollection.Add oSmartClass
                
                ElseIf LCase(Trim(sSmartSelector)) = LCase(sSelectionRule) Then
                    bContinue = False
                    bSelectionRule = True
                    oSelectorCollection.Add oSmartClass
                End If
            End If
            
            If bContinue Then
                ' Get the Parent of the current Smart Class
                Set oParentObject = oSmartClass.Parent
                If TypeOf oParentObject Is IJSmartClass Then
                    ' the current parent is an IJSmartClass object,
                    ' continue loop thur IJSmartClass parents
                    Set oSmartClass = oParentObject
                    Set oParentObject = Nothing
                Else
                    ' the current parent is NOT an IJSmartClass object,
                    ' reach end of parent loop
                    Set oSmartClass = Nothing
                    bContinue = False
                End If
            End If
        Wend
    
    End If
         
    ' Check requested Seletion Rule was found
    If Len(Trim(sSmartSelector)) > 0 Then
        If bSelectionRule Then
            Exit Sub
        ElseIf Not bSeachParents Then
            Exit Sub
        End If
    ElseIf Not bSeachParents Then
        ' Caller did not request search of All Selection Rules
        Exit Sub
    End If
                
    ' Continue search up SmartOccurrence Parents (if any)
    ' Need to check if given SmartOccurrence
    ' is a member of a CustomAssembly relationship
    ' it if is:
    '   then get the Parent of the SmartOccurrence/SmartItem
    '   to continue up the list
    GetParentSmartObject oSmartOccurrence, oParentObject
    
    If Not oParentObject Is Nothing Then
        oSelectorCollection.Add oSmartOccurrence
                
        GetSmartSelector oParentObject, _
                         sSmartSelector, oSelectorCollection, _
                         bSeachParents
        oSelectorCollection.Add oParentObject
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, Method).Number
End Sub
