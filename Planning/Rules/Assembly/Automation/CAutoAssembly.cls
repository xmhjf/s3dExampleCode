VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAutoAssembly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 1998-2002 Intergraph Corporation.  All Rights Reserved.
'
' File: SHIAutoRule.cls
'
' Author: Kamrooz k. Hormoozi
'
' Abstract: Implementats the rule
'
' Description:
' The default rule is implemented and functions which uses to walk thorugh the system and
' Assembly hierarchy are also implemented here
'
'  History
'   May 02 2002           oss\kaho   Created.
'   Jan 17 2017           Gopi M     DI-CP-308270 Assembly properties customization.
'********************************************************************
Option Explicit

Private Const MODULE As String = "PlnAutoAssemblyCmd.CAutoAssembly"

' ***************************************************************************
'
' Function
'   GetDefaultNamingRule()
'
' Abstract
'   Returns the default Korean naming rule for the given object type
'
' ***************************************************************************

Private Function GetDefaultkoreanNamingRule(ByVal ObjectType As String) As GSCADGenericNamingRulesFacelets.NameRuleHolderFacelets
    Const METHOD As String = "GetDefaultkoreanNamingRule"
    On Error GoTo ErrorHelper
    
    Dim RuleHelper As IJDNamingRulesHelper
    Set RuleHelper = New GSCADNameRuleHlpr.NamingRulesHelper
    ' Get collection of available rules
    Dim Rules As IJElements
    On Error Resume Next
    RuleHelper.GetEntityNamingRulesGivenName ObjectType, Rules
    ' RuleHelper.GetEntityNamingRulesGivenProgID ObjectType, Rules
    ' On Error GoTo ErrorHelper
    
    ' No rules collection? ..then return no rule
    If Rules Is Nothing Then
      
        Set GetDefaultkoreanNamingRule = Nothing
        
        Exit Function
    End If
    
    ' No rules? ..then return no rule
    If Rules.Count = 0 Then
        Set GetDefaultkoreanNamingRule = Nothing
        Exit Function
    End If
    
    ' Return the first (default) naming rule in
    ' It works as design. The desired naming rule sould be the first rule
    ' in the list of the naming rule for assembly (check reference DataBase)
    Set GetDefaultkoreanNamingRule = Rules(1)
    
    Exit Function
ErrorHelper:
    MsgBox Err.Description
End Function

Public Sub SetDefaultkoreanNamingRule(ByVal ObjType As String, ByVal oAssembly As IJAssembly)
    Const METHOD As String = "SetDefaultkoreanNamingRule"
    On Error GoTo ErrorHelper
    
    Dim DefaultRule As GSCADGenericNamingRulesFacelets.NameRuleHolderFacelets
    Set DefaultRule = GetDefaultkoreanNamingRule(ObjType)

    If Not DefaultRule Is Nothing Then
        Dim RuleHelper As IJDNamingRulesHelper
        Set RuleHelper = New GSCADNameRuleHlpr.NamingRulesHelper
        
        Dim Semantic As GSCADGenNameRuleAE.IJNameRuleAE
        RuleHelper.AddNamingRelations oAssembly, DefaultRule, Semantic
    End If
    
    Exit Sub
ErrorHelper:
    MsgBox Err.Description
End Sub

' ***************************************************************************
'
' Function
'   CreateAssembly()
'
' Abstract
'   Creates a new sub-assembly of the currently selected assembly
'
' Description
'   The function instantiates an Assembly Factory, and asks it to create a
'   new Assembly Entity.
'
' Notes
'   This method does not handle errors by itself.  Errors are passed to the
'   calling function.
'
' ***************************************************************************

Public Function CreateAssembly( _
        ByVal oParent As GSCADAssembly.IJAssembly, _
        ByVal oChildren As IJElements) As GSCADAssembly.IJAssembly
    Const METHOD As String = "CreateAssembly"
    On Error GoTo ErrorHandler
   
    ' Create assembly factory
    Dim oFactory As GSCADAssembly.IJAssemblyFactory
    Set oFactory = New GSCADAssembly.AssemblyFactory
    
    ' Create new assembly
    Dim oNewAssembly As GSCADAssembly.IJAssembly
    Set oNewAssembly = oFactory.CreateAssembly(GetActiveConnection.GetResourceManager(GetActiveConnectionName), oParent)
    
    ' Set Assembly Contents
    Dim oChild As IJAssemblyChild
    For Each oChild In oChildren
        oNewAssembly.AddChild oChild

    Next
    
    ' Return Created Assembly
    Set CreateAssembly = oNewAssembly
    
    ' Clean Up
    Set oChild = Nothing
    Set oNewAssembly = Nothing
    Set oFactory = Nothing

    Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

Public Sub FilterOutSytems(ByVal oElements As IJElements, oFElements As IJElements, oSElements As IJElements, oTElements As IJElements)
   Const METHOD As String = "FilterOutSystem"
   On Error GoTo ErrorHandler
   
   Set oFElements = New JObjectCollection
   Set oSElements = New JObjectCollection
   Set oTElements = New JObjectCollection
   
   Dim oCollectSet As IJElements
   Set oCollectSet = New JObjectCollection
   
   Dim oCollectList As IJElements
   Set oCollectList = New JObjectCollection
   
   Dim oCollection As IJElements
   Set oCollection = New JObjectCollection
   
   Dim oObject As Object
   Dim index As Integer
   
   For index = 1 To oElements.Count
        Set oObject = oElements.Item(index)
        If TypeOf oObject Is IJPlatePart Or TypeOf oObject Is IJPlateSystem Then
             oCollectSet.Add oObject
        ElseIf TypeOf oObject Is IJBeamPart Or TypeOf oObject Is IJBeamSystem Then
             oCollectList.Add oObject
        ElseIf TypeOf oObject Is IJStiffenerPart Or TypeOf oObject Is IJStiffenerSystem Then
             oCollection.Add oObject
        End If
    Next index
    
    AddElements oFElements, oCollectSet
    oCollectSet.Clear
    Set oCollectSet = Nothing
    AddElements oSElements, oCollectList
    oCollectList.Clear
    Set oCollectSet = Nothing
    AddElements oTElements, oCollection
    oCollection.Clear
    Set oCollection = Nothing
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub


' ***************************************************************************
'
' Function
'   SetWorkAndEquipment(oNewAssembly As GSCADAssembly.IJAssembly)
'
' Abstract
'   [in] the function has the newassembly as input
'   [out] Nothing
'   The function find  all workcenters and take the default workcenter which is the first
'   workcenter in the list and put it in the SetWorkCenter function. On basis on the workcenter
'   the function find the Equipment where the first is default BuildEquipment, the second is default
'   WeldEquipment and the third is default ExitEquipment
'
' ***************************************************************************
Public Sub SetWorkAndEquipment(oNewAssembly As GSCADAssembly.IJAssembly)
    Const METHOD = "SetWorkAndEquipment"
    On Error GoTo ErrorHandler
    
    Dim ListOfWorkcenters() As String
    Dim ListOfEquipments() As String
    Dim objWorkcenterQuery As IJSRDWorkcenterQuery
    Dim objSRDQuery As IJSRDQuery
    Dim objPlnAssociateCatalog As IJPlnAssociateCatalog
    Set objPlnAssociateCatalog = oNewAssembly
    Dim oWorkCenterName As String
    
    ' Create the non-persistent SRDServices.SRDQuery to get the Query Object for Workcenter
    Set objSRDQuery = New SRDQuery
    
    ' Fetch the SRDWorkcenterQuery (persisted in the catalog db) to perform SQL queries on the catalog database
     Set objWorkcenterQuery = objSRDQuery.GetWorkcenterQuery()
    
    'get list of all workcenter
    objWorkcenterQuery.GetAllWorkcenters ListOfWorkcenters
    
    'check about the workcenter is bulkload
    If UBound(ListOfWorkcenters) - LBound(ListOfWorkcenters) >= 0 Then
    
        'add the default workcenter
        oWorkCenterName = ListOfWorkcenters(0)
        
        If objPlnAssociateCatalog.WorkcenterName = vbNullString Then
            'save the default value of the workcenter
            objPlnAssociateCatalog.SetWorkcenter oWorkCenterName
        Else
            oWorkCenterName = objPlnAssociateCatalog.WorkcenterName
        End If
        
        'get list of equipment on basis of the default workcenter
        objWorkcenterQuery.GetEquipmentInWorkcenter oWorkCenterName, ListOfEquipments

        If UBound(ListOfEquipments) - LBound(ListOfEquipments) >= 3 Then
            'save the default value of the exitequiment
            objPlnAssociateCatalog.SetExitEquipment ListOfEquipments(0), oWorkCenterName
            
            'save the default value of the weldequiment
            objPlnAssociateCatalog.SetWeldEquipment ListOfEquipments(2), oWorkCenterName
            
            'save the default value of the buildequipment
            objPlnAssociateCatalog.SetBuildEquipment ListOfEquipments(3), oWorkCenterName
            
        End If
    End If
Cleanup:
    Set objSRDQuery = Nothing
    Set objWorkcenterQuery = Nothing
Exit Sub
ErrorHandler:
    MsgBox Err.Description
    GoTo Cleanup
End Sub


Private Function GetActiveConnection() As IJDAccessMiddle
    Const METHOD = "GetActiveConnection"
    On Error GoTo ErrorHandler
    
    Dim oCmnAppGenericUtil As IJDCmnAppGenericUtil
    Set oCmnAppGenericUtil = New CmnAppGenericUtil
    
    oCmnAppGenericUtil.GetActiveConnection GetActiveConnection

    Set oCmnAppGenericUtil = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function GetActiveConnectionName() As String

    Dim jContext As IJContext
    Dim oDBTypeConfiguration As IJDBTypeConfiguration
    
    'Get the middle context
    Set jContext = GetJContext()
    
    'Get IJDBTypeConfiguration from the Context.
    Set oDBTypeConfiguration = jContext.GetService("DBTypeConfiguration")
    
    'Get the Model DataBase ID given the database type
    GetActiveConnectionName = oDBTypeConfiguration.get_DataBaseFromDBType("Model")
    
    Set jContext = Nothing
    Set oDBTypeConfiguration = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

