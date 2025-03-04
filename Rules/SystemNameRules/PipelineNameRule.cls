VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PipelineNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2001, Intergraph Corporation.  All rights reserved.
'
'  Project: SystemsAndSpecs SystemNameRules
'  Class:   PipelineNameRule
'
'  Abstract: The file contains a sample implementation of a naming rule for Pipeline systems.
'
'  Author: Mark Mohon
'
'  History:
'
'  08/09/2001   Mohon   Name now uses string representation of FluidCode
'  02/22/2002   TM      Added checking for duplicate names.
'  06/20/2002   TM      Added a count value to the name so removed duplicate name checking.
'  07/10/2002   TM      Removed the counter added on 06/20, per requirement of our specification
'                         writer.
'
'******************************************************************

Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "PipelineNameRule: "
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"       ' define fixed-width number field for name

Dim m_oErrors As IJEditErrors

'Variables for localized strings
Private m_strDuplicateName As String
Private m_strHasChanged As String


Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
    InitStrings
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'********************************************************************
' Description:
' Creates a name for the object passed in. The name is based on the parents
' name and object name.It is something like this: "Base Name" + "Object Name" + Index.
' "Base Name" is derived from the Naming Parents and the Index is unique for the "Base Name"
' It is assumed that all Naming Parents and the Object implement IJPRodModelItem.
' The Naming Parents are added in AddNamingParents() of the same interface.
' Both these methods are called from naming rule semantic.
'
' Arguments:
'  oObject - Input.  Child object that needs to have the naming rule naming.
'  elements - Input.  Naming parents collection.
'  oActiveEntity - Input.  Naming rules active entity on which the NamingParentsString is stored.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    
    Dim oChildNamedItem As IJNamedItem
    Dim oParentNamedItem As IJNamedItem
    Dim strChildName As String
    Dim oPipeline As IJPipelineSystem
    Dim lFluidCode As Long
    Dim strFluidCode As String
    Dim oCodeListData As IJDCodeListMetaData
    
    Dim strNamedParentsString As String
    
    Dim jContext As IJContext
    Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    
    'For testing for duplicate names
    Dim boolIsDuplicate As Boolean
    Dim strOrigName As String
    Dim strNewName As String
    On Error GoTo IJNameRule_ComputeNameError
    
    'Get the middle context
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    
    'Verify inputs
    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If
    
    On Error Resume Next
    Set oParentNamedItem = elements.Item(1)
    Set oPipeline = oObject
    Set oChildNamedItem = oObject
    On Error GoTo IJNameRule_ComputeNameError
    
    If oPipeline Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If
    
    If Not oParentNamedItem Is Nothing Then
        'Start building the pipeline name by appending the name and the sequence number
        strChildName = oParentNamedItem.Name + oPipeline.SequenceNumber
        
        'Get the fluid code string
        lFluidCode = oPipeline.FluidCode
        
        On Error Resume Next
        Set oCodeListData = oPipeline
        On Error GoTo IJNameRule_ComputeNameError
        
        If Not oCodeListData Is Nothing Then
            strFluidCode = oCodeListData.ShortStringValue("FluidCode", lFluidCode)
            Set oCodeListData = Nothing
        Else
            strFluidCode = lFluidCode
        End If
                
        'Verify that we haven't already named this pipeline.
        strNamedParentsString = oActiveEntity.NamingParentsString
        If (strChildName + strFluidCode) <> strNamedParentsString Then
        
            oActiveEntity.NamingParentsString = strChildName + strFluidCode
        
            strChildName = strChildName + "-" + strFluidCode
        
            'Set the name of the pipeline
            oChildNamedItem.Name = strChildName
        
            'Check for duplicate names.
            TestForDuplicateName oObject, boolIsDuplicate, strOrigName, strNewName
            'TM - 4/23/04 - we used to raise duplicate name errors, but that confused
            ' the users.  Simply rename the system and don't raise the error.
            'If boolIsDuplicate Then
            '    m_oErrors.Add ernDuplicateName, _
            '        "IJNameRule_ComputeName", _
            '        m_strDuplicateName & strOrigName & m_strHasChanged & strNewName, _
            '            "NAMING", "", 0, "", "", "", 0
            '    On Error Resume Next
            '    Err.Raise E_FAIL
            'End If
            
        End If
        
    End If
    
    'Clean up
    Set jContext = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    
    Set oChildNamedItem = Nothing
    Set oParentNamedItem = Nothing
    Set oPipeline = Nothing
    
Exit Sub

IJNameRule_ComputeNameError:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' All the Naming Parents that need to participate in an objects naming are added here to the
' IJElements collection. The parents added here are used in computing the name of the object in
' ComputeName() of the same interface. Both these methods are called from naming rule semantic.
'
' Arguments:
'  oEntity - Input.  Child object that needs to have the naming rule naming.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo IJNameRule_GetNamingParentsError

    Dim oSysParent As IJSystem
    Dim oSysChild As IJSystemChild
    Dim oUnitSys As IJUnitSystem
    Dim oLastChild As IJSystemChild
    Dim oPipeline As IJPipelineSystem

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
    Set oSysChild = oEntity
    
    'For our purposes, the parent will be the closest unit system above the pipeline
    ' in the pipeline's branch of the system hierarchy.  (Remember, a generic system
    ' is a unit system.)  If no unit system is found in the pipeline's branch of the
    ' system hierarchy, the system closest to the root object in this pipeline's
    ' branch of the hierarchy is the parent.  Travel up the hierarchy until either
    ' a unit system is found or the root object is reached.
    Do While IJNameRule_GetNamingParents.Count = 0
        On Error Resume Next
        Set oSysParent = Nothing
        Set oSysParent = oSysChild.GetParent
        Set oUnitSys = oSysParent
        Set oLastChild = Nothing
        Set oLastChild = oSysChild
        Set oSysChild = Nothing
        Set oSysChild = oSysParent
        Set oPipeline = Nothing
        Set oPipeline = oLastChild
        On Error GoTo IJNameRule_GetNamingParentsError
        
        If Not oUnitSys Is Nothing Then
            Call IJNameRule_GetNamingParents.Add(oSysParent)
            Set oUnitSys = Nothing
        ElseIf oSysChild Is Nothing Then
            If oPipeline Is Nothing Then
                Call IJNameRule_GetNamingParents.Add(oLastChild)
            Else
                Call IJNameRule_GetNamingParents.Add(oSysParent)
            End If
        End If
    Loop 'While
    
    Set oSysParent = Nothing
    Set oSysChild = Nothing
    Set oUnitSys = Nothing
    Set oLastChild = Nothing
Exit Function

IJNameRule_GetNamingParentsError:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function


Private Sub InitStrings()

    Dim oLocalizer As IJLocalizer
    Set oLocalizer = New IMSLocalizer.Localizer
    
    If Not oLocalizer Is Nothing Then
        If Not IsInDebugMode() Then
            oLocalizer.Initialize App.Path + "\" + App.EXEName
        Else
            oLocalizer.Initialize "M:\SystemsAndSpecs\Client\Bin\" + App.EXEName
        End If
        
        m_strDuplicateName = oLocalizer.GetString(IDS_DUPLICATENAME, "Duplicate system name:  ")
        m_strHasChanged = oLocalizer.GetString(IDS_HASCHANGED, " has been changed to ")
    End If
    
    Exit Sub
    
End Sub

