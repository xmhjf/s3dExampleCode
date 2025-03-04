VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserDefinedNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002-2009, Intergraph Corporation.  All rights reserved.
'
'  Project: SystemNamingRules
'
'  Abstract: The file contains an implementation of a naming rule for SystemsAndSpecs
'
'  Author: T. Merchant  (Following the pattern of work by Bathina Balakrishna Reddy)
'
' 01/30/02      T. Merchant                 Created
' 05/06/09      MJF                         TR 164444  GetNamingParents should return Nothing
'
'******************************************************************

Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "UserDefinedNameRule: "
Private Const DEFAULTBASE = "Defaultname"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"       ' define fixed-width number field for name
Private Const PARTROLE = "part"

Dim m_oErrors As IJEditErrors

'Variables for localized strings
Private m_strDuplicateName As String
Private m_strHasChanged As String

Private m_strBlankName As String
Private m_strSystem As String
Private m_strNew As String
Private m_strGeneric As String
Private m_strConduit As String
Private m_strDucting As String
Private m_strEquipment As String
Private m_strPipeline As String
Private m_strPiping As String
Private m_strStructural As String
Private m_strUnit As String
Private m_strArea As String


Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
    InitStrings
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'********************************************************************
' Description:
' This is where the name is usually generated.  In our case, we only want
' to verify that the new name is not shared by any sibling systems.
'
' Arguments:
'  oObject - Input.  Child object that needs to have the naming rule naming.
'  oParents - Input.  Naming parents collection.
'  oActiveEntity - Input.  Naming rules active entity on which the NamingParentsString is stored.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal oParents As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo IJNameRule_ComputeNameError
    
    'For testing for duplicate names
    Dim boolIsDuplicate As Boolean
    Dim strOrigName As String
    Dim strNewName As String
        
    'Check for blank names.
    If IsNameBlank(oObject, strNewName) Then
        m_oErrors.Add ernBlankName, _
            "IJNameRule_ComputeName", _
            m_strBlankName & strNewName, _
                "NAMING", "", 0, "", "", "", 0
        On Error Resume Next
        Err.Raise E_FAIL
    End If
        
    'Check for duplicate names.
    TestForDuplicateName oObject, boolIsDuplicate, strOrigName, strNewName
    If boolIsDuplicate Then
        m_oErrors.Add ernDuplicateName, _
            "IJNameRule_ComputeName", _
            m_strDuplicateName & strOrigName & m_strHasChanged & strNewName, _
                "NAMING", "", 0, "", "", "", 0
        On Error Resume Next
        Err.Raise E_FAIL
    End If
        
    Exit Sub

IJNameRule_ComputeNameError:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' Returns naming parents of the object.  In this case, there are no naming parents so return Nothing
'
' Arguments:
'  oEntity - Input.  Child object that needs to have the naming rule naming.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    
    On Error GoTo label
    Set IJNameRule_GetNamingParents = Nothing
    Exit Function
    

label:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' Initialize localized strings
'
' Arguments:
'  None.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitStrings()

    Dim oLocalizer As IJLocalizer
    
    On Error GoTo Error
    
    Set oLocalizer = New IMSLocalizer.Localizer
    
    If Not oLocalizer Is Nothing Then
        If Not IsInDebugMode() Then
            oLocalizer.Initialize App.Path + "\" + App.EXEName
        Else
            oLocalizer.Initialize "M:\SystemsAndSpecs\Client\Bin\" + App.EXEName
        End If
        
        m_strDuplicateName = oLocalizer.GetString(IDS_DUPLICATENAME, "Duplicate system name:  ")
        m_strHasChanged = oLocalizer.GetString(IDS_HASCHANGED, " has been changed to ")
        m_strBlankName = oLocalizer.GetString(IDS_BLANKNAME, _
            "Blank system names are not allowed.  Name has been changed to:  ")
        m_strSystem = oLocalizer.GetString(IDS_SYSTEM, "System")
        m_strNew = oLocalizer.GetString(IDS_NEW, "New")
        m_strGeneric = oLocalizer.GetString(IDS_GENERIC, "Generic")
        m_strConduit = oLocalizer.GetString(IDS_CONDUIT, "Conduit")
        m_strDucting = oLocalizer.GetString(IDS_DUCTING, "Ducting")
        m_strEquipment = oLocalizer.GetString(IDS_MACHINERY, "Equipment")
        m_strPipeline = oLocalizer.GetString(IDS_PIPELINE, "Pipeline")
        m_strPiping = oLocalizer.GetString(IDS_PIPING, "Piping")
        m_strStructural = oLocalizer.GetString(IDS_STRUCTURAL, "Structural")
        m_strUnit = oLocalizer.GetString(IDS_UNIT, "Unit")
        m_strArea = oLocalizer.GetString(IDS_AREA, "Area")
    End If
    
    Exit Sub

Error:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "InitStrings", Err.Description
    Err.Raise E_FAIL
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' Check if the name of the system is blank.  If so, generate a new name.
'
' Arguments:
'  oObject - Input.  System object we are checking.
'  strGeneratedname - Output.  Generated new name for the system.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsNameBlank(oObject As Object, strGeneratedName As String) As String
    Dim oParent As IJDesignParent
    Dim oChild As IJDesignChild
    Dim oObjectName As IJNamedItem
    Dim iNdx As Long
    Dim bSystemExists As Boolean
    Dim bFoundMatch As Boolean
    Dim strSystemType As String
    Dim oChildrenCollection As New JObjectCollection
    Dim oChildren As IJDObjectCollection
    Dim oChildName As IJNamedItem
    
    IsNameBlank = False
    strGeneratedName = ""
    
    Set oObjectName = oObject
    If oObjectName Is Nothing Then Exit Function
    
    If oObjectName.Name = "" Then
        
        IsNameBlank = True
        Set oChild = oObject
        If Not oChild Is Nothing Then
            Set oParent = oChild.GetParent
            Set oChild = Nothing
        End If
        
        If Not oParent Is Nothing Then
            
            'Create a starting name for the system
            oParent.GetCount iNdx, "SystemChildren"
            GetSystemType oObject, strSystemType
    
            strGeneratedName = m_strNew & strSystemType & m_strSystem & iNdx
            
            bFoundMatch = False
            bSystemExists = False
    
            Set oChildren = oChildrenCollection
            oParent.GetChildren oChildren, "SystemChildren"
    
            Do While Not bSystemExists
                For Each oChildName In oChildren
                    If oChildName.Name = strGeneratedName Then
                        iNdx = iNdx + 1
                        strGeneratedName = m_strNew & strSystemType & m_strSystem & iNdx
                        bFoundMatch = True
                        Exit For
                    End If
                Next
    
                If bFoundMatch = True Then
                    bSystemExists = False
                    bFoundMatch = False
                Else
                    bSystemExists = True
                End If
            Loop
            oObjectName.Name = strGeneratedName
        End If
    End If
    Exit Function
    
Error:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IsNameBlank", Err.Description
    Err.Raise E_FAIL
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' Determine what kind of system we are working with.
'
' Arguments:
'  oObject - Input.  System object we are checking.
'  strSystemType - Output.  A string that contains the system type description.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetSystemType(oSystem As Object, strSystemType As String)
    
    On Error GoTo Error
    
    strSystemType = ""
    If TypeOf oSystem Is IJGenericSystem Then
        strSystemType = m_strGeneric
    ElseIf TypeOf oSystem Is IJConduitSystem Then
        strSystemType = m_strConduit
    ElseIf TypeOf oSystem Is IJDuctingSystem Then
        strSystemType = m_strDucting
    ElseIf TypeOf oSystem Is IJMachinerySystem Then
        strSystemType = m_strEquipment
    ElseIf TypeOf oSystem Is IJPipelineSystem Then
        strSystemType = m_strPipeline
    ElseIf TypeOf oSystem Is IJPipingSystem Then
        strSystemType = m_strPiping
    ElseIf TypeOf oSystem Is IJStructuralSystem Then
        strSystemType = m_strStructural
    ElseIf TypeOf oSystem Is IJUnitSystem Then
        strSystemType = m_strUnit
    ElseIf TypeOf oSystem Is IJAreaSystem Then
        strSystemType = m_strArea
    End If
    Exit Sub
    
Error:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetSystemType", Err.Description
    Err.Raise E_FAIL
End Sub
