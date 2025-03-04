VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlnJointFolderNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project: NameRules
'
'  Abstract: Naming rule for joint folders
'
'  Author: Ole von Mortensen    oss/olmo
'
'  Changes:
'  10/27/04  CRS  Recognize a set of valid folder names.  Check for
'                 existence of a parent folder.
'
'******************************************************************

'********************************************************************
' GENERIC ABOUT FOLDERNAMES
' Folders will appear in the following hierarchy:
' Assembly
'   |-Part, seq 1
'   |-Part, seq 2
'   |-Part, seq 3
'   |-Part, UnSeq
'   |-Welds (a holder for all welds)
'      |-WeldsSeq2 (welds involved in attaching part 2 to part 1)
'      |-WeldsSeq3 (welds involved in attaching part 3 to part 2)
'      |-UnSeqWelds (welds belonging to unsequenced parts)
'      |-UnAssignWelds (welds belonging to unassigned parts)
'      |-DelayedWelds (welds which belongs to assemblies further down, but has been postponed)
'
' There are 3 static holder names (Welds, UnSeqWelds, DelayedWelds).
' and 1 dynamic name (WeldsSeq3), which must take the sequence as input.
'********************************************************************

Option Explicit
Implements IJNameRule
                                                        
Private Const Module = "JointFolderNameRule: "
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000000"    ' define fixed-width number field for name

Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"

Private m_oErrors As IJEditErrors       ' To collect and propagate the errors.
Private m_oError As IJEditError         ' The error to raise.

Private Sub Class_Initialize()
   Set m_oErrors = CreateObject(ERRORPROGID)
End Sub

Private Sub Class_Terminate()
   Set m_oErrors = Nothing
   Set m_oError = Nothing
End Sub

'********************************************************************
' Method
'   IJNameRule_ComputeName
'
' Description:
'   Creates a name for the PlnJointFolder passed in. The name is based on the string "H" and
'   an Index. The Index is unique for the Folder. It is assumed that all Naming Parents
'   and the Object implement IJNamedItem. The Naming Parents are added in AddNamingParents()
'   of the same interface. Both these methods are called from naming rule semantic.
'
' Parameters:
'   [in] pEntity        Object to be named.
'   [in] pParents       Collection of parents of the object to be named.
'   [in] pActiveEntity  Naming rules active entity on which the NamingParentsString is stored.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, _
                                   ByVal elements As IJElements, _
                                   ByVal oActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler
    
    ' get the named item
    Dim strName As String
    Dim oNamedItem As IJNamedItem
    Set oNamedItem = oObject
    
    ' first check for type of folder
    Dim oPlnJointFolder As IJPlnJointFolder
    Set oPlnJointFolder = oObject
    
    Dim lFolderType As Long
    lFolderType = oPlnJointFolder.FolderType
    
    Select Case lFolderType
        Case PLN_FolderType_Root
            strName = "Welds"
        Case PLN_FolderType_Postponed
            strName = "Postponed"
        Case PLN_FolderType_Unsequenced
            strName = "Unsequenced"
        Case PLN_FolderType_Unassigned
            strName = "Unassigned"
        Case PLN_FolderType_Subsequent
            strName = "Subsequent"
        Case PLN_FolderType_Joining
            GoTo CalculateName
        Case Else
            'this is for custom planning joint folder.it will get the display name from DB and set it  name to the folder
            strName = GetDisplayString(oObject, "FolderType", lFolderType)
    End Select
    
    oNamedItem.Name = strName
    GoTo CleanUp

CalculateName:
    Dim oNaming As Object
    For Each oNaming In elements
        Dim oNamingParent As IJNamedItem
        Set oNamingParent = oNaming
            
        strName = oNamingParent.Name
        oNamedItem.Name = "Joining " + strName
        Set oNamingParent = Nothing
    Next oNaming
    
CleanUp:
    Set oNamedItem = Nothing
    Set oPlnJointFolder = Nothing
    
    Exit Sub
    
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
    GoTo CleanUp
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method
'   IJNameRule_GetNamingParents
'
'Description
'   This description is not implemented yet, but supposed to be at a later point. Name of
'   the folder should be calculated the following way
'   1. Get the part objects the which the welds belong to
'   2. Find the lowest sequence number of the 2 parts
'   3. Calc the Name by applying W-seqnumber-Random number, e.g. W-5-00000057
'   4. In case of sequence change in one of the parts, name must be updated
'   5. In case all planning joints are removed from folder, folder should be removed
'   6. In case both parts are unsequenced all planning joints should be moved to UnSequencedWelds folder
'   7. In some way the naming rule should take IJSequence as input
'
'   As temp solution we just create a standard name not involving any parents, but this method
'   should return the 2 parts involved in naming
'
' Parameters:
'   [in] pEntity                        Child object that needs to have the naming rule naming.
'   [out] IJNameRule_GetNamingParents,  JObjectCollection with naming parents
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler
    
    ' create the collection
    Dim oParents As IJElements
    Set oParents = New IMSCoreCollections.JObjectCollection
    
    ' check if the folder has been related to any AssemblyChild
    Dim oPlnJointFolder As IJPlnJointFolder
    Set oPlnJointFolder = oEntity
    
    Dim oAssyChild As IJAssemblyChild
    Set oAssyChild = oPlnJointFolder.GetRelatedAssyChild
    
    If Not oAssyChild Is Nothing Then
        oParents.Add oAssyChild
    End If
    
    Set IJNameRule_GetNamingParents = oParents

CleanUp:
    Set oParents = Nothing
    Set oAssyChild = Nothing
    Set oPlnJointFolder = Nothing
    
    Exit Function

ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
    GoTo CleanUp
End Function
 
