VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarkFldrNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: GSCADStrMfgNameRules
'
'  Abstract: The file contains an implementation of the default naming rule
'            for the MarkingFolder object in StructManufacturing UE.
'
'  History:'

'     Venkat (vkukutla)  2006/08/21     Created for giving a default Name Rule for Marking Folder-DI:101206
'
'***************************************************************************

Option Explicit

Implements IJNameRule

Const vbInvalidArg = &H80070057

Private Const MODULE = "MarkingFolderNameRule: "
Private Const DEFAULTBASE = "MKFR"

'*********************************************************************************************
' Description:
'   Creates a name for the object passed in. The name is based on the parents
'   name. It is something like this: "Base Name" + "Parent Name" + Index.
'   It is assumed that all Naming Parents and the Object implement IJNamedItem.
'   The Naming Parents are added in AddNamingParents() of the same interface.
'   Both these methods are called from the naming rule semantic.
'
' Notes: This default Naming Rule was implemented based on the following condition.
'
'   Reflection of Model Part (Reference Plane, Plate, Profile, etc ...) with a prefix to
'   indicate the marking object.
'   Ex: "MKFR_B2-A23-BL2".
'
'*********************************************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler

    Dim oNameCounter As IJNameCounter
    Dim oChildNamed As IJNamedItem
    Dim oParent As Object
    Dim oParentNamed As IJNamedItem

    Dim strChildNewName As String
    
    Dim strParentName As String
    Dim strNamedParentsString As String
    
    Dim strLocation As String
    strLocation = vbNullString
    
    Dim nCount As Long
    Dim oActiveConnection As IJDAccessMiddle

    Set oActiveConnection = GetActiveConnection
    
    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, MODULE, METHOD
    End If

    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    
    'Get the Name count and location ID
    nCount = oNameCounter.GetCountEx(oActiveConnection.GetResourceManager(GetActiveConnectionName), _
                         "GSCADStrMfgNamingRule_MarkingDefaultGenerator", strLocation)
'                        "MfgMarkingFolder", strLocation)

    Set oChildNamed = oObject

        If elements.Count > 0 Then
            'It is assumed that only one parent object has been passed in.
            Set oParent = elements.Item(1)
            ' the object which supports IJNamedItem is only the valid one.
            If TypeOf oParent Is IJNamedItem Then
                Set oParentNamed = oParent
                strParentName = oParentNamed.Name

                If strLocation <> vbNullString Then
                    strChildNewName = DEFAULTBASE + "_" + strParentName + "-" + strLocation
                Else
                    strChildNewName = DEFAULTBASE + "_" + strParentName
                End If
                
                strNamedParentsString = strParentName
                
            End If
        End If
        
'       'If strChildName is empty then apply the default.Worst case.
        If (Len(strChildNewName) = 0) Then
'            'Default name will be (ex; "MKFR_NoParent000001", "MKFR_NoParent000001", ..., "MKFR_NoParent999999")

            If strLocation <> vbNullString Then
                strChildNewName = DEFAULTBASE + "_NoParent" + "-" + strLocation
            Else
                 strChildNewName = DEFAULTBASE + "_NoParent"
            End If '
            strNamedParentsString = "MarkingFolderDefaultName"
        End If
    
        'Set the NamingParentsString property of the Active Entity.
        If oActiveEntity.NamingParentsString <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strNamedParentsString
        End If
        
        Dim strChildCurrentName As String
        Dim strChildNameTruncated As String
        'Get current name of the object
        strChildCurrentName = oChildNamed.Name
        On Error Resume Next   'In case if it is not having a name
        'Get name of the object without Count.So truncate count part of the string
        strChildNameTruncated = Left(strChildCurrentName, (InStrRev(strChildCurrentName, "-") - 1))
        On Error Resume Next   'In case for User defined Name, we may not get the truncated string.
        
        'Apply the Default Name to Object
        If strChildNameTruncated <> strChildNewName Then
            oChildNamed.Name = strChildNewName + "-" + Format(nCount, "") 'Object having User defined Name OR without Name comes here and default name is assigned here.
        Else
         oChildNamed.Name = strChildCurrentName
        End If

    
    Set oActiveConnection = Nothing
    Set oNameCounter = Nothing
    Set oChildNamed = Nothing
    Set oParent = Nothing
    Set oParentNamed = Nothing

Exit Sub

ErrorHandler:

    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 4002, , "RULES")
End Sub

'****************************************************************************************************
'Description
'   All the Naming Parents that need to participate in an objects naming are added here to the
'   IJElements collection. The parents added here are used in computing the name of the object in
'   ComputeName() of the same interface. Both these methods are called from the naming rule semantic.
'****************************************************************************************************
Private Function IJNameRule_GetNamingParents(ByVal oMfgMarking As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    ' Need to be implemented if it is clarified how to get the parent object
    ' (i.e. Model Part such as Reference Plane, Plate, Profile, etc ...)
    ' from the given oMfgMarking object.


    Dim oMfgChild As IJMfgChild
    Set oMfgChild = oMfgMarking

    Dim oMfgParent As IJMfgParent
    Set oMfgParent = oMfgChild.GetParent
     
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    If Not (oMfgParent Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oMfgParent)
    End If

    Set oMfgParent = Nothing
    Set oMfgChild = Nothing

Exit Function

ErrorHandler:

    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 4002, , "RULES")
End Function
 