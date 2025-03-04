VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssyMarginNameRule"
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
'            for the Assembly Margin object in StructManufacturing UE.
'
'  History:
'     Yeun-gyu Kim  01/24/00    Initial creation.
'     Marcel Veldhuizen 2004/04/22  Included correct error handling
'     Venkat (vkukutla)  2006/07/05  Modified GetCount into GetCountEx and modified the Naming Rule
'                                    to include work share location ID in order to avoid replication problems.CR:91673
'     Venkat (vkukutla)  2006/07/24  Existing error handler is replaced with Struct Manufacturing specific Error logger.CR:91673
'
'***************************************************************************

Option Explicit

Implements IJNameRule

Const vbInvalidArg = &H80070057

Private Const MODULE = "AssemblyMarginNameRule: "
Private Const DEFAULTBASE = "Assembly Margin"

'   Creates a name for the object passed in.
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler

    Dim oActiveConnection As IJDAccessMiddle
    Dim oChildNamed       As IJNamedItem
    
    Set oChildNamed = oObject

    Set oActiveConnection = GetActiveConnection

    Dim oNameCounter        As IJNameCounter

    Dim strNamedParentsString As String
    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, MODULE, METHOD
    End If

    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    
    Dim BOStringName As String
    If TypeOf oObject Is IJAssyMarginParent Then
        BOStringName = Trim("Assembly margin parent")
    ElseIf TypeOf oObject Is IJAssyMarginChild Then
        BOStringName = Trim("Assembly margin child")
    End If
    
    Dim strLocation As String
    strLocation = vbNullString
    Dim nCount As Long
    'Get name count and location ID
    nCount = oNameCounter.GetCountEx(oActiveConnection.GetResourceManager(GetActiveConnectionName), BOStringName, strLocation)

    If Not strLocation = vbNullString Then
       strNamedParentsString = BOStringName + " " + "-" + strLocation + "-" + Format(nCount, "########")
    Else
       strNamedParentsString = BOStringName + " " + "-" + Format(nCount, "########")
    End If

    'Set the NamingParentsString property of the Active Entity.
    If oActiveEntity.NamingParentsString <> strNamedParentsString Then
        oActiveEntity.NamingParentsString = strNamedParentsString
    End If
    
    If oChildNamed.Name <> strNamedParentsString Then
        oChildNamed.Name = strNamedParentsString
    End If
    
    Set oActiveConnection = Nothing
    Set oChildNamed = Nothing
    Set oNameCounter = Nothing

Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 6002, , "RULES")
End Sub

'****************************************************************************************************
'Description
'   All the Naming Parents that need to participate in an objects naming are added here to the
'   IJElements collection. The parents added here are used in computing the name of the object in
'   ComputeName() of the same interface. Both these methods are called from the naming rule semantic.
'****************************************************************************************************
Private Function IJNameRule_GetNamingParents(ByVal oAssyMargin As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oPlatePart As GSCADAsmHlpers.IJAssemblyChild
    Dim oIJAssemblyChild As GSCADAsmHlpers.IJAssemblyChild
    Dim oMfgChild As MfgChildHelperLib.IJMfgChild
    Dim oParent As Object

    'Retrieve the PlatePart object from the given oMfgTemplate object.
    Set oMfgChild = oAssyMargin
'    On Error Resume Next

    Set oPlatePart = oMfgChild.GetParent
    On Error GoTo ErrorHandler
    
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    Set oMfgChild = Nothing
    If oPlatePart Is Nothing Then
        GoTo ErrorHandler
    End If

    'Add Block object as the first parent object into the collection.
    Set oIJAssemblyChild = oPlatePart
    Do While True
        On Error Resume Next
        Set oParent = oIJAssemblyChild.Parent
        On Error GoTo ErrorHandler
        
        If oParent Is Nothing Then Exit Do                      'There was no block found, terminate loop.
        If Not TypeOf oParent Is IJAssemblyChild Then Exit Do   'There was no block found, terminate loop.
        
        If TypeOf oParent Is IJBlockItem Then
            'Add the found block as the naming parent.
            Call IJNameRule_GetNamingParents.Add(oParent)
            Exit Do
        End If
        Set oIJAssemblyChild = Nothing
        Set oIJAssemblyChild = oParent
        Set oParent = Nothing
    Loop

    'Add the PlatePart object as the naming parent.
    Call IJNameRule_GetNamingParents.Add(oPlatePart)

    Set oIJAssemblyChild = Nothing
    Set oParent = Nothing
    Set oPlatePart = Nothing
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 6002, , "RULES")
End Function
 
