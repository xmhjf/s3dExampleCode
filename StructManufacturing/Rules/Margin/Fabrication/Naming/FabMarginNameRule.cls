VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FabMarginNameRule"
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
'            for the FabricationMargin object in StructManufacturing UE.
'
'  History:
'     Yeun-gyu Kim  01/06/00    Initial creation.
'     Yeun-gyu Kim  03/06/00    Updated to set the block as the 1st naming parent correctly
'                               and the related codes.
'     Yeun-gyu Kim  04/20/00    Updated IJNameRule_ComputeName method to get the Block object
'                               correctly if given.
'     Yeun-gyu Kim  07/27/01    Updated IJNameRule_GetNamingParents method to get the parent object
'                               by using IJAssembly ----AssemblyHierarchy----IJAssemblyChild relationship.
'     ByungSam Lee  11/16/01    Changed the Relationship btw PlatePart and Manufacturing Object to
'                               IJMfgParent ----------StrMfgHierarchy - --------IJStrMfgChild
'    Marcel Veldhuizen 2004/04/22 Included correct error handling

 '     Thakur        2005/04/26 Modified ComputeName method  to avoid duplicate name of Template
 '                                         as more than oe template is supported for a plate.
 
'     Venkat (vkukutla)  2006/07/05  Modified GetCount into GetCountEx and modified the Naming Rule
'                                    to include work share location ID in order to avoid replication problems.CR:91673
'     Venkat (vkukutla)  2006/07/24  Existing error handler is replaced with Struct Manufacturing specific Error logger.CR:91673

'***************************************************************************

Option Explicit

Implements IJNameRule

Const vbInvalidArg = &H80070057
Private Const IJDFabMargin = "{367C2C7E-02FA-45BB-A50C-60CF6B4D1D85}"
Private Const Module = "FabMarginNameRule: "
Private Const DEFAULTBASE = "Bxx"
Private Const WELDCOMPENSATIONTYPE As Long = 1000

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
        Err.Raise vbInvalidArg, Module, METHOD
    End If

    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    
   Dim BOStringName As String
    If TypeOf oObject Is IJConstMargin Then
        Dim oConstantMargin As IJConstMargin
        Set oConstantMargin = oObject
        If oConstantMargin.Type = WELDCOMPENSATIONTYPE Then
            BOStringName = Trim("Weld Compensation")
        Else
            BOStringName = Trim("Constant margin")
        End If
    ElseIf TypeOf oObject Is IJObliqueMargin Then
        BOStringName = Trim("Oblique margin")
    End If
    
    Dim strLocation As String
    strLocation = vbNullString
    Dim nCount As Long
    'Get the name count and location ID.
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
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 6001, , "RULES")
End Sub

'****************************************************************************************************
'Description
'   All the Naming Parents that need to participate in an objects naming are added here to the
'   IJElements collection. The parents added here are used in computing the name of the object in
'   ComputeName() of the same interface. Both these methods are called from the naming rule semantic.
'****************************************************************************************************
Private Function IJNameRule_GetNamingParents(ByVal oFabMargin As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oPlatePart As GSCADAsmHlpers.IJAssemblyChild
    Dim oIJAssemblyChild As GSCADAsmHlpers.IJAssemblyChild
    Dim oMfgChild As MfgChildHelperLib.IJMfgChild
    Dim oParent As Object

    'Retrieve the PlatePart object from the given oMfgTemplate object.
    Set oMfgChild = oFabMargin
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
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 6001, , "RULES")
End Function
 
