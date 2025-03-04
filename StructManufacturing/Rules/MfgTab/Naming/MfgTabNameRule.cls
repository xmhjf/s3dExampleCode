VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MfgTabNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2010, Intergraph Corporation.  All rights reserved.
'
'  Project: GSCADStrMfgNameRules
'
'  Abstract: The file contains an implementation of the default naming rule
'            for the MfgTab in StructManufacturing UE.
'
'  History:
'     Suma Mallena  22/10/10    Initial creation.
'
'***************************************************************************

Option Explicit

Implements IJNameRule

Const vbInvalidArg = &H80070057

Private Const MODULE = "MfgTabNameRule: "

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
    If TypeOf oObject Is IJStrMfgTab Then
        Dim oMfgTab As IJStrMfgTab
        Set oMfgTab = oObject
        Dim oIJDPart As IJDPart
        
        Set oIJDPart = oMfgTab.GetCatalogPart
        Dim strName As String
        strName = oIJDPart.PartNumber
        
        BOStringName = Trim(strName)
     
        Set oMfgTab = Nothing
    End If
    
    Dim strLocation As String
    strLocation = vbNullString
    Dim nCount As Long
    
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
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 9001, , "RULES")
End Sub
'****************************************************************************************************
'Description
'   All the Naming Parents that need to participate in an objects naming are added here to the
'   IJElements collection. The parents added here are used in computing the name of the object in
'   ComputeName() of the same interface. Both these methods are called from the naming rule semantic.
'****************************************************************************************************
Private Function IJNameRule_GetNamingParents(ByVal oMfgTab As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oPart       As GSCADAsmHlpers.IJAssemblyChild    'In reality this should be IJMfgParent but GSCADAsmHlpers does not expose it.
    Dim oMfgChild   As MfgChildHelperLib.IJMfgChild
    Dim oParent     As Object

    'Retrieve the PlatePart object from the given Tab object by using
    Set oMfgChild = oMfgTab
    Set oPart = oMfgChild.GetParent
    
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    Set oMfgChild = Nothing
    If oPart Is Nothing Then
        GoTo ErrorHandler
    End If
        
    'Add the Part object as the naming parent.
    Call IJNameRule_GetNamingParents.Add(oPart)

    Set oParent = Nothing
    Set oPart = Nothing
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 9001, , "RULES")
End Function
 
