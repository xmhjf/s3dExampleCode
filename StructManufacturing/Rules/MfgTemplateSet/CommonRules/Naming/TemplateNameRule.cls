VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateNameRule"
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
'            for the MfgTemplate object in StructManufacturing UE.
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
Private Const IJDMfgTemplateSet = "{0D5FB0AA-7C0B-4DC3-9F7C-583741D6F542}"
Private Const MODULE = "TemplateNameRule: "
Private Const DEFAULTBASE = "Bxx"

'*********************************************************************************************
' Description:
'   Creates a name for the object passed in. The name is based on the parents
'   name and object name. It is something like this: "Base Name" + "Object Name" + Index.
'   "Base Name" is derived from the Naming Parents and the Index is unique for the "Base Name".
'   It is assumed that all Naming Parents and the Object implement IJNamedItem.
'   The Naming Parents are added in AddNamingParents() of the same interface.
'   Both these methods are called from the naming rule semantic.
'
' Notes: This default Naming Rule was implemented based on the SHI's Part Naming standard and
'        the following is standard pattern.
'
'        TBBB-SSS    (T: An initial letter of Template, BBB: Block name, SSS: Single part name)
'
'        Single part name consists of Serial number and Secondary member category name.
'
'        It is also assumed that the Block as Naming Parent is passed in as the first one of the
'        "elements" collection.
'
'*********************************************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler

    Dim oActiveConnection As IJDAccessMiddle

    Set oActiveConnection = GetActiveConnection

    Dim oNameCounter        As IJNameCounter
    Dim oChildNamed         As IJNamedItem
    Dim oParent             As Object
    Dim oParentNamed        As IJNamedItem

    Dim strChildName        As String
    Dim strParentName       As String
    Dim strBlockName        As String
    Dim strNamedParentsString As String
    Dim nCount              As Long
    Dim oGlobalQuery        As GSCADStructMfgGlobals.StructMfgGlobalsQuery
    Dim strChildNewName  As String
    Set oGlobalQuery = New GSCADStructMfgGlobals.StructMfgGlobalsQuery
    Dim nTempCount As Long
    Dim bSetNewName  As Boolean
    
    Dim strLocation As String
    strLocation = vbNullString
    
    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, MODULE, METHOD
    End If

    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    'get the name count and location ID from below method
    nCount = oNameCounter.GetCountEx(oActiveConnection.GetResourceManager(GetActiveConnectionName), _
                                        "GSCADStrMfgNamingRule_TemplateDefaultGenerator", strLocation)

    Set oChildNamed = oObject

    If elements.Count > 0 Then
        For Each oParent In elements
            'Regards the object which supports IJNamedItem as only the valid one.
            If TypeOf oParent Is IJNamedItem Then
                Set oParentNamed = oParent
                'As long as using "For Each" statement, it returns an object random from the collection.
                ' so, some special treatment is necessary to distinguish the Block object.
                If TypeOf oParent Is IJBlockItem Then
                    strBlockName = oParentNamed.Name
                Else    'i.e. It implicts PlatePart object.
                    strParentName = oParentNamed.Name
                    Dim oElements As IJElements
                    Set oElements = oGlobalQuery.GetMfgPart(oParent, IJDMfgTemplateSet)
                End If
                
                Set oParentNamed = Nothing
           End If
        Next oParent

        'Assign the default name if there was no name taken from the Naming Parent object.
        If (Len(strBlockName) = 0) Then strBlockName = DEFAULTBASE
        If (Len(strParentName) = 0) Then strParentName = "X99-9XXX-9P"
       
        Dim nTemplateNumber As Integer
        If Not oElements Is Nothing Then
            nTemplateNumber = oElements.Count
        End If
        
        'PATTERN: TBBB-SSS  (T: An initial letter of Template, BBB: Block name, SSS: Single part name)
        If (InStr(1, strParentName, "<") > 0 And InStr(1, strParentName, ">") > 0) Then
            Dim MyPos As Integer
            MyPos = InStr(1, strParentName, ">") + 1 'The additional one is for removing the "-"
            strChildName = "T" + strBlockName + "-" + Right(strParentName, Len(strParentName) - MyPos)
        Else
            strChildName = "T" + strBlockName + "-" + strParentName
        End If
        strChildNewName = strChildName + "_" + CStr(nTemplateNumber)
        
        'Check if this name already exist
        If Not oElements Is Nothing Then
            Dim oNamedItem As IJNamedItem
            For nTempCount = 1 To oElements.Count
                Set oNamedItem = oElements.Item(nTempCount)
                
                Dim iPos As Integer
                iPos = InStr(1, oNamedItem.Name, strChildNewName, vbTextCompare)
                If iPos > 0 Then
                    nTemplateNumber = nTemplateNumber + 1
                    strChildName = strChildName + "_" + CStr(nTemplateNumber)
                    bSetNewName = True
                    Set oNamedItem = Nothing
                    Exit For
                End If
                Set oNamedItem = Nothing
            Next
        End If
        If Not (bSetNewName) And (Not (oElements Is Nothing)) Then
            If oElements.Count > 0 Then
                strChildName = strChildName + "_" + CStr(nTemplateNumber)
            End If
        End If
        Set oElements = Nothing
        strNamedParentsString = strBlockName + "/" + strParentName
        
        If strLocation <> vbNullString Then
            strChildName = strChildName + "-" + strLocation
        End If
        
    Else
        'Default name will be (ex; "TBxx-1P", "TBxx-2P", ..., "TBxx-nP")
        If strLocation <> vbNullString Then
            strChildName = "T" + DEFAULTBASE + "-" + strLocation + "-" + LTrim(Str(nCount)) + "P"
        Else
            strChildName = "T" + DEFAULTBASE + "-" + LTrim(Str(nCount)) + "P"
        End If
        
        strNamedParentsString = "TemplateDefaultName"
    End If

    'Set the NamingParentsString property of the Active Entity.
    If oActiveEntity.NamingParentsString <> strNamedParentsString Then
        oActiveEntity.NamingParentsString = strNamedParentsString
    End If
    
    If oChildNamed.Name <> strChildName Then
        oChildNamed.Name = strChildName
    End If
    
    Set oActiveConnection = Nothing

    Set oNameCounter = Nothing
    Set oChildNamed = Nothing
    Set oParent = Nothing
    Set oParentNamed = Nothing

Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3001, , "RULES")
End Sub

'****************************************************************************************************
'Description
'   All the Naming Parents that need to participate in an objects naming are added here to the
'   IJElements collection. The parents added here are used in computing the name of the object in
'   ComputeName() of the same interface. Both these methods are called from the naming rule semantic.
'****************************************************************************************************
Private Function IJNameRule_GetNamingParents(ByVal oMfgTemplate As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oPlatePart As GSCADAsmHlpers.IJAssemblyChild
    Dim oIJAssemblyChild As GSCADAsmHlpers.IJAssemblyChild
    Dim oMfgChild As MfgChildHelperLib.IJMfgChild
    Dim oParent As Object

    'Retrieve the PlatePart object from the given oMfgTemplate object.
    Set oMfgChild = oMfgTemplate
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
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 3001, , "RULES")
End Function
 
