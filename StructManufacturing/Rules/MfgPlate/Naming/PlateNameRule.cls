VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 1999 & 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: GSCADStrMfgNameRules
'
'  Abstract: The file contains an implementation of the default naming rule
'            for the MfgPlate object in StructManufacturing UE.
'
'  History:
'     Yeun-gyu Kim  12/14/99    Initial creation.
'     Yeun-gyu Kim  03/06/00    Updated to set the block as the 1st naming parent correctly
'                               and the related codes.
'     Yeun-gyu Kim  04/20/00    Updated IJNameRule_ComputeName method to get the Block object
'                               correctly if given.
'     Yeun-gyu Kim  03/21/02    Changed the Relationship btw PlatePart and Manufacturing Object to
'                               IJMfgParent ----------StrMfgHierarchy - --------IJStrMfgChild
'                               and reflected the change that PlatePart is no longer support IJAssembly.
'     M. Veldhuizen 01/14/03    Remove the blockname from the parts name in order to avoid miscommunication with
'                               GNEST
'     Yeun-gyu Kim  01/25/03    Did general clean up in accordance with Marcel's change above and
'                               block object is no longer needed as the parent object for naming.
'     Marcel Veldhuizen 2004/04/21 Included correct error handling
'
'     Venkat (vkukutla)  2006/07/05  Modified GetCount into GetCountEx and modified the Naming Rule
'                                    to include work share location ID in order to avoid replication problems.CR:91673
'     Venkat (vkukutla)  2006/07/24  Existing error handler is replaced with Struct Manufacturing specific Error logger.CR:91673
'
'***************************************************************************

Option Explicit

Implements IJNameRule

Const vbInvalidArg = &H80070057

Private Const MODULE = "PlateNameRule: "
'Private Const DEFAULTBASE = "Bxx"

'*********************************************************************************************
' Description:
'   Creates a name for the object passed in. The name is based on the parents
'   name and object name. It is something like this: "Base Name" + "Object Name" + Index.
'   "Base Name" is derived from the Naming Parents and the Index is unique for the "Base Name".
'   It is assumed that all Naming Parents and the Object implement IJNamedItem.
'   The Naming Parents are added in AddNamingParents() of the same interface.
'   Both these methods are called from the naming rule semantic.
'
' Notes: This default Naming Rule was implemented based on the SHI's Part Naming standard.
'
'   1. There is two cases of Part Naming as follows:
'
'       CASE 1: BBB-SSS         (BBB: Block name, SSS: Single part name)
'       CASE 2: BBB-EEEE-SSS    (BBB: Block name, EEEE: Element name, SSS: Single part name)
'
'       Element name consists of Location name and Main member category name.
'       Single part name consists of Serial number and Secondary member category name.
'
'   2. It is also assumed that the Block as Naming Parent is passed in as the first one of the
'      "elements" collection.
'
'*********************************************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler

    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, MODULE, METHOD
    End If

    Dim oChildNamed As IJNamedItem
    Set oChildNamed = oObject
    
    Dim strChildName As String
    Dim strParentName As String
    Dim strNamedParentsString As String
    
    Dim strLocation As String
    strLocation = vbNullString
    
    Dim oNameCounter As IJNameCounter
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    
    Dim nCount As Long
    Dim oActiveConnection As IJDAccessMiddle
    Dim bPanel As Boolean
    bPanel = False
    
    Set oActiveConnection = GetActiveConnection
    'get the name count and location ID
    nCount = oNameCounter.GetCountEx(oActiveConnection.GetResourceManager(GetActiveConnectionName), _
                                             "GSCADStrMfgNamingRule_PlateDefaultGenerator", strLocation)
    
'''
''' Naming limitation in GNEST:                     Yeun-gyu Kim on 07/01/2003
'''
'''  (1) Max. 32 chars including the token delimiter (-), max. 8 chars for each token and max. number of token is 4.
'''      Max. length of name will be like for example "1234567-1234567-12345678-1234567".
'''
'''  (2) Can not use following special char.
'''      < > / \ # + - * @ = & , ! ? : % ( )
'''
'''   Note: It was confirmed from Kitano-san on 06/01/2003 that ( ) are also the special chars
'''         which can not be used and was recommended to use only the standard alphabet & numeric chars.
'''

'   There is only one object that is PlatePart in elements collection now since GetNamingParents function
'    was changed so. Anyway the codes below were kept as if there are more than one objects for the future use.
    If elements.Count > 0 Then
        Dim oParent As Object
        Dim oParentNamed As IJNamedItem
'        Dim strBlockName As String

        'Parts that are assigned to a block needs to be renamed without the blockinformation
        'Remember that parts that are processed but still in the unassigned folder have <NoBlk> as blockname
        'which should also be removed
        
        For Each oParent In elements
            'Regards the object which supports IJNamedItem as only the valid one.
            If TypeOf oParent Is IJNamedItem Then
                Set oParentNamed = oParent
                'As long as using "For Each" statement, it returns an object random from the collection.
                ' so, some special treatment is necessary to distinguish the Block object.
                'If TypeOf oParent Is IJBlockItem Then
                '    strBlockName = oParentNamed.Name
                'Else
                    strParentName = oParentNamed.Name
                'End If
                
                Set oParentNamed = Nothing
           End If
           If TypeOf oParent Is IJAssembly Then
                strParentName = "Panel-" & strParentName & "-"
                strParentName = strParentName & oNameCounter.GetCount(oActiveConnection.GetResourceManager(GetActiveConnectionName), strParentName)
                bPanel = True
           End If		   
        Next oParent
        
        Set oParent = Nothing
        Set oParentNamed = Nothing
        
        If (InStr(1, strParentName, "<") > 0 And InStr(1, strParentName, ">") > 0) Then
            ' This case means strParentName already contains the block name. Either a true blockname or "NoBlk"
            ' and it assumes that strParentName has always <B0.1.1.1.1>-xxxx-yyyy or <NoBlk>-xxxx-yyyy convention.
            Dim MyPos As Integer
            
            MyPos = InStr(1, strParentName, ">") + 1 ' The additional one is for removing the "-"
            ' Take out the blockinformation regardless of the true blockname or NoBlk.

            If strLocation <> vbNullString Then
                strChildName = Right(strParentName, Len(strParentName) - MyPos) + "-" + strLocation
            Else
                strChildName = Right(strParentName, Len(strParentName) - MyPos)
            End If
        Else
'            strChildName = GetLastToken(strParentName, "-")
            ' Just pass strParentName as strChildName.
            If strLocation <> vbNullString Then
                If Not bPanel Then
                    strChildName = strParentName + "-" + strLocation
                Else
                    strChildName = strParentName
                End If
            Else
                strChildName = strParentName
            End If
            
        End If
        
        strNamedParentsString = strParentName
    Else
        'This case will never happen but remained for just in case.
        
        'Check/Give the name if it does not have any generated default name.

'        If GetFirstToken(oChildNamed.Name, "-") <> DEFAULTBASE Then
'            'Default name will be the CASE 1 (ex; "Bxx-1P", "Bxx-2P", ..., "Bxx-nP")
            'Default name will be the CASE 1 (ex; "1P", "2P", ..., "nP")
            If strLocation <> vbNullString Then
                strChildName = strLocation + "-" + LTrim(Str(nCount)) + "P"
            Else
                strChildName = LTrim(Str(nCount)) + "P"
            End If
'        End If
        strNamedParentsString = "PlateDefaultName"
        
        Set oActiveConnection = Nothing
        Set oNameCounter = Nothing
    End If
    
    'Remove Extra Spaces before and after the name
    strNamedParentsString = Trim(strNamedParentsString)
    strChildName = Trim(strChildName)

    'Set the NamingParentsString property of the Active Entity.
    If oActiveEntity.NamingParentsString <> strNamedParentsString Then
        oActiveEntity.NamingParentsString = strNamedParentsString
    End If
    
    If oChildNamed.Name <> strChildName Then
        oChildNamed.Name = strChildName
    End If
    
    Set oChildNamed = Nothing

Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1034, , "RULES")
End Sub

'****************************************************************************************************
'Description
'   All the Naming Parents that need to participate in an objects naming are added here to the
'   IJElements collection. The parents added here are used in computing the name of the object in
'   ComputeName() of the same interface. Both these methods are called from the naming rule semantic.
'****************************************************************************************************
Private Function IJNameRule_GetNamingParents(ByVal oMfgPlate As Object) As IJElements
'
'   It was cleared that the block object is no longer needed to return as the parent object
'    since the ComputeName function does not need it any more.
'   So, all the related lines were commented out.   Yeun-gyu Kim on 01/25/2003
'
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oPlatePart As GSCADAsmHlpers.IJAssemblyChild    'In reality this should be IJMfgParent but GSCADAsmHlpers does not expose it.
'    Dim oIJAssemblyChild As GSCADAsmHlpers.IJAssemblyChild
    Dim oMfgChild As MfgChildHelperLib.IJMfgChild
    Dim oParent As Object

    'Retrieve the PlatePart object from the given oMfgPlate object by using Hierarchy relation
    Set oMfgChild = oMfgPlate
    Set oPlatePart = oMfgChild.GetParent
    
    If oPlatePart Is Nothing Then
        Dim oMfgPlatePart As IJMfgPlatePart
        Set oMfgPlatePart = oMfgPlate
        
        ' Try to get the plate part from the Mfg Plate AE(using DetailPlatePartInput relation)
        oMfgPlatePart.GetDetailedPart oPlatePart
    End If
    
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    Set oMfgChild = Nothing
    If oPlatePart Is Nothing Then
        GoTo ErrorHandler
    End If
    
'    'Add Block object as the first parent object into the collection.
'    Set oIJAssemblyChild = oPlatePart
'    Do While True
'        On Error Resume Next
'        Set oParent = oIJAssemblyChild.Parent
'        On Error GoTo ErrorHandler
'
'        If oParent Is Nothing Then Exit Do                      'There was no block found, terminate loop.
'        If Not TypeOf oParent Is IJAssemblyChild Then Exit Do   'There was no block found, terminate loop.
'
'        If TypeOf oParent Is IJBlockItem Then
'            'Add the found block as the naming parent.
'            Call IJNameRule_GetNamingParents.Add(oParent)
'            Exit Do
'        End If
'        Set oIJAssemblyChild = Nothing
'        Set oIJAssemblyChild = oParent
'        Set oParent = Nothing
'    Loop

    'Add the PlatePart object as the naming parent.
    Call IJNameRule_GetNamingParents.Add(oPlatePart)

'    Set oIJAssemblyChild = Nothing
    Set oParent = Nothing
    Set oPlatePart = Nothing
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1034, , "RULES")
End Function

'*********************************************************************************************
' Description:
'   Returns the case of Part Naming after determining it from the inputted strParentName.
'
'   According to the SHI's Part Naming standard, there is two cases of Part Naming as follows:
'
'       CASE 1: BBB-SSS         (BBB: Block name, SSS: Single part name)
'       CASE 2: BBB-EEEE-SSS    (BBB: Block name, EEEE: Element name, SSS: Single part name)
'
'*********************************************************************************************
Private Function DetermineCaseOfNaming(strParentName As String) As Integer
    Const METHOD = "DetermineCaseOfNaming"
    On Error GoTo ErrorHandler

    Select Case GetMainCategoryName(strParentName, "-")
        Case "SH":      'Shell Base Assembly.
            DetermineCaseOfNaming = 1
        Case "DE":      'Shell Base Assembly.
            DetermineCaseOfNaming = 1
        Case "GI":      'Girder Base Assembly.
            DetermineCaseOfNaming = 1
        Case "TB":      'T.BHD Base Assembly.
            DetermineCaseOfNaming = 1
        Case Else:
            DetermineCaseOfNaming = 2
    End Select

    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1034, , "RULES")
End Function

 
