VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JigNameRule"
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
'            for the MfgJig object in StructManufacturing UE.
'
'  History:
'     Yeun-gyu Kim  01/06/00    Initial creation.
'     Yeun-gyu Kim  03/06/00    Updated to set the block as the naming parent instead of
'                               the PlatePart which was temporary used.
'     Yeun-gyu Kim  07/27/01    Updated IJNameRule_GetNamingParents method to get the parent object
'                               by using IJAssembly ----AssemblyHierarchy----IJAssemblyChild relationship.
'     ByungSam Lee  11/16/01    Changed the Relationship btw PlatePart and Manufacturing Object to
'                               IJMfgParent ----------StrMfgHierarchy - --------IJStrMfgChild
'     Marcel Veldhuizen 2004/04/22  Included correct error handling
'     K Kamph       2005/11/15  TR-88534 JigNameRule is changed so the names generated will be
'                               Unique and Intelligent
'
'     Venkat (vkukutla)  2006/07/05  Modified GetCount into GetCountEx and modified the Naming Rule
'                                    to include work share location ID in order to avoid replication problems.CR:91673
'     Venkat (vkukutla)  2006/07/24  Existing error handler is replaced with Struct Manufacturing specific Error logger.CR:91673
'***************************************************************************

Option Explicit


Implements IJNameRule

Const vbInvalidArg = &H80070057

Private Const Module = "JigNameRule: "
Private Const DEFAULTBASE = "J"
Private Const IID_IJPinJig As String = "{FE221533-5879-11D5-B86E-0000E2300200}"

'****************************************************************************************************
'Description
'   Returns the GetActiveConnectionName.
'****************************************************************************************************
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
    
End Function


'*********************************************************************************************
' Description:
'   Creates a name for the object passed in. The name is based on the parents
'   name and object type. It is something like this: "Base Name" + "Object Name" + Index.
'   "Base Name" is derived from the Naming Parents and the Index is unique for the "Base Name".
'   It is assumed that all Naming Parents and the Object implement IJNamedItem.
'   The Naming Parents are added in AddNamingParents() of the same interface.
'   Both these methods are called from the naming rule semantic.
'
'  TR-88534 ComputeName is changed so the name will be generated as:
'
'  If the parent is a Plate then the name will be:
'  "PJ_"+name of parent
'
'  If the parent is an Assembly then the name will be:
'  "AJ_"+name of parent+"-"+serial number
'
'   Serial number will be sequential and unique within each Assembly
'
'*********************************************************************************************
Private Sub IJNameRule_ComputeName(ByVal oPinJigBeingNamed As Object, ByVal oPinJigParentsColl As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler

    Dim oChildNamed As IJNamedItem
    Set oChildNamed = oPinJigBeingNamed
    
    If oChildNamed Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If

    Dim oNameCounter As IJNameCounter
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService

    Dim oCmnAppGenericUtil As IJDCmnAppGenericUtil
    Set oCmnAppGenericUtil = New CmnAppGenericUtil
    
    Dim oActiveConnection As IJDAccessMiddle
    oCmnAppGenericUtil.GetActiveConnection oActiveConnection
    
    Dim oResourceManager As Object
    Set oResourceManager = oActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim NameToBeSet As String
    Dim strNamedParentsString As String
    Dim nCount As Long
    
    Dim strLocation As String
    strLocation = vbNullString
    
    If oPinJigParentsColl.Count > 0 Then
        Dim oParent As Object
        For Each oParent In oPinJigParentsColl
        
            'Consider first parent that supports IJNamedItem.
            If TypeOf oParent Is IJNamedItem Then
            
                ' Not required immediately for parents that are Plate
                ' parts, but will be needed for CR 121461, so might as
                ' well do it right now.
                
                Dim oGlobalQuery As GSCADStructMfgGlobals.StructMfgGlobalsQuery
                Set oGlobalQuery = New GSCADStructMfgGlobals.StructMfgGlobalsQuery
    
                Dim oSiblings As IJElements
                Set oSiblings = oGlobalQuery.GetMfgPart(oParent, IID_IJPinJig)
                
                Dim NumSiblings As Long
                If Not oSiblings Is Nothing Then
                    NumSiblings = oSiblings.Count
                Else
                    NumSiblings = 0
                End If
                
                ' NOTE: As of V7.0.42.11, a pin jig for an individual
                ' plate part cannot have siblings.

                Dim oParentNamed As IJNamedItem
                Set oParentNamed = oParent
                
                Dim strParentName As String
                strParentName = oParentNamed.Name
    
                strNamedParentsString = strParentName
                
                If TypeOf oParent Is IJPlatePart Then
                    NameToBeSet = "PJ" + "_" + strParentName + "-"
                Else 'TypeOf oParent Is IJAssembly
                    NameToBeSet = "AJ" + "_" + strParentName + "-"
                End If
                
                nCount = oNameCounter.GetCountEx(oActiveConnection.GetResourceManager(GetActiveConnectionName), _
                                                 NameToBeSet, strLocation)
                                                 
                ' This is the "workshare" location entered during Site DB creation.
                If strLocation <> vbNullString Then
                    NameToBeSet = NameToBeSet + strLocation + "-"
                End If
                
                If NumSiblings = 0 Then Exit For ' Should not happen!
                
                Dim MyOrigName As String
                MyOrigName = oChildNamed.Name
                
                Dim NameAnchorLength As Long
                NameAnchorLength = Len(NameToBeSet)



                If Len(MyOrigName) > NameAnchorLength Then
                    If Left(MyOrigName, NameAnchorLength) = NameToBeSet Then
                        If NumSiblings + 1 >= CLng(Mid(MyOrigName, NameAnchorLength + 1)) Then
                           ' Name is already well-formed
                            NameToBeSet = MyOrigName
                            Exit For
                        End If
                    End If
                End If



                ' Iterate through all siblings to find
                ' first name that is not already taken.
                
                Dim SiblingNames() As String
                ReDim SiblingNames(1 To NumSiblings) As String
                Dim i As Long
                For i = 1 To NumSiblings
                    If Not oSiblings.Item(i) Is oPinJigBeingNamed Then
                        Dim oSiblingName As IJNamedItem
                        Set oSiblingName = oSiblings.Item(i)
                        SiblingNames(i) = oSiblingName.Name
                        Set oSiblingName = Nothing
                    End If
                Next
                
                For i = 1 To NumSiblings
                    Dim CandidateName As String
                    CandidateName = NameToBeSet + CStr(i)
                  
                    Dim ExistsCandidateName As Boolean
                    ExistsCandidateName = False
                    Dim j As Long
                    For j = 1 To NumSiblings
                         If CandidateName = SiblingNames(j) Then
                            ExistsCandidateName = True
                            Exit For
                         End If
                    Next j
                    
                    If Not ExistsCandidateName Then
                        NameToBeSet = CandidateName
                        Exit For
                    End If
                Next i
                
                Exit For ' Stop iterating over Pin jig's parents
            End If
        Next
    Else ' Shouldn't ever get here for a sane pin jig!
        nCount = oNameCounter.GetCountEx(oResourceManager, _
                                         "GSCADStrMfgNamingRule_PinJigDefaultGenerator", _
                                         strLocation)

        'Default name will be (ex; "J_NoParent-000001", "J_NoParent-000002", ..., "J_NoParent-999999")
        If strLocation <> vbNullString Then
            NameToBeSet = DEFAULTBASE + "_NoParent-" + "-" + strLocation + "-" + Format(nCount, "000000")
        Else
            NameToBeSet = DEFAULTBASE + "_NoParent-" + Format(nCount, "000000")
        End If
        
        strNamedParentsString = "JigDefaultName"
    End If

    'Set the NamingParentsString property of the Active Entity.
    If oActiveEntity.NamingParentsString <> strNamedParentsString Then
        oActiveEntity.NamingParentsString = strNamedParentsString
    End If
    
    If oChildNamed.Name <> NameToBeSet Then
        oChildNamed.Name = NameToBeSet
    End If

    GoTo CleanUp

ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 5001, , "RULES")
    
CleanUp:
    Set oSiblings = Nothing
    Set oCmnAppGenericUtil = Nothing
    Set oGlobalQuery = Nothing
    Set oNameCounter = Nothing
    Set oChildNamed = Nothing
    Set oParent = Nothing
    Set oParentNamed = Nothing
    Set oActiveConnection = Nothing
    Set oResourceManager = Nothing
    Set oPinJigParentsColl = Nothing
    Set oPinJigBeingNamed = Nothing
    Set oActiveEntity = Nothing

End Sub

'****************************************************************************************************
'Description
'   All the Naming Parents that need to participate in an objects naming are added here to the
'   IJElements collection. The parents added here are used in computing the name of the object in
'   ComputeName() of the same interface. Both these methods are called from the naming rule semantic.
'****************************************************************************************************
Private Function IJNameRule_GetNamingParents(ByVal oMfgJig As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oMfgChild As MfgChildHelperLib.IJMfgChild
    Dim oParent As Object

    'Retrieve the PlatePart object from the given oMfgJig object.
    Set oMfgChild = oMfgJig
    On Error Resume Next
    Set oParent = oMfgChild.GetParent
    On Error GoTo ErrorHandler
    
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
    'Add parent object into the collection.
    Do While Not oParent Is Nothing
        If TypeOf oParent Is IJPlatePart Or TypeOf oParent Is IJAssembly Then
            'Add the found object as the naming parent.
            Call IJNameRule_GetNamingParents.Add(oParent)
            Exit Do
        End If
        Set oMfgChild = oParent
        Set oParent = oMfgChild.Parent
    Loop

    Set oMfgJig = Nothing
    Set oParent = Nothing
    Set oMfgChild = Nothing

Exit Function
ErrorHandler:
     Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 5001, , "RULES")
End Function

 
