VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShrinkageNameRule"
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
'            for the MfgShrinkageobject in StructManufacturing UE.
'
'  History:
'     Yeun-gyu Kim  01/24/00    Initial creation.
'     Marcel Veldhuizen 2004/04/22  Included correct error handling
'
'     Venkat (vkukutla)  2007/07/05  Modified GetCount into GetCountEx and modified the Naming Rule
'                                    to include work share location ID in order to avoid replication problems.CR:91673
'     Venkat (vkukutla)  2006/07/24  Existing error handler is replaced with Struct Manufacturing specific Error logger.CR:91673
'
'     Triveni           2006/07/09   DI- 103073  Shrinkage Stabilization
'
'***************************************************************************

Option Explicit

Implements IJNameRule

Const vbInvalidArg = &H80070057

Private Const MODULE = "ShrinkageNameRule: "
Private Const DEFAULTBASE = "Shrinkage"

'   Creates a name for the object passed in.
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)

   Const METHOD = "IJNameRule_ComputeName"
   On Error GoTo ErrorHandler

    Dim oChildNamed       As IJNamedItem
    Set oChildNamed = oObject
    Dim strNamedParentsString As String
    Dim BOStringName As String
    
    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, MODULE, METHOD
    End If
    
    If TypeOf oObject Is IJScalingShr Then
        Dim oShrinkage As IJScalingShr
        Set oShrinkage = oObject
        
        'Get the ParentPart of ShrinkageObject
        Dim oMfgChild As IJMfgChild
        Set oMfgChild = oObject
        Dim oParentPart As IJMfgParent
        Set oParentPart = oMfgChild.GetParent
        Set oMfgChild = Nothing
        Dim oNamedItem As IJNamedItem
        Dim sPartName As String
        Set oNamedItem = oParentPart
        sPartName = oNamedItem.Name
        Set oNamedItem = Nothing
        Set oObject = Nothing

        'Name Formula for:
        '"AssemblyShrinkage": is AssyName - AssyType - ShrinkageType
        'AssemblyPart: PartName - "AssemblyBased" - ShrkType
        'Part: PartName - "Part" - ShrkType
        If oShrinkage.ShrinkageMode = AssemblyShrinkage Then  ' gets the "Name" & "Type" of Assembly
            
            Dim lAssemblyType As Long
            Dim objAssembly As IJAssemblyBase
            Set objAssembly = oParentPart
            lAssemblyType = objAssembly.Type
            
            'To get Assembly Type As String
            Dim oCodeListMetaData As IJDCodeListMetaData
            Dim oCodeList As IJDCodeListValue
            Set oCodeListMetaData = oParentPart
            Set oParentPart = Nothing
            Dim strAssyType As String
            Dim iIndex As Integer
            Dim oTempcodeListColl As IJDInfosCol
            Set oTempcodeListColl = oCodeListMetaData.CodelistValueCollection("AssemblyType")
                If Not oTempcodeListColl Is Nothing Then
                    If oTempcodeListColl.Count > 0 Then
                        For iIndex = 1 To oTempcodeListColl.Count
                            Set oCodeList = oTempcodeListColl.Item(iIndex)
                            If lAssemblyType = oCodeList.ValueID Then
                                strAssyType = oCodeList.LongValue
                                Exit For
                            End If
                        Next iIndex
                    End If
                End If
            Set oCodeList = Nothing
            Set oCodeListMetaData = Nothing
            Set oTempcodeListColl = Nothing
           
            BOStringName = sPartName + " - " + strAssyType + " -"  'Appends PartName with AssemblyType
            
        
        ElseIf oShrinkage.ShrinkageMode = AssemblyPartShrinkage Then
            BOStringName = sPartName + " - " + Trim("AssemblyBased -")  'Append PartName with ShrinkageMode String
        ElseIf oShrinkage.ShrinkageMode = PartShrinkage Then
            BOStringName = sPartName + " - " + Trim("Part -")
        End If

        'to append Shrinkage Type to the StringName
        If oShrinkage.ShrinkageType = ScalingType Then
            BOStringName = BOStringName + " " + Trim("Edge")
        ElseIf oShrinkage.ShrinkageType = GlobalType Then
            BOStringName = BOStringName + " " + Trim("Axis")
        Else
            BOStringName = BOStringName + " " + Trim("Vector")
        End If
    End If
    
    Set oShrinkage = Nothing
    
    'Set the NamingParentsString property of the Active Entity.
    strNamedParentsString = BOStringName
    If oActiveEntity.NamingParentsString <> strNamedParentsString Then
        oActiveEntity.NamingParentsString = strNamedParentsString
    End If
    
    If oChildNamed.Name <> strNamedParentsString Then
        oChildNamed.Name = strNamedParentsString
    End If
    
    Set oChildNamed = Nothing
    Set oParentPart = Nothing

Exit Sub

ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 7001, , "RULES")
End Sub
'****************************************************************************************************
'Description
'   All the Naming Parents that need to participate in an objects naming are added here to the
'   IJElements collection. The parents added here are used in computing the name of the object in
'   ComputeName() of the same interface. Both these methods are called from the naming rule semantic.
'****************************************************************************************************
Private Function IJNameRule_GetNamingParents(ByVal oScalingShrinkage As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oPlatePart As GSCADAsmHlpers.IJAssemblyChild    'In reality this should be IJMfgParent but GSCADAsmHlpers does not expose it.

    Dim oMfgChild As MfgChildHelperLib.IJMfgChild
    Dim oParent As Object

    'Retrieve the PlatePart object from the given Shrinkage object by using
    Set oMfgChild = oScalingShrinkage
    Set oPlatePart = oMfgChild.GetParent
    
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

    Set oParent = Nothing
    Set oPlatePart = Nothing
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 7001, , "RULES")
End Function
 
