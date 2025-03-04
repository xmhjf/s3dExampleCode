Attribute VB_Name = "AssemblyConnCommon"

'*******************************************************************
'
'Copyright (C) 2006-15 Intergraph Corporation. All rights reserved.
'
'File : AssemblyConnCommon.bas
'
'Author : R. P
'
'Description :
'    Generic methods useful for assemly connections are placed in this module
'
'History:
'
' 11/10/05   R.P    Replaced the common.bas in AssemblyConnection project
'                           with this one
' Jun-21-2006 AS    Separated all the Split Migration Utilities from AssemblyConnCommon
'
' 19/Sep/06   SS   TR#105292 - changes to base plate asm connection to add plate part
'                  as an assembly child of its parent parent
' 09/aug/14   NK   CR-250020 - Create AC for ladder support or cage support to ladder rail
' 04/Nov/14   NK   CR-CP-262343  Create AC ladder rung penetrating support
' 03/Nov/14  MDT/GH CR-CP-250198  Lapped AC for traffic items
' 03/Dec/14  CSK/CM CR-250022 Connections and free end cuts at ends of ladder rails
' 13/Jan/15  CM/MDT/CSK CR-250025 and CR-254702 To Support Splits and Miters
' 19/Feb/15   NK   CR-CP-267170  Mitered Stadnard AC to be considered even for Mbr End to End Offset case
' 24/Mar/15  MDT/RPK  TR-269306 Handrail placed on builtup member, cannot create assembly connection, errors,
'                     Added  checks related to design/Builtup member as the bounding member
' 08/Aug/15  svsmylav CR-272424: Added new method 'PartClassExistsInDB' to check if given smart class is in the DB.
' 08/Aug/15  knukala  CR-CP-226692: Removing Standard AC method and adding an optional attribute to PatrClassExistInDB() method.
' 30/Oct/15  svsmylav TR-280722: Added two new methods 'IsMbrSplitByPlateCase' and 'IsInterfaceSupported'.
' 11/Dec/15  PYk      Moved constant E_FAIL, Enum MemberCategoryAndType to CommonEnumsAndConstants.bas file and
'                      Moved GetFrameConnectionType, AreMembersIdentical to MemberUtilities.bas file.
' 13/May/16  svsmylav  TR-294128: to avoid record exceptions, modified 'ValidSectionType' to have
'                      not 'Nothing' checks before dereferencing respective variables.
'********************************************************************
Option Explicit
Private Const MODULE = "AssemblyConnCommon"
Private Const CUSTOMPLATEPROGID = "StructCustomPlatePart.StructCustomPlatePart"
Private Const STRUCTFEATUREPROGID = "StructFeature.StructFeature2"
Private Const STRUCTMEMPARTPROGID = "SPSMembers.SPSMemberPartPrismatic"
Public Const IJPlane = "{4317C6B3-D265-11D1-9558-0060973D4824}"
Public Const SPSvbError = vbObjectError + 512
Public Const IJPlatePart = "{780F26C2-82E9-11D2-B339-080036024603}"

'*************************************************************************
'Function
'GenerateNameForFeature
'
'Abstract
'Establishes a name rule for the given feature
'
'Arguments
'oFeature is input, the feature to be named
'
'Return
'Errors are written to the error log file and cleared.
'
'Exceptions
'
'***************************************************************************

Public Sub GenerateNameForFeature(oFeature As IJStructFeature2)
    Const METHOD = "GenerateNameForFeature"
    On Error GoTo ErrorHandler
    
    Dim NamingRules As IJElements
    Dim oNameRuleHolder As IJDNameRuleHolder
    Dim oNameRuleHlpr As IJDNamingRulesHelper
    Dim oNameRuleAE As IJNameRuleAE
    
    Set oNameRuleHlpr = New NamingRulesHelper
    
    oNameRuleHlpr.GetEntityNamingRulesGivenProgID STRUCTFEATUREPROGID, NamingRules
    
    If NamingRules.Count > 0 Then
        Set oNameRuleHolder = NamingRules.Item(1)
    End If
    
    Call oNameRuleHlpr.AddNamingRelations(oFeature, oNameRuleHolder, oNameRuleAE)
    
    Set NamingRules = Nothing
    Set oNameRuleHolder = Nothing
    Set oNameRuleAE = Nothing
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
Public Sub GenerateNameForMember(obj As Object)
Const METHOD = "GenerateNameForMember"
On Error GoTo ErrorHandler

    Dim NameRule As String
    Dim found As Boolean
    found = False
    On Error Resume Next
      
    Dim NamingRules As IJElements
    Dim oNameRuleHolder As GSCADGenericNamingRulesFacelets.IJDNameRuleHolder
    Dim oActiveNRHolder As GSCADGenericNamingRulesFacelets.IJDNameRuleHolder
    Dim oNameRuleHlpr As GSCADNameRuleSemantics.IJDNamingRulesHelper
    Set oNameRuleHlpr = New GSCADNameRuleHlpr.NamingRulesHelper
    Call oNameRuleHlpr.GetEntityNamingRulesGivenProgID(STRUCTMEMPARTPROGID, NamingRules)
    Dim ncount As Integer
    Dim oNameRuleAE As IJNameRuleAE
      
    For ncount = 1 To NamingRules.Count
        Set oNameRuleHolder = NamingRules.Item(1)
    Next ncount

    Call oNameRuleHlpr.AddNamingRelations(obj, oNameRuleHolder, oNameRuleAE)
    Set oNameRuleHolder = Nothing
    
    Set oActiveNRHolder = Nothing
    Set oNameRuleHolder = Nothing
    Set oNameRuleAE = Nothing
 
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub


'*************************************************************************
'Function
'ValidSectionType
'
'Abstract
'Checks whether the member part's cross-section is a valid one
'
'Arguments
'oMemb is input, the member part
'
'Return
'The function is a boolean and is returned true when the section type is valid.
'
'Exceptions
'
'***************************************************************************

Public Function ValidSectionType(oMemb As ISPSMemberPartPrismatic) As Boolean
    Const METHOD = "ValidSectionType"
    On Error GoTo ErrorHandler
    Dim strSectionType As String
    ValidSectionType = False
    
    If Not oMemb Is Nothing Then
        Dim oSPSCrossSection As ISPSCrossSection
        Set oSPSCrossSection = oMemb.CrossSection
        If Not oSPSCrossSection Is Nothing Then
            Dim oCrossSection As IJCrossSection
            Set oCrossSection = oSPSCrossSection.definition
            If Not oCrossSection Is Nothing Then
                strSectionType = oCrossSection.Type
            End If
        End If
    End If
        
    If strSectionType = "W" Or strSectionType = "S" Or strSectionType = "HP" Or strSectionType = "M" _
    Or strSectionType = "HSSC" Or strSectionType = "CS" Or strSectionType = "PIPE" _
    Or strSectionType = "L" Or strSectionType = "C" Or strSectionType = "MC" _
    Or strSectionType = "WT" Or strSectionType = "MT" Or strSectionType = "ST" _
    Or strSectionType = "2L" Or strSectionType = "RS" Or strSectionType = "HSSR" _
    Then
        ValidSectionType = True
    End If
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

' make the plate assembly child of asm conn or its parent
Public Sub AddAsmConnToAsmParent(oAsmConnSmartOcc As IJSmartOccurrence, oPlatePart As IJStructCustomPlatePart)
Const METHOD = "AddAsmConnToAsmParent"
On Error GoTo ErrorHandler

    Dim oAssemblyparent As IJAssembly
    If TypeOf oAsmConnSmartOcc Is IJAssembly Then
        Set oAssemblyparent = oAsmConnSmartOcc
    Else
        Dim pRelationHelper As IMSRelation.DRelationHelper
        Dim pCollectionHelper As IMSRelation.DCollectionHelper
        Set pRelationHelper = oAsmConnSmartOcc
        
        Set pCollectionHelper = pRelationHelper.CollectionRelations("IJFullObject", "toAssembly1")
        If Not pCollectionHelper Is Nothing Then
            'the input smartocc may not have a CA parent, so need to check for count before accessing element
            If pCollectionHelper.Count > 0 Then 'added to fix TR#110987 - rperingo 12/7/06
                Dim oParentSmartOcc As IJSmartOccurrence
                Set oParentSmartOcc = pCollectionHelper.Item(1)
                If TypeOf oParentSmartOcc Is IJAssembly Then
                    Set oAssemblyparent = oParentSmartOcc
                End If
            End If
        End If
    End If
    
    If Not oAssemblyparent Is Nothing Then oAssemblyparent.AddChild oPlatePart
    
    Set oAssemblyparent = Nothing
    Set pRelationHelper = Nothing
    Set pCollectionHelper = Nothing
    Set oAssemblyparent = Nothing
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Function bIsTrafficItem(MemberPart As ISPSMemberPartPrismatic, Optional eMemberType As Integer, Optional eMemberTypeCategory As Integer) As Boolean
                          
    'Checks whether member part is a traffic element or not.
                                   
    Const METHOD = "bIsTrafficItem"
    On Error GoTo ErrorHandler
    
    bIsTrafficItem = False
    
    If MemberPart Is Nothing Then
        Exit Function
    End If
    
    eMemberType = MemberPart.MemberType.Type
    eMemberTypeCategory = MemberPart.MemberType.TypeCategory
    
    If (eMemberTypeCategory = MemberCategoryAndType.HandRailElement Or eMemberTypeCategory = MemberCategoryAndType.StairElement Or _
        eMemberTypeCategory = MemberCategoryAndType.LadderElement) Then
        bIsTrafficItem = True
    Else
        bIsTrafficItem = False
    End If
     
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
    
End Function
'********************************************************************
' Routine: PartClassExistsInDB
' Description:  Check if the given PartClass exist in Catalog
'
'Note: implementation is taken from CheckPartClassExist method in SMAssyConRul\RulesCommon.bas.
' However, this is done to avoid the selector that handles plant AC to be depending on module specific to marine content
'********************************************************************
Public Function PartClassExistsInDB(sPartClassName As String, oClassMoniker As IUnknown) As Boolean
On Error GoTo ErrorHandler
    Const sMethod = "PartClassExistsInDB"
    On Error GoTo ErrorHandler
    
    Dim oMidctx As IJMiddleContext
    Set oMidctx = New GSCADMiddleContextProj.GSCADMiddleContext
    Dim oResourceManager As IUnknown
    Set oResourceManager = oMidctx.GetResourceManager("Catalog")
    
    On Error Resume Next
    Dim oNamingCntxObj As NamingContextObject
    Set oNamingCntxObj = New NamingContextObject

    Set oClassMoniker = oNamingCntxObj.ObjectMoniker(oResourceManager, sPartClassName)
    On Error GoTo ErrorHandler
    
    If Not oClassMoniker Is Nothing Then
        PartClassExistsInDB = True
    End If
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, sMethod
    
End Function

'*************************************************************************
'Function
'    IsMbrSplitByPlateCase
'
'Abstract
'    This method helps checking for generic assembly connection where
'    connectables on port1 and port2 are different, for member split by
'    plate case
'
'input
'    oPort1 as IJPort
'    oPort2 as IJPort
'
'Return
'    Boolean value: True if it is member split by plate case, otherwise False
'    Optionally oPointConn can be collected as IJPoint3D
'
'Exceptions
'
'***************************************************************************

Public Function IsMbrSplitByPlateCase(oPort1 As IJPort, oPort2 As IJPort, _
                                Optional ByRef oPointConn As IJPoint = Nothing) As Boolean
    Const METHOD = "IsMbrSplitByPlateCase"
    On Error GoTo ErrHandler
    
    IsMbrSplitByPlateCase = False 'Initialize
    
    'Check for invalid inputs:
    '-> check for null
    If oPort1 Is Nothing Then Exit Function   '*** EXIT ***
    If oPort2 Is Nothing Then Exit Function   '*** EXIT ***
    
    '-> check if bounding and bounded parts are same (other type of generic AC)
    Dim oPart1 As Object
    Dim oPart2 As Object
    
    Set oPart1 = oPort1.Connectable
    Set oPart2 = oPort2.Connectable
    If oPart1 Is oPart2 Then Exit Function   '*** EXIT ***
    
    '-> check for member split by point (or similar case)
    If TypeOf oPart1 Is ISPSMemberPartPrismatic And _
        TypeOf oPart2 Is ISPSMemberPartPrismatic Then Exit Function   '*** EXIT ***
        
    '-> check for plate part
    Dim oMemberPart As ISPSMemberPartPrismatic
    Dim oPlatePart As Object
    If TypeOf oPart1 Is ISPSMemberPartPrismatic Then
        If IsInterfaceSupported(oPart2, IJPlatePart) Then
            Set oMemberPart = oPart1
            Set oPlatePart = oPart2
        End If
    ElseIf TypeOf oPart2 Is ISPSMemberPartPrismatic Then
        If IsInterfaceSupported(oPart1, IJPlatePart) Then
            Set oMemberPart = oPart2
            Set oPlatePart = oPart1
        End If
    End If
    If oPlatePart Is Nothing Then Exit Function   '*** EXIT ***
    
    'Get member system and root plate system
    Dim oParent As Object
    Dim oDesignChild As IJDesignChild
    Dim oMemberSystem As ISPSMemberSystem
    Dim oPlateSystem As Object
    
    'Member
    Set oDesignChild = oMemberPart
    Set oMemberSystem = oDesignChild.GetParent
    
    'Plate
    Dim iCount As Integer
    
    Set oDesignChild = oPlatePart
    For iCount = 1 To 2
        Set oParent = oDesignChild.GetParent
        Set oDesignChild = oParent
    Next iCount
    Set oPlateSystem = oDesignChild
    
    'Get inputs of each split connection on the member system: if connectable of
    ' an input port happens to be the plate system, then it is member-split-by-plate case
    Dim eleSplits As IJElements
    Dim eleInputs As IJElements
    Dim ii As Long
    Dim jj As Long
    Dim iSC As ISPSSplitMemberConnection

    Set eleSplits = oMemberSystem.SplitConnections
    Dim oPort As IJPort
        
    For ii = 1 To eleSplits.Count
        Set iSC = eleSplits(ii)
        Set eleInputs = iSC.InputObjects
        For jj = 1 To eleInputs.Count
            On Error Resume Next
            If TypeOf eleInputs(jj) Is IJPort Then
                Set oPort = eleInputs(jj)
                If oPort.Connectable Is oPlateSystem Then
                    Set oPointConn = iSC
                    IsMbrSplitByPlateCase = True
                    Exit For 'exit jj loop
                End If
            End If
            On Error GoTo 0
        Next jj
        If IsMbrSplitByPlateCase = True Then Exit For 'exit ii loop
    Next ii

    Exit Function
ErrHandler:
    HandleError MODULE, METHOD
End Function
'*************************************************************************
'Function
'    IsInterfaceSupported
'
'Abstract
'    This method checks if the given interface is supported by
'    given object.
'    Example: can be used to check if connectable of a port(i.e. part)
'    is a plate part by passing the connectable and "IJPlatePart" string.
'
'input
'    oBusinessObject As iJDObject
'    sInterface As String (similar to POM method:
'    'SupportsInterface' which accepts interface name or interface id/GUID)
'
'Return
'    Boolean value: True if the interface is supported, otherwise False
'
'Exceptions
'
'***************************************************************************

Public Function IsInterfaceSupported(oBusinessObject As iJDObject, sInterface As String) As Boolean
    Const MT = "Congifuration.IsInterfaceSupported"

    'Get POM
    On Error Resume Next
    Dim oPOM As IJDPOM

    Set oPOM = oBusinessObject.ResourceManager
    
    'Get object moniker
    Dim oMoniker As IMoniker
    Set oMoniker = oPOM.GetObjectMoniker(oBusinessObject)
    Err.Clear
    On Error GoTo ErrorHandler
    
    If Not oPOM Is Nothing Then
        If Not oMoniker Is Nothing Then
            'Set return value to be the same as that returned by the method on POM
            IsInterfaceSupported = oPOM.SupportsInterface(oMoniker, sInterface)
        End If
    End If
    
    Exit Function

ErrorHandler:
    HandleError MODULE, MT
End Function


