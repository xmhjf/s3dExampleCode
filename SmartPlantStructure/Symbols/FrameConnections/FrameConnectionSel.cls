VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrameConnectionSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : FrameConnectionSel.cls
'
'Author : R. Peringoth
'
'Description :
'    root selector for frame  connections
'
'History:
'   15-Jan-2004     MH             simplified "rule" code.  TR 52552 Check for IJSurface, not IJPlane.
'   14-April-2004   RP             Added code to check coplanarity of members for vertical corner brace
'   15-May-2006     MH             Curved Member impacts
'   15-Sept-2006    RP             Moved IsTwoMemberSystemsOkforVCB()to the common module
'                                  now checks if all three members are in the same quadrant (TR#104603)
'   26-Sept-2006    MH             TR 104591  ReadPathPointOnObjectsForFC
'
'   01-July-2009    RP             CR165819 - Added code in SetRelatedObjects() to handle the FC's
'                                  association to CanRules
'*******************************************************************************************************

Option Explicit

Private Const MODULE = "FrameConnectionSel"
Private m_oLocalizer As IJLocalizer

Const m_SelectorProgid As String = "SPSFCMacros.FrameConnectionSel"
Const m_SelectorName As String = "SPSFCMacros.FrameConnectionSel"
 
Implements SPSMembers.ISPSFCInputHelper
Implements SPSMembers.ISPSFCInputHelperEx
Implements IJDUserSymbolServices


Public Sub SelectorInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  
  Exit Sub
ErrorHandler:
  pIH.ReportError
End Sub
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
  On Error GoTo ErrorHandler
  ' TODO - Add your question below
  
  Exit Sub
ErrorHandler:
  pQH.ReportError
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)

  On Error GoTo ErrorHandler

'From Patrice,
'        Mike,
'
'        Returning ""  deletes the smart occurrence.
'        Returning "Undefined" deletes the members of the custom assembly.
'        Returning nothing should be an error.
'
'        Patrice

    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    Dim oFC As ISPSFrameConnection
    Dim strSelection As String

    Set oFC = pSL.SmartOccurrence
    IHStatus = ISPSFCInputHelper_ExecuteSelectionRule(oFC, strSelection)

    If IHStatus <> SPSFCInputHelper_Ok Then
        strSelection = "Undefined"
    ElseIf strSelection = "" Then
        strSelection = "Undefined"
    ElseIf strSelection = "Unsupported" Then
        strSelection = "Undefined"
    End If

    pSL.Add strSelection

    Exit Sub

ErrorHandler:
  pSL.ReportError
End Sub


' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
' ** End CM **


' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
'*************************************************************************
'Function
'IJDUserSymbolServices_GetDefinitionName
'
'Abstract
'Used during the execution of IJDDefinitionCollection::GetDefinitionByProgId to get the definition name
'based upon the definitionParameters passed in. It returns the definition name (pDefName) if it already
'exists within the collection. The name of a definition is the identifier of the definition object
'in the definition collection and assures its uniqueness in the given resource manager.
'
'Arguments
'definitionParameters
'
'Return
'
'Exceptions
'
'***************************************************************************

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function
'*************************************************************************
'Function
'IJDUserSymbolServices_InitializeSymbolDefinition
'
'Abstract
'Cleans the previous definition up and initializes the new one (Input, Output, Representation,
'RepresenationEvaluation, ...) constructing the symbol definition by (re)defining the inputs, outputs,
'representations, and representation.
'Note:  The previous setting of the definition must be reset in this method before constructing the new definition.
'
'Arguments:
'pDefinition  Symbol definition passed by reference that will be initialized in this method.
'
'Return
'S_OK  Operation succeeded.
'E_FAIL  Operation failed (no detail).
'
'Exceptions
'
'***************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
  Dim pDFact As New DefinitionFactory
  pDFact.InitAbstractSelector pSelector
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.definition = pSelector
'  pIH.InitAs m_FamilyProgid
  SelectorInputs pIH
  Dim pQH As IJDQuestionsHelper
  Set pQH = New QuestionHelper
  pQH.Selector = pSelector
  SelectorQuestions pQH
  
   Dim inputsProp As IMSDescriptionProperties
   inputsProp = pSelector.IJDInputs.Property
   pSelector.IJDInputs.Property = inputsProp Or igCOLLECTION_VARIABLE

End Sub
'*************************************************************************
'Function
'IJDUserSymbolServices_InstanciateDefinition
'
'Abstract
'Instantiates a persistent symbol definition object and initializes it for the first time,
'returning a pointer (ppSymbolDefDisp) to the IDispatch interface of the initialized symbol definition.
'
'Arguments:
'codeBase specifies the URL (or UNC) of the .cab file that can provides the dll associated to the symbol definition object (ActiveX® control packaging).
'definitionParameters  Definition parameters.
'pResourceMgr  resource manager to which the symbol definition will be connected.
'
'Return:
'S_OK  Operation succeeded.
'E_FAIL  Operation failed (no detail).
'
'Exceptions:
'
'***************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
'This function is not being called as we call the ISPSFCInputHelper_ExecuteSelectionRule in the FrameConnectionSel class which has the logic of selecting a particular item
Public Sub CMSelector(pRep As IJDRepresentation)
  Dim pSL As IJDSelectorLogic
  Set pSL = New SelectorLogic
  pSL.Representation = pRep
  SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

Private Property Get ISPSFCInputHelperEx_AMPOption() As SPSMembers.SPSFCAMPOptions
    ISPSFCInputHelperEx_AMPOption = SPSFCAMPOption_NONE
End Property

'====================
'BEGIN USER-CUSTOMIZABLE AREA

'*************************************************************************
'Function
'ISPSFCInputHelper_ExecuteSelectionRule
'
'Abstract
'ExecuteSelectionRule gets called by the command, and by the SmartOccurrence logic.
'The purpose is to implement code that can decide what kind of specific frameConnection to place.
'The set of available connections in the catalog needs to be known.
'The caller expects this function to return a string which is the name of the specific
'frame connection selected.  If none are suitable, then return the string "Unsupported"
'which is a connection type that has no connectivity to other members.
'
'
'Arguments
'FC as FrameConnection
'name as String
'
'Return
'Returns the catalog smartitem name as string and SPSFCInputHelperStatus representing the error occurred
'
'Exceptions
'
'***************************************************************************

Private Function ISPSFCInputHelper_ExecuteSelectionRule(ByVal oFC As SPSMembers.ISPSFrameConnection, ByRef selection As String) As SPSMembers.SPSFCInputHelperStatus

    On Error GoTo ErrorHandler

    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    Dim object1 As Object, object2 As Object
    Dim bSupportedHasPath As Boolean

    IHStatus = ISPSFCInputHelper_GetRelatedObjects(oFC, object1, object2)
    If IHStatus <> SPSFCInputHelper_Ok Then
        selection = "Unsupported"
        GoTo wrapup
    End If

    IHStatus = SPSFCInputHelper_UnexpectedError
    selection = "Unsupported"

    'If only one object is set, then let object1 be it.
    If object1 Is Nothing And Not object2 Is Nothing Then
        Set object1 = object2
        Set object2 = Nothing
    End If
        
    'If no input objects ...
    If object1 Is Nothing And object2 Is Nothing Then
        IHStatus = SPSFCInputHelper_Ok

    'Only one supporting object
    ElseIf object2 Is Nothing Then
    
        If ObjectIsOkSurface(object1) Then
            selection = "Surface-Default"

        ElseIf Not TypeOf object1 Is ISPSMemberSystem Then
            selection = "Unsupported"
    
        Else
            selection = MakeSingleMemberSelection(oFC, object1)
            
        End If

    Else
        'Two supporting objects...

        If TypeOf object1 Is ISPSMemberSystem And TypeOf object2 Is ISPSMemberSystem Then
            
            ' set the flag, if the supported is a curved member
            bSupportedHasPath = False
            If Not oFC Is Nothing Then
                If Not oFC.MemberSystem.LogicalAxis.CurveObject Is Nothing Then
                    bSupportedHasPath = True
                End If
            End If
             
            If IsTwoMemberSystemsOkforVCB(oFC, object1, object2) Then

                selection = "VerticalCornerBrace-WP2"
                
            ElseIf (bSupportedHasPath = True) Then
                ' While placing a curved member, reaching this point of having two member systems as supporting
                ' objects, happens only if user picked both the member systems from graphic view. So make the selection based
                ' on first object. For TR-CP·110509
                selection = MakeSingleMemberSelection(oFC, object1)

            Else
                
                selection = MakeSingleMemberSelection(oFC, object2)
                    
            End If

        'One is a MemberSystem
        ElseIf TypeOf object1 Is ISPSMemberSystem Then
        
            selection = MakeSingleMemberSelection(oFC, object1)
        
        ElseIf TypeOf object2 Is ISPSMemberSystem Then
        
            selection = MakeSingleMemberSelection(oFC, object2)
        
        Else
            selection = "Unsupported"

        End If

    End If 'if two objects

wrapup:
    ISPSFCInputHelper_ExecuteSelectionRule = SPSFCInputHelper_Ok
    Exit Function

ErrorHandler:
    Err.Clear
    ISPSFCInputHelper_ExecuteSelectionRule = IHStatus
    Exit Function
End Function



'*************************************************************************
'Function
'ISPSFCInputHelper_GetRelatedObjects
'
'Abstract
'Returns the objects in the ReferenceCollection used   as input by the FrameConnection.
'
'Arguments
'FC As ISPSFrameConnection
'RelatedObject1 As Object
'RelatedObject2 As Object
'
'Return
'Returns the input objects as arguments and SPSFCInputHelperStatus representing the error occurred
'
'Exceptions
'
'***************************************************************************
'GetRelatedObjects
'Returns the objects in the ReferenceCollection used for compute input by the FrameConnection.
'
'Do not return the object that is the same as the FrameConnection's MembersSystem, although
'it is, and must be, in the ReferenceCollection.
'
'First, return objects from the ReferenceCollection.
'If that was two objects, you're done.
'Otherwise, check for objects from FC.Joint.GetPointOn.  This can also be one or two objects.
'Any MemberSystem object returned by FC.Joint.GetPointOn must also be in the referenceCollection,
'otherwise this is an invalid connection.
'
'Although the ReferenceCollection frequently does contain more than one reference to the same
'object ( to respond to different interfaces ) GetRelatedObjects should "compress" that list.

Private Function ISPSFCInputHelper_GetRelatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, RelatedObject1 As Object, RelatedObject2 As Object) As SPSMembers.SPSFCInputHelperStatus

    Const METHOD = "ISPSFCInputHelper_GetRelatedObjects"
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim count As Long

    On Error GoTo ErrorHandler
    IHStatus = SPSFCInputHelper_UnexpectedError

    Set oRC = GetRefColl(FC)
  
    count = oRC.IJDEditJDArgument.GetCount
    If count > 1 Then
        Set RelatedObject1 = oRC.IJDEditJDArgument.GetEntityByIndex(2)
    End If
    
    If count > 2 Then
        Set RelatedObject2 = oRC.IJDEditJDArgument.GetEntityByIndex(3)
    End If

    IHStatus = SPSFCInputHelper_Ok
    ISPSFCInputHelper_GetRelatedObjects = IHStatus
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    ISPSFCInputHelper_GetRelatedObjects = IHStatus
End Function


'*************************************************************************
'Function
'ISPSFCInputHelper_SetRelatedObjects
'
'Abstract
'Sets the objects in the ReferenceCollection used as input by the FrameConnection.
'SetRelatedObjects is called by the command to enable this code to set connections to the FrameConnection's ReferenceCollection.
'
'Different connections will establish different relations according to what needs to be watched.
'This connection is the RootSelector connection, and serves as a hub for the "ByRule".
'
'See documentation and metadata for more details.
'
'Arguments
'FC As ISPSFrameConnection
'RelatedObject1 As Object
'RelatedObject2 As Object
'
'Return
'Returns the input objects as arguments and SPSFCInputHelperStatus representing the error occurred
'
'Exceptions
'
'***************************************************************************
Private Function ISPSFCInputHelper_SetRelatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, ByVal RelatedObject1 As Object, ByVal RelatedObject2 As Object) As SPSMembers.SPSFCInputHelperStatus

    ' our objective is to establish RefColl relations for the sake of searching using the SO rules.
    '
    Const METHOD = "ISPSFCInputHelper_SetRelatedObjects"
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim oCanRuleSupported As ISPSCanRule
    
    
    On Error GoTo ErrorHandler
    IHStatus = SPSFCInputHelper_UnexpectedError
    
    If Not IsFCCleared(FC) Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_FRAME_PREVIOUSREL_NOTCLEAR, "Frame connection is invalid because the previous connection was not removed. Change connection to unsupported and the re-select the supporting member.")
    End If

    If RelatedObject1 Is Nothing And Not RelatedObject2 Is Nothing Then
        Set RelatedObject1 = RelatedObject2
        Set RelatedObject2 = Nothing
    End If
    
    If Not RelatedObject1 Is Nothing Then
        If TypeOf RelatedObject1 Is ISPSFrameConnection Then
            IHStatus = SPSFCInputHelper_InvalidTypeOfObject
            Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_FRAME_CANNOTUSE_ASOBJ1, "Cannot use this frame connection as Obj1 to SetRelatedObject. See To Do List messages in the Troubleshooting Guide for more information.")
        End If
    End If
    If Not RelatedObject2 Is Nothing Then
        If TypeOf RelatedObject2 Is ISPSFrameConnection Then
            IHStatus = SPSFCInputHelper_InvalidTypeOfObject
            Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_FRAME_CANNOTUSE_ASOBJ2, "Cannot use this frame connection as Obj2 to SetRelatedObject. See To Do List messages in the Troubleshooting Guide for more information.")
        End If
    End If

    Set oRC = GetRefColl(FC)
    
    'user may be changing to an unsupported connection
    'if the member has an end can then we need to add that to the reference collection
    Set oCanRuleSupported = FC.GetCrossSectionObject(SPSFCSupported)
    
    oRC.IJDEditJDArgument.RemoveAll
    If oCanRuleSupported Is Nothing Then 'no can on the end, so establish relation with memb sys
        oRC.IJDEditJDArgument.SetEntity 1, FC.MemberSystem
    Else
        FC.SetCrossSectionObject SPSFCSupported, oCanRuleSupported
    End If
    
    If Not RelatedObject1 Is Nothing Then
        oRC.IJDEditJDArgument.SetEntity 2, RelatedObject1
    End If
    
    If Not RelatedObject2 Is Nothing Then
        oRC.IJDEditJDArgument.SetEntity 3, RelatedObject2
    End If

    IHStatus = SPSFCInputHelper_Ok
    ISPSFCInputHelper_SetRelatedObjects = IHStatus
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    ISPSFCInputHelper_SetRelatedObjects = IHStatus
End Function


'*************************************************************************
'Function
'ISPSFCInputHelper_UserAttributeMgmt
'
'Abstract
'Function on ISPSFACInputHelper to return the UserAttributeMgmt interface
'
'Arguments
'
'Return
'The interface is returned.
'
'Exceptions
'
'***************************************************************************
Private Property Get ISPSFCInputHelper_UserAttributeMgmt() As SPSMembers.IJUserAttributeMgmt
    Set ISPSFCInputHelper_UserAttributeMgmt = Nothing
End Property

'*************************************************************************
'Function
'ISPSFCInputHelper_ValidateLocatedObjects
'
'Abstract
'Validates the objects to be set as input to the ReferenceCollection for the FrameConnection.
'ValidateLocatedObjects is called by the command during Mouse Move events of PlaceMemberSystem command
'Depending on the type of FrameConnection selected in the ribbon bar its repective ValidateLocatedObjects
'function will be called to determine if that object is valid for placing the selected type of FrameConnection
'the located objects are validated and returned as relatedobjects in the argument
'
'Arguments
'FC As ISPSFrameConnection
'LocatedObject1 As Object
'LocatedObject2 As Object
'RelatedObject1 As Object
'RelatedObject2 As Object
'RelatedPositions X,Y,Z as Doubles
'LocatedPositions X,Y,Z as Doubles
'The other arguments are not being used and is for future enhancements
'
'Returns the input objects as arguments and SPSFCInputHelperStatus representing the error occurred
'
'Exceptions
'
'***************************************************************************
Private Function ISPSFCInputHelper_ValidateLocatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, _
    ByVal options As Long, ByVal snapDistance As Double, _
    ByVal LocatedObject1 As Object, ByVal LocatedObject2 As Object, _
    ByVal LocateX As Double, ByVal LocateY As Double, ByVal LocateZ As Double, _
    RelatedObject1 As Object, RelatedObject2 As Object, _
    RelatedObjectX As Double, RelatedObjectY As Double, RelatedObjectZ As Double) As SPSMembers.SPSFCInputHelperStatus
Const METHOD = "ISPSFCInputHelper_ValidateLocatedObjects"
' FC is optional and can be Nothing
' what we want to do here is if snap is enabled, then
' based on LocatedObject1's type we can return the x,y,z of nearby points of interest such as
' grid-intersections or pointOn joints, endpoints
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    Dim oFC As ISPSFrameConnection
    Dim oRelObj1 As Object, oRelObj2 As Object
    Dim bSupportedHasPath As Boolean

    bSupportedHasPath = False
    If Not FC Is Nothing Then
        If Not FC.MemberSystem.LogicalAxis.CurveObject Is Nothing Then
            bSupportedHasPath = True
        End If
    End If
    
    Set LocatedObject1 = SwitchLocatedPartToMemberSystem(LocatedObject1)
    Set LocatedObject2 = SwitchLocatedPartToMemberSystem(LocatedObject2)

    If Not LocatedObject1 Is Nothing Then
        If TypeOf LocatedObject1 Is ISPSFrameConnection Then
        
            Set oFC = LocatedObject1
            IHStatus = oFC.InputHelper.GetRelatedObjects(oFC, oRelObj1, oRelObj2)
        
            If IHStatus <> SPSFCInputHelper_Ok Then
                Set oRelObj1 = Nothing
                Set oRelObj2 = Nothing
                IHStatus = SPSFCInputHelper_Ok
            End If
            
            'user is modifying a connection of MS, and located dependent FC.
            'just return the located FC.MemberSystem
            If Not FC Is Nothing Then
                If FC.MemberSystem Is oRelObj1 Or FC.MemberSystem Is oRelObj2 Then
                    Set oRelObj1 = Nothing
                    Set oRelObj2 = Nothing
                End If
            End If
        
            If bSupportedHasPath Then
                Set RelatedObject1 = oFC.MemberSystem
                
            ElseIf oRelObj1 Is Nothing And oRelObj2 Is Nothing Then
                Set RelatedObject1 = oFC.MemberSystem
            
            ElseIf Not oRelObj1 Is Nothing And Not oRelObj2 Is Nothing Then
                If TypeOf oRelObj1 Is ISPSMemberSystem And TypeOf oRelObj2 Is ISPSMemberSystem Then
                    Set RelatedObject1 = oRelObj1
                    Set RelatedObject2 = oRelObj2
                ElseIf TypeOf oRelObj1 Is ISPSMemberSystem Then
                    Set RelatedObject1 = oRelObj1
                    Set RelatedObject2 = oFC.MemberSystem
                End If
            ElseIf Not oRelObj1 Is Nothing Then
                If TypeOf oRelObj1 Is ISPSMemberSystem Then
                    Set RelatedObject1 = oRelObj1
                    Set RelatedObject2 = oFC.MemberSystem
                End If
            End If
              
            If RelatedObject1 Is Nothing And RelatedObject2 Is Nothing Then
                Set RelatedObject1 = oFC.MemberSystem
            End If

        ElseIf TypeOf LocatedObject1 Is ISPSSplitMemberConnection Then
        
            If bSupportedHasPath Then
                Set RelatedObject1 = Nothing
                Set RelatedObject2 = Nothing
        
            Else
                Dim ii As Long
                Dim eles As IJElements
                Dim oSC As ISPSSplitMemberConnection
                Dim oSCResult As IJStructILCConnectionResult
                Dim iSplitStatus As SplitStatus
                
                Set oSC = LocatedObject1
                Set eles = oSC.InputObjects
                If eles.count = 2 Then
                    Set oSCResult = oSC
                    iSplitStatus = oSCResult.SplitParentStatusResult
                    If iSplitStatus = ssSplitSecond Then
                        Set RelatedObject1 = eles(1)
                        Set RelatedObject2 = eles(2)
                    Else
                        Set RelatedObject1 = eles(2)
                        Set RelatedObject2 = eles(1)
                    End If
                End If
            End If
        Else
            
            Set RelatedObject1 = LocatedObject1
            Set RelatedObject2 = LocatedObject2
        
        End If
    End If

    RelatedObjectX = LocateX
    RelatedObjectY = LocateY
    RelatedObjectZ = LocateZ

wrapup:
    ISPSFCInputHelper_ValidateLocatedObjects = SPSFCInputHelper_Ok
    
Exit Function
ErrorHandler:
    ISPSFCInputHelper_ValidateLocatedObjects = IHStatus
End Function


Public Function MakeSingleMemberSelection(oFC As ISPSFrameConnection, oSuppingMembSys As ISPSMemberSystem) As String

    Const METHOD = "MakeSingleMemberSelection"
    
    Dim selection As String
    Dim oSuppedMembSys As ISPSMemberSystem
    Dim geomCondition As Long
    
    Set oSuppedMembSys = oFC.MemberSystem
    
    'default selection, unless special conditions exist...
    selection = "Axis-Along"

    geomCondition = GeometricalCondition(oFC, oSuppingMembSys)

    'if types are same ...
    '    if colinear and PGSame, use Colinear
    '    elseif end-matched, use AxisEnd

    If oSuppedMembSys.MemberType.TypeCategory = oSuppingMembSys.MemberType.TypeCategory And _
        oSuppedMembSys.MemberType.Type = oSuppingMembSys.MemberType.Type Then
        
        ' if end-matched ...
        If (geomCondition And 2) = 2 Then
            selection = "Axis-End"
        End If
    
    'if supported type is Girt or Purlin, use Seated-Top, except if parallel
    ElseIf oSuppedMembSys.MemberType.Type = 106 Or oSuppedMembSys.MemberType.Type = 107 Then
        If (geomCondition And 4) <> 4 Then
            selection = "Seated-Top"
        End If
    
    'brace onto another type that is end-matched uses Axis-End
    ElseIf oSuppedMembSys.MemberType.TypeCategory = 3 Then      ' brace onto anything
        If (geomCondition And 2) = 2 Then                               ' end-matched
            selection = "Axis-End"
        End If
    End If

    MakeSingleMemberSelection = selection
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    MakeSingleMemberSelection = "Unsupported"
    Exit Function

End Function
    
  
Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub
