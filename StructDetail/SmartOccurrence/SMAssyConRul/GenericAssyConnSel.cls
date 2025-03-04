VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenericAssyConnSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Selector:
'   ---------
'   It selects the list of available AC Items for the Stiffner Hitting the multiple ports
'
'   CR-CP-205474     Define classes for Stiffener End Cuts that include additional ports (schema)
'
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parameterize the selection
'History :
'    15/June/2016    knukala   TR-CP-295640: Generic AC is updated when changing the corner feature type.
'
Option Explicit

Const m_SelectorProgID As String = CUSTOMERID + "AssyConRul.GenericAssyConnSel"
Const m_SelectorName As String = CUSTOMERID + "AssyConRul.GenericAssyConnSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\GenericAssyConnSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    strError = "Setting bounded and bounding objects as inputs."
    
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT


    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs", strError).Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add your question below
    
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler

    Dim strError As String
    Dim oAssyConn As StructDetailObjects.AssemblyConn
            
    Const METHOD = m_SelectorProgID & "::SelectorLogic"
    Dim sMsg As String
    
'    MsgBox "GenericAssyConnSel"
    
    ' ---------------------------------
    ' Get the assembly connection ports
    ' ---------------------------------
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    
    Set oAppConnection = pSL.SmartOccurrence
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If lStatus <> 0 Then
        Exit Sub
    End If
        
    ' -------------------------------------------------------------------------------
    ' Verify that the multiple objects have been selected for assembly conn
    ' -------------------------------------------------------------------------------
    
    Dim oConnAttrbs As GSCADSDCreateModifyUtilities.IJSDConnectionAttributes
    Set oConnAttrbs = New SDConnectionUtils
    
    If LessThan(oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Count, 1, 0) Then
        Exit Sub
   End If
    
    Dim bIsTubeBounded As Boolean
    Dim bIsTubeBounding As Boolean
    Dim bIsTubeByTube As Boolean
    
    bIsTubeBounded = False: bIsTubeBounding = False: bIsTubeByTube = False
    
    bIsTubeBounded = IsTubularMember(oBoundedData.MemberPart)
    
    
    ' --------------------------------------------------------------------
    ' Collect information on the relationship between bounding and bounded
    ' --------------------------------------------------------------------
    Dim i As Long
    
    Dim oBoundingObject As Object
    Dim oPortOnBounding As IJPort
    Dim oStructPort As IJStructPort
    
    Dim oPlateObjects As New Collection
    Dim oIsPlateEdge As New Collection
    
    Dim oMemberObjects As New Collection
    Dim oMemberPenetratesWeb As New Collection
    Dim oMemberAlias As New Collection
    
    Dim sectionAlias As Long
    Dim penetratesWeb As Boolean
    Dim oEdgeMap As JCmnShp_CollectionAlias
    
    Set oAssyConn = New AssemblyConn
    Set oAssyConn.object = pSL.SmartOccurrence
    
    Dim oBoundingObj As Object
    Dim bEndToEnd As Boolean, oEndToEndPos As IJDPosition, oTan1 As IJDVector, oTan2 As IJDVector
    
    
    Dim oColl As JCmnShp_CollectionAlias
    
    Set oColl = New Collection
    
    For i = 1 To oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Count
        oColl.Add oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Item(i)
    Next
    
    
    'Adding the bounding port of the assembly connection to AuxiliaryPorts collection
    oColl.Add oAssyConn.Port2
        
    For i = 1 To oColl.Count
        
        Set oPortOnBounding = oColl.Item(i)
        
        Set oBoundingObject = oPortOnBounding.Connectable
                If TypeOf oBoundingObject Is IJProfile Then
                    bIsTubeBounding = IsTubularMember(oBoundingObject)
                End If
                
                If bIsTubeBounded And bIsTubeBounding Then
                    bIsTubeByTube = True
                    Exit For
                End If
                
        If TypeOf oBoundingObject Is IJConnectable Then
            If TypeOf oBoundingObject Is IJPlate Then
                oPlateObjects.Add oBoundingObject
            
                ' presuming here that only one port on the plate has been selected
                If TypeOf oPortOnBounding Is IJStructPort Then
                    Set oStructPort = oPortOnBounding
                    If (oStructPort.ContextID And CTX_LATERAL) Then
                        oIsPlateEdge.Add True
                    Else
                        oIsPlateEdge.Add False
                    End If
                End If
            ElseIf TypeOf oBoundingObject Is ISPSMemberPartPrismatic Then
                Set oEdgeMap = New Collection
                Set oEdgeMap = GetEdgeMap(oAppConnection, oPortOnBounding, oBoundedData.AxisPort, sectionAlias, penetratesWeb)
                
                oMemberObjects.Add oBoundingObject
                oMemberAlias.Add sectionAlias
                oMemberPenetratesWeb.Add penetratesWeb
            ElseIf TypeOf oBoundingObject Is IJProfile Then
                'Treat a profile as a member and add it to the oMemberObjects collection
                Set oEdgeMap = New Collection
                Set oEdgeMap = GetEdgeMap(oAppConnection, oPortOnBounding, oBoundedData.AxisPort, sectionAlias, penetratesWeb)
                oMemberObjects.Add oBoundingObject
                oMemberAlias.Add sectionAlias
                oMemberPenetratesWeb.Add penetratesWeb
            End If
        End If
    Next i
    
    
    If bIsTubeByTube Then
        pSL.Add "Generic_Stiff_AC_01"
    End If
            
    ' -------------------------------------------------------------------
    ' Special case for:
    ' - bounded by plate face and member
    ' - web penetrated
    ' - bounding member has single edge
    ' - position is to edge and outside
    ' -------------------------------------------------------------------
    Dim alias As Long
    Dim topFlangeBoundingEdge As eBounding_Edge
    Dim btmFlangeBoundingEdge As eBounding_Edge
    Dim mappedEdge As JXSEC_CODE
    
    ' One plate and one member
    If oPlateObjects.Count = 1 And oMemberObjects.Count = 1 Then
        
        ' Plate face and pentrates web
        If (oIsPlateEdge.Item(1) = False) And (oMemberPenetratesWeb.Item(1) = True) Then
            
            alias = oMemberAlias.Item(1)
            
            For i = 1 To oColl.Count
               Set oPortOnBounding = oColl.Item(i)
               If TypeOf oPortOnBounding.Connectable Is ISPSMemberPartCommon Or _
                  TypeOf oPortOnBounding.Connectable Is IJProfile Then
                    GetNonPenetratedIntersectedEdge oAppConnection, oPortOnBounding, oBoundedData.AxisPort, btmFlangeBoundingEdge, mappedEdge, True
                    GetNonPenetratedIntersectedEdge oAppConnection, oPortOnBounding, oBoundedData.AxisPort, topFlangeBoundingEdge, mappedEdge, False
               End If
            Next
            ' Bounded bottom to bounding bottom, and top above (edge and outside case 1)
            If (alias = 3) And (btmFlangeBoundingEdge = Bottom_Flange_Right) And (topFlangeBoundingEdge = Above) Then
                pSL.Add "Generic_Stiff_WPEF_PF_1" ' Web-penetrated member edge and face, plate face, configuration #1
            
            ' Bounded top to bounding top, and bottom below (edge and outside case 2)
            ElseIf (alias = 1) And (topFlangeBoundingEdge = Top_Flange_Right) And (btmFlangeBoundingEdge = Below) Then
                pSL.Add "Generic_Stiff_WPEF_PF_1" ' Web-penetrated member edge and face, plate face, configuration #1
            End If
        End If
    End If
    
    ' ----------------------------------
    '
    ' ----------------------------------
    
    If oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Count = 1 And _
        TypeOf oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Item(1) Is ISPSMemberPartPrismatic Then
            pSL.Add "Generic_Stiff_AC_01"
    Else
        pSL.Add "Generic_Stiff_AC_01"
    End If
    
    ' --------------------------------------------------------------------------------
    ' The below smart item is for the following condition
    ' Special case for any two boundaries.  Places a corner feature where the two meet
    ' --------------------------------------------------------------------------------
    If (oPlateObjects.Count + oMemberObjects.Count) = 2 And Not oPlateObjects.Count = 2 Then
        pSL.Add "Generic_Stiff_WP_2B_CF"
    End If
          
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic", strError).Number
End Sub
  

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pSelector.IJDInputs.RemoveAllInput
    pSelector.IJDRepresentations.RemoveAllRepresentation
        
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgID, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************



