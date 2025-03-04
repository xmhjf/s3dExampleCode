VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketDefCM2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' NOTE: This replaces deprecated BracketDefCM, adding support for other types
'       reinforcement besides just flange

Const MODULE = "S:\StructDetail\Data\SmartOccurrence\BracketRules\BracketDefCM2"


Private sError As String
Private sMETHOD As String

Public Sub Bracket_FinalConstruct( _
    ByVal oAggregatorDescription As IJDAggregatorDescription)

    On Error GoTo ErrorHandler
    sMETHOD = "Bracket_FinalConstruct"
    
    Dim oCAO As Object
    Set oCAO = oAggregatorDescription.CAO
    
    If TypeOf oCAO Is IJSmartPlate Then
    
        ' For Bracket Plate Part:
        ' =======================
        
        ' Get bracket plate
        Dim pBracket As IJSmartPlate
        Set pBracket = oCAO
        
        ' Create the SDCreateModifyUtilities and place bracket
        Dim oSmartPlateUtils As New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
        Dim oBracketOps As GSCADSDCreateModifyUtilities.IJSDSmartPlateOperations
        
        Set oBracketOps = oSmartPlateUtils
        
        oBracketOps.PlaceBracket pBracket
        
        Set pBracket = Nothing
        Set oSmartPlateUtils = Nothing
      
    Else
        sError = "Unexpected Custom Assembly Object"
        GoTo ErrorHandler
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number

End Sub

'***********************************************************************
' METHOD:  CMCreateBracketReinforcement
'
' DESCRIPTION:  Conditional that determines if there is a bracket reinforcement
'
'***********************************************************************
Public Sub CMCreateBracketReinforcement(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    
    sMETHOD = "CMCreateBracketReinforcement"
    
    bIsNeeded = True

    Dim oBracketByPlane As New StructDetailObjects.BracketByPlane
    Dim oBracketPlate As Object
    Dim oBracketContourSO As IJSmartOccurrence
    Set oBracketContourSO = pMD.CAO
    
    Set oBracketPlate = oBracketByPlane.GetBracketPlateFromBracketContourSO(oBracketContourSO)
    Set oBracketByPlane.object = oBracketPlate

    'Get Supports
    Dim pRefPlane As IJPlane
    Dim oUPoint As IJPoint
    Dim oVPoint As IJPoint
    Dim strRootSelector As String
    Dim pSupportColl As IJElements

    oBracketByPlane.GetInputs pRefPlane, oUPoint, oVPoint, strRootSelector, pSupportColl

    Set pRefPlane = Nothing
    Set oUPoint = Nothing
    Set oVPoint = Nothing
    
    Dim strIASelectionRuleRoot As String

    strIASelectionRuleRoot = "IABracketRules_RootBracketByPlaneMTSel2"
    'define interface from which to get the question
    
    Dim pHelper As New StructDetailObjects.Helper
    Dim retValue As Variant
    Dim retValueType As EResultType
    Dim sTableName As String
    Dim lCodeList As Long
    Dim sShortDesc As String
    Dim sLongDesc As String
    Dim sPlaceBracketReinforcement As String
    Dim strBracketType As String
    Dim strIASelectionRule As String
    Dim dLTValue As Double

    'use helper to get the answer from the question to know what type of reinforcement is requested
    Call pHelper.GetCustomAttributeParamValue(oBracketContourSO, strIASelectionRuleRoot, "PlaceReinforcement", _
            retValue, retValueType, sTableName, lCodeList, sShortDesc, sLongDesc)

    'check if answer is from code list; if so, get short desc, if not, use retValue
    If Len(Trim(sTableName)) > 0 Then
        sPlaceBracketReinforcement = sShortDesc
    Else
        sPlaceBracketReinforcement = retValue
    End If
    
    If sPlaceBracketReinforcement = "Flange" Then
        If pSupportColl.Count >= 3 Then
            'Currently not supporting Flange creation for 3S, 4S and 5S brackets
            bIsNeeded = False
        Else
            ' 2 Supports only
            strIASelectionRule = "IABracketRules_BracketByPlaneMT2SSel"
            Call pHelper.GetCustomAttributeParamValue(oBracketContourSO, strIASelectionRule, "BracketByPlaneType", _
                    retValue, retValueType, sTableName, lCodeList, strBracketType, sLongDesc)
                
            'Create Flange only if Bracket Contour is 2SLinearTrimWithFlange
            If strBracketType = "2SLinearTrimWithFlange" Then
                bIsNeeded = True
            Else
                bIsNeeded = False
            End If
        End If
    
        ' Trim the Plate system after resymbolization of the bracket contour
        ' to get the actual final geometry when creating brackets without flange.
        'If it is a flanged bracket(bIsNeeded = True) it is done after creation of flange.
        If bIsNeeded = False Then
            Dim oCAO As Object
            Set oCAO = pMD.CAO
        
            ' Get Bracket-by-Plane Smart Occurrence:
            Dim oBracketByPlaneSO As IJSmartOccurrence
            Set oBracketByPlaneSO = oCAO
        
            ' Create helper for final construction of bracket:
            Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
            Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
            
            'If the Bracket Plate is already flanged, remove the flange.
            Dim oBrPlateSys As IJPlateSystem
            Dim oFlangeAE As Object
            
            Set oBrPlateSys = oBracketUtils.GetBracketByPlaneFromBracketContour(oBracketByPlaneSO)
            
            If Not oBrPlateSys Is Nothing Then
                Set oFlangeAE = oBrPlateSys.FlangeActiveEntity(Nothing)
                
                If Not oFlangeAE Is Nothing Then
                    oBrPlateSys.RemoveFlange Nothing
                    
                    Set oFlangeAE = Nothing
                End If
            End If
                    
            'Trim the plate system with the bracket contour.
            'Since the Bracket-by-Plane Smart Occurrence has access to all
            'the data (plate system), only it as input is needed.
            oBracketUtils.TrimPlateSystemIntoBracket oBracketByPlaneSO
    
            Set oCAO = Nothing
            Set oBrPlateSys = Nothing
        End If
    ElseIf sPlaceBracketReinforcement = "EdgeReinforcement" Then
        If pSupportColl.Count = 2 Then
            strIASelectionRule = "IABracketRules_BracketByPlaneMT2SSel"
        Else ' add in support for more supports as it becomes available
            strIASelectionRule = ""
        End If
        
        ' create only if symbol supports reinforcement information
        If Not strIASelectionRule = "" Then
            Call pHelper.GetCustomAttributeParamValue(oBracketContourSO, strIASelectionRule, "BracketByPlaneType", _
                    retValue, retValueType, sTableName, lCodeList, strBracketType, sLongDesc)
                    
            dLTValue = CalculateLTValue(oBracketContourSO, oBracketPlate)
                
            'Create stiffener only if Bracket Contour with contour for landing curve
            If (strBracketType = "2SLinearTrimReinforce") And (dLTValue > 0) Then
                bIsNeeded = True
            Else
                bIsNeeded = False
            End If
        Else ' no selection rule supported
            bIsNeeded = False
        End If
        
        If bIsNeeded = False Then
            ' add code to remove existing reinforcement here
        End If
    ElseIf sPlaceBracketReinforcement = "BucklingStiffener" Then
        If pSupportColl.Count = 2 Then
            strIASelectionRule = "IABracketRules_BracketByPlaneMT2SSel"
        Else ' add in support for more supports as it becomes available
            strIASelectionRule = ""
        End If
        
        ' create only if symbol supports reinforcement information
        If Not strIASelectionRule = "" Then
            Call pHelper.GetCustomAttributeParamValue(oBracketContourSO, strIASelectionRule, "BracketByPlaneType", _
                    retValue, retValueType, sTableName, lCodeList, strBracketType, sLongDesc)
                
            dLTValue = CalculateLTValue(oBracketContourSO, oBracketPlate)
                
            'Create stiffener only if Bracket Contour with contour for landing curve
            If (strBracketType = "2SLinearTrimReinforce") And (dLTValue > 0) Then
                bIsNeeded = True
            Else
                bIsNeeded = False
            End If
        Else ' no selection rule supported
            bIsNeeded = False
        End If
        
        If bIsNeeded = False Then
            ' add code to remove existing reinforcement here
        End If
    Else
        bIsNeeded = False
    End If
    
    Set oBracketByPlaneSO = Nothing
    Set oBracketUtils = Nothing
    Set pSupportColl = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

Public Sub ConstructBracketReinforcement( _
    ByVal oMemberDescription As IJDMemberDescription, _
    ByVal oResourceManager As IUnknown, _
    ByRef oReinforcement As Object)

    On Error GoTo ErrorHandler
    sMETHOD = "ConstructBracketReinforcement"
    
    ' need answer again to know which type to create
    'define interface from which to get the question
    Dim strIASelectionRuleRoot As String
    strIASelectionRuleRoot = "IABracketRules_RootBracketByPlaneMTSel2"
    
    Dim pHelper As New StructDetailObjects.Helper
    Dim retValue As Variant
    Dim retValueType As EResultType
    Dim sTableName As String
    Dim lCodeList As Long
    Dim sShortDesc As String
    Dim sLongDesc As String
    Dim oBracketContourSO As IJSmartOccurrence
    Set oBracketContourSO = oMemberDescription.CAO

    'use helper to get the answer from the question
    Call pHelper.GetCustomAttributeParamValue(oBracketContourSO, strIASelectionRuleRoot, "PlaceReinforcement", _
            retValue, retValueType, sTableName, lCodeList, sShortDesc, sLongDesc)
    
    Dim sPlaceBracketReinforcement As String

    'check if answer is from code list; if so, get short desc, if not, use retValue
    If Len(Trim(sTableName)) > 0 Then
        sPlaceBracketReinforcement = sShortDesc
    Else
        sPlaceBracketReinforcement = retValue
    End If
    
    ' Get Bracket-by-Plane Smart Occurrence:
    Dim oBracketByPlaneSO As IJSmartOccurrence
    Set oBracketByPlaneSO = oMemberDescription.CAO
    
    ' Create helper for reinforcement creation:
    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
    Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
    
    Dim oBracketPlate As IJPlate
    Set oBracketPlate = oBracketUtils.GetBracketByPlaneFromBracketContour(oBracketByPlaneSO)
    
    If sPlaceBracketReinforcement = "Flange" Then
        
        'First check whether the Flange is already there.
        'If available no need to create another.
        Dim bFlangeSOExists As Boolean
        bFlangeSOExists = False
        
        bFlangeSOExists = CheckFlangeSOExists(oBracketPlate)
        If bFlangeSOExists = False Then
            oBracketUtils.CreateFlange oBracketByPlaneSO, _
                                       "BracketFlangeContour", _
                                       oReinforcement
      
            ' Trim the Plate system after resymbolization of the bracket contour
            ' and resymbolization of Flange contour to get the
            ' actual final geometry when creating flanged brackets.
            oBracketUtils.TrimPlateSystemIntoBracket oBracketByPlaneSO
            
        End If
    ElseIf sPlaceBracketReinforcement = "EdgeReinforcement" Then
        
        Dim bReinforcementExists As Boolean
        Dim oOldER As IJStiffener
        Dim oEdge As Object
        bReinforcementExists = CheckReinforcementExists(oBracketPlate, oOldER)
        
        ' if not yet exists, get free edge on bracket that is not connected
        If (bReinforcementExists = False) Then
    
            ' get edge port for ER
            Dim colPorts As IJElements
            Dim oConnectable As IJConnectable
            Dim oConnections As IJElements
            Set oConnectable = oBracketPlate
            
            oConnectable.enumPorts colPorts, PortEdge
            Set oConnectable = Nothing
                
            Dim oPort As Object
            Dim oIJPort As IJPort
            For Each oPort In colPorts
                Set oIJPort = oPort
                oIJPort.enumConnections oConnections
                If oConnections Is Nothing Then
                    Set oEdge = oPort
                    Exit For
                End If
                Set oConnections = Nothing
            Next
        
        Else ' get inputs and use already known edge
            Dim oERAttr As IJEdgeReinforcementAttributes
            Dim oOldPlate As Object
            Set oERAttr = New GSCADCreateModifyUtilities.CEdgeReinforcementUtils
            oERAttr.GetERCreationInputs oOldER, oOldPlate, oEdge
            Set oOldPlate = Nothing
        End If
        
        If Not oEdge Is Nothing Then
    
            ' get the boundary points from the symbol output
            Dim oIJDSymbol As IJDSymbol
            Dim oBoundRefA As Object
            Dim oBoundRefB As Object
            Dim oBoundSupportA As Object
            Dim oBoundSupportB As Object
            If TypeOf oBracketContourSO Is IJDSymbol Then
                Set oIJDSymbol = oBracketContourSO
                
                On Error Resume Next ' BindToOutput throws error if output not found
                Set oBoundRefA = oIJDSymbol.BindToOutput("ERPoints", "Output_ER_Point_BoundPointA")
                Set oBoundRefB = oIJDSymbol.BindToOutput("ERPoints", "Output_ER_Point_BoundPointB")
                Set oBoundSupportA = oIJDSymbol.BindToOutput("ER_Boundary1", "Output_ER_Boundary_1")
                Set oBoundSupportB = oIJDSymbol.BindToOutput("ER_Boundary2", "Output_ER_Boundary_2")
                On Error GoTo ErrorHandler
                
                ' create ER on bracket
                Dim oIJERDefinition As IJEdgeReinforcementDefinition
                Set oIJERDefinition = New GSCADCreateModifyUtilities.CEdgeReinforcementUtils
                
                On Error GoTo ErrorHandler ' reset to handle placement errors
                Set oReinforcement = oIJERDefinition.PlaceEdgeReinforcement( _
                                                 oResourceManager, _
                                                 oBracketPlate, _
                                                 oEdge, _
                                                 oOldER)
                SetCrossSection oReinforcement, "FB", "F 100x10", BottomSideOfInferiorFlange, 1
                
                ' replace with default naming rule
                Dim oNamedItem As IJNamedItem
                Dim oNamedBracket As IJNamedItem
                Set oNamedBracket = oBracketPlate
                Set oNamedItem = oReinforcement
                oNamedItem.Name = oNamedBracket.Name & "-EdgeReinforcement"
                
                ' add boundaries using boundary points here using profile utils
                ' must have 2 boundaries for ER to be correctly bound and not ambiguous
                Dim oBoundUtil As IJProfileOperations
                Dim oBoundColl As Collection
                Dim oSupport As Object
                
                If Not oBoundRefA Is Nothing And Not oBoundRefB Is Nothing Or _
                   Not oBoundRefA Is Nothing And Not oBoundSupportA Is Nothing Or _
                   Not oBoundSupportA Is Nothing And Not oBoundSupportB Is Nothing Then
                    
                    Set oBoundColl = New Collection
                    If Not oBoundRefA Is Nothing Then
                        oBoundColl.Add oBoundRefA
                    End If
                    If Not oBoundSupportA Is Nothing Then
                        Set oSupport = GetSupportFromIndex(oBracketContourSO, oBracketUtils, oBoundSupportA)
                        oBoundColl.Add oSupport
                    End If
                    If Not oBoundRefB Is Nothing Then
                        oBoundColl.Add oBoundRefB
                    End If
                    If Not oBoundSupportB Is Nothing Then
                        Set oSupport = GetSupportFromIndex(oBracketContourSO, oBracketUtils, oBoundSupportB)
                        oBoundColl.Add oSupport
                    End If

                    ' TEMPORARILY comment out until boundary issues resolved
'                    If oBoundColl.Count = 2 Then
'                        Set oBoundUtil = New GSCADCreateModifyUtilities.ProfileUtils
'                        oBoundUtil.ApplyProfileSystemBoundaries oReinforcement, oBoundColl
'                    End If
                    
                ' else no boundaries in symbol
                End If
            End If
        End If
        
    ElseIf sPlaceBracketReinforcement = "BucklingStiffener" Then
    
        Dim oOldBuckStiff As IJStiffener
        bReinforcementExists = CheckReinforcementExists(oBracketPlate, oOldBuckStiff)
        
        '!!!Is there really any way to know not to edit again since this gets called
        '   several times in the sequence of creating the bracket via the UI?
        
'        If (bReinforcementExists = False) Then
        
            Dim oBSContour As Object
            If TypeOf oBracketContourSO Is IJDSymbol Then
                Set oIJDSymbol = oBracketContourSO
                
                On Error Resume Next ' BindToOutput throws error if output not found
                Set oBSContour = oIJDSymbol.BindToOutput("BucklingStiffContour", "Output_BucklingStiffContour")
                On Error GoTo ErrorHandler
                
                ' using the contour, create a stiffener by projection on the bracket
                Dim oProfileDefn As IJProfileDefinition
                Set oProfileDefn = New GSCADCreateModifyUtilities.ProfileUtils
                
                Dim oBucklingStiffener As IJStiffener
                Dim oBracketPlane As IJPlane
                
                If Not oBSContour Is Nothing Then
                    Set oBracketPlane = oBracketPlate
                    Set oBucklingStiffener = oProfileDefn.PlaceStiffener_Projected(sptLongitudinal, oResourceManager, _
                        oBracketPlate, oBSContour, oBracketPlane, oOldBuckStiff)
                    If Not oBucklingStiffener Is Nothing Then
                        SetCrossSection oBucklingStiffener, "FB", "F 100x10", BottomSideOfInferiorFlange, 1
                        
                        Dim oStiffAttach As IJStiffenerAttachmentMethod
                        Set oStiffAttach = oBucklingStiffener
                        oStiffAttach.AttachmentMethod = JS_TOPOLOGY_STIFF_NoOverlap
                        oStiffAttach.LandCrvProjectionMethod = SS_LCPM_Normal
                        
                        ' replace with default naming rule
                        Set oNamedBracket = oBracketPlate
                        Set oNamedItem = oBucklingStiffener
                        oNamedItem.Name = oNamedBracket.Name & "-BucklingStiffener"
                        
                    End If
                End If
            End If
'        End If
    End If
    Set oBracketPlate = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number

End Sub

Public Sub DeleteBracketReinforcement( _
    ByVal oReinforcementMember As IJDMemberDescription _
)

    On Error GoTo ErrorHandler
    'Get Plate System, then RemoveFlange.
    
    'Previously did nothing for flange.  Question is does reinforcement need to
    'be deleted here, or like flange be deleted in CMCreate when bIsNeeded is false
    sMETHOD = "DeleteBracketReinforcement"
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number

End Sub

Private Function CheckFlangeSOExists(oBracketPlateSys As Object) As Boolean
    Const sMETHOD = "CheckFlangeSOExists"
    On Error GoTo ErrorHandler
    
    'Initialize the function output
    CheckFlangeSOExists = False
    
    Dim oBracketPlate As IJPlateSystem
    Dim oFlangeAE As IJPlateFlange_AE
    Dim oFlange As Object
    Dim oFlangeSymbol As Object
                
    Set oBracketPlate = oBracketPlateSys
    
    Set oFlange = oBracketPlate.FlangeActiveEntity(Nothing)
    If Not oFlange Is Nothing Then
        Set oFlangeAE = oFlange
        Set oFlangeSymbol = oFlangeAE.FlangeSymbol
    End If
        
    If Not oFlangeSymbol Is Nothing Then
        'Flange already exists
        CheckFlangeSOExists = True
    End If
    
cleanup:
    Set oFlangeSymbol = Nothing
    Set oFlangeAE = Nothing
    Set oFlange = Nothing
    Set oBracketPlate = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
    GoTo cleanup
End Function

Private Function CheckReinforcementExists(oBracketPlateSys As Object, oOldStiff As IJStiffener) As Boolean

    On Error Resume Next
    Dim exists As Boolean
    exists = False
    
    Dim oBracketSys As IJSystem
    Dim oChildCol As IJDTargetObjectCol
    Dim oChild As Object
    
    Set oBracketSys = oBracketPlateSys
    Set oChildCol = oBracketSys.GetChildren
    If Not oChildCol Is Nothing Then
        If oChildCol.Count > 0 Then
            Dim i As Integer
            For i = 1 To oChildCol.Count
            
                ' temporarily just check for stiffener, assume ER or BS
                ' may need more explicit check in the future
                Set oChild = oChildCol.Item(i)
                If TypeOf oChild Is IJStiffener Then
                    exists = True
                    Set oOldStiff = oChild
                    Exit For
                End If
            Next
        End If
    End If
    
    CheckReinforcementExists = exists
End Function

'NOTE: Handling of cross section may change depending on further discussions with user/requirements writers

' Function:
'    CreateCrossSection
' Abstract:
'   Create a Cross section to be applied to the profile
' Description:
'   Creates a hard code profile section
Private Function CreateCrossSection(ByVal strSectionType As String, _
                                    ByVal strSectionName As String) As IJCrossSection
                                                      
    'Dim oRefDataQuery As GSCADRefDataServices.StructQuery
    Dim oRefDataService As RefDataMiddleServices.StructServices
    Dim xsecColl As IJDCollection
    Dim oTempCrossSection As IJDAttributes
    Dim lSize As Long
    Dim collIndex As Integer
    Dim strName As String
    On Error GoTo ErrorHandler
    
    'Set oRefDataQuery = New GSCADRefDataServices.StructQuery
    Set oRefDataService = New RefDataMiddleServices.StructServices
    Set xsecColl = oRefDataService.GetCrossSections("ShipShapes", strSectionType)
    Set oRefDataService = Nothing
    
    lSize = xsecColl.Size
    
    Dim i As Integer
    For i = 1 To lSize
        Set oTempCrossSection = xsecColl.Item(i)
        If oTempCrossSection.CollectionOfAttributes("ISTRUCTCrossSection").Item("SectionName").Value = strSectionName Then
            collIndex = i
            Exit For
        End If
    Next i
    
    Set CreateCrossSection = xsecColl.Item(collIndex)
    Set xsecColl = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
              
Private Sub SetCrossSection(oProfile As Object, _
                            ByVal strCrossSectionType As String, _
                            ByVal strCrossSectionName As String, _
                            ByVal eMountingFace As ProfileFaceName, _
                            ByVal loadPoint As Long)
                           
    Dim oIJProfile As IJProfile
    Dim oIJProfileAttributes As IJProfileAttributes
    Dim oCrossSection As IJCrossSection
    
    On Error GoTo ErrorHandler
    
    Set oIJProfile = oProfile
'    If oIJProfile Is Nothing Then
'        Err.Raise E_INVALIDARG, MODULE & ":" & "SetCrossSection", "Profile not found"
'    End If
    
    Set oIJProfileAttributes = New GSCADCreateModifyUtilities.ProfileUtils
    Set oCrossSection = CreateCrossSection(strCrossSectionType, strCrossSectionName)
    If oCrossSection Is Nothing Then
        Err.Raise E_FAIL, MODULE & ":" & "SetCrossSection", "Failed to get CrossSection"
    End If
    
    oIJProfileAttributes.SetCrossSection oIJProfile, oCrossSection, eMountingFace, loadPoint
    
    Set oIJProfile = Nothing
    Set oIJProfileAttributes = Nothing
    Set oCrossSection = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Function CalculateLTValue(oBracketSO As IJSmartOccurrence, oBracketPlate As IJPlate) As Double

    Dim oUnsupportedEdge As Object
    Dim oIJDSymbol As IJDSymbol
    
    CalculateLTValue = 0 ' initialize output
    
    If TypeOf oBracketSO Is IJDSymbol Then
        Set oIJDSymbol = oBracketSO
        
        On Error Resume Next
        Set oUnsupportedEdge = oIJDSymbol.BindToOutput("UnsupportedEdgeLength", "Output_UnsupportedEdgeLength")
        If Not oUnsupportedEdge Is Nothing Then
        
            ' calculate the L/T to use to decide whether reinforcement is required
            Dim dThickness As Double
            dThickness = oBracketPlate.thickness
            If dThickness > 0 Then
            
                '!!!how to get the length of a contour
                Dim oCurve As IJCurve
                If TypeOf oUnsupportedEdge Is IJCurve Then
                    Set oCurve = oUnsupportedEdge
                    Dim dEdgeLength As Double
                    dEdgeLength = oCurve.Length
                    CalculateLTValue = dEdgeLength / dThickness
'                Else
'                    MsgBox "CalculateLTValue - contour is NOT an IJCurve"
                End If
            End If
        End If
    End If

End Function

Private Function GetSupportFromIndex(oBracketContourSO As IJSmartOccurrence, _
                                     oBracketUtils As IJBracketAttributes, _
                                     oParameter As IJDParameterContent) As Object
    On Error Resume Next
    'translate the support # into the object to use
    Dim oSupports As IJElements
    Dim nSupportCount As Long
    oBracketUtils.GetSupportsFromBracketContourSO oBracketContourSO, oSupports, nSupportCount
    
    ' note that elememts is 0-based so subtract from support number to get actual index
    Dim oSupport As Object
    Set oSupport = oSupports.Item(oParameter.UomValue - 1)
    
'debugging
'If Not oSupport Is Nothing Then
'    Dim oSupportName As IJNamedItem
'    If TypeOf oSupport Is IJNamedItem Then
'        Set oSupportName = oSupport
'        MsgBox "GetSupportFromIndex - support found=" & oSupportName.Name
'        Set oSupportName = Nothing
'    End If
'End If
    Set GetSupportFromIndex = oSupport

End Function


