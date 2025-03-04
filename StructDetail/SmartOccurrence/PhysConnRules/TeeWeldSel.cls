VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TeeWeldSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_SelectorProgid As String = "PhysConnRules.TeeWeldSel"
Const m_SelectorName As String = "PhysConnRules.TeeWeldSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\PhysConnRules\TeeWeldSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)

    On Error GoTo ErrorHandler
    
    'Add port inputs
    pIH.SetInput INPUT_CONN_OBJECT1
    pIH.SetInput INPUT_CONN_OBJECT2
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)

    On Error GoTo ErrorHandler
    
    pQH.SetQuestion "Category", gsNormal, "TeeWeldCategory", "SetCategory", "PhysConnRules.TeeWeldSel"
    pQH.SetQuestion "ClassSociety", gsLloyds, "ClassSocietyCol", "SetClassSociety", "PhysConnRules.TeeWeldSel"
    pQH.SetQuestion "BevelAngleMethod", gsConstant, "BevelMethod", "SetBevelAngleMethod", "PhysConnRules.TeeWeldSel"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)

    On Error GoTo ErrorHandler
    
    ' Get Class Arguments
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    
    Set oPhysConn.object = pSLH.SmartOccurrence
        
    ' Customer data area starts,change it accordingly as needed
    ' Set mounting angle thresholds
    Dim arrayAngleThreshold() As Double
    Dim bCheckToBeSplit As Boolean
    Dim oSysChild As IJSystemChild
    Dim oPCParent As Object
    
    bCheckToBeSplit = True
    Set oSysChild = oPhysConn.object
    Set oPCParent = oSysChild.GetParent
    Set oSysChild = Nothing
    If Not oPCParent Is Nothing Then
        If TypeOf oPCParent Is IJStructPhysicalConnection Then
            bCheckToBeSplit = False
        End If
        Set oPCParent = Nothing
    End If
    
    Dim bToBeSplit As Boolean
    
    bToBeSplit = False
    If bCheckToBeSplit = True Then
        Dim bAngleComputed As Boolean
        Dim strCategory As String
        
        strCategory = pSLH.Answer("Category")
        bAngleComputed = ComputeSplitAngles(strCategory, oPhysConn, arrayAngleThreshold)
        If bAngleComputed Then
            oPhysConn.IsPhysConnToBeSplit arrayAngleThreshold, bToBeSplit
        End If
    End If
    
    If bToBeSplit = True Then
        pSLH.Add "RootTeeWeld"
    Else
        pSLH.Add "LeafTeeWeld"
    End If
                
    Set oPhysConn = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
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
    
    Dim pSelectorFact As New DefinitionFactory
    pSelectorFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.Definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, _
                                                                                           ByVal DP As Variant, _
                                                                                           ByVal pRM As Object) As Object
    Dim pSelectorFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pSelectorFact.InstanciateSelector(m_SelectorProgid, CB, _
                                                                                IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, _
                                                                                       ByVal pRepName As String, _
                                                                                       ByVal pOutputColl As Object, _
                                                                                       arrayOfInputs() As Variant)
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, _
                                                                                  ByVal pTransactionMgr As Object) As Boolean
End Function

Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSLH As IJDSelectorLogic
    Set pSLH = New SelectorLogic
    pSLH.Representation = pRep
    SelectorLogic pSLH
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

Private Function ComputeSplitAngles(strCategory As String, oPhysConn As StructDetailObjects.PhysicalConn, splitAngles() As Double) As Boolean
    On Error GoTo ErrorHandler
    
     Dim oProfilePart As New StructDetailObjects.ProfilePart
     Dim oPlatePart As New StructDetailObjects.PlatePart
     Dim oStiffenedPlate As Object
     Dim oParentSystem As Object
     Dim oRootPlate As Object
     Dim oStructDetailUtilHelper As New GSCADStructDetailUtil.StructDetailHelper
     Dim oStructDetailObjectsHelper As New StructDetailObjects.Helper
    
     Dim bIsPlateSystem As Boolean
     Dim bSplitAnglesComputed As Boolean
     bSplitAnglesComputed = False
              
     Dim eAttachmentMethod As JS_TOPOLOGY_STIFF_MOUNTING_POINT
     Dim firstAngle As Double
     Dim webthickness As Double
    
    'The values for ERROR_TOLERANCE and NOSE_SIZE need to match the values used in the leaf selector
    'They should be redefined in a common .bas file
    Const ERROR_TOLERANCE As Double = 0.0012                 '1.2 mm
    Const NOSE_SIZE As Double = 0.003                        '3 mm
    Const MAX_ROOT_GAP As Double = 0.003                     '3 mm
    
    ' We will only split the physical connection between a bounded stiffener
    ' and a bounding plate if the plate is on the system stiffened by the
    ' stiffener
    If oPhysConn.ConnectedObject1Type = SDOBJECT_STIFFENER Then
        If oPhysConn.ConnectedObject2Type = SDOBJECT_PLATE Then
            'This connection may need to be split
            'Now, is connected object 2 the stiffened plate of connected object 1?
            Set oProfilePart.object = oPhysConn.ConnectedObject1
            Set oPlatePart.object = oPhysConn.ConnectedObject2
            
            oProfilePart.GetStiffenedPlate oStiffenedPlate, bIsPlateSystem
              
            On Error Resume Next
            
            oStructDetailUtilHelper.IsPartDerivedFromSystem oPlatePart.object, oParentSystem
            
            If Not oParentSystem Is Nothing Then  'Need to get attachment method from parent profile system
                Set oRootPlate = oStructDetailObjectsHelper.Object_RootParentSystem(oPlatePart.object)
            Else
                Set oRootPlate = oPlatePart.object
            End If
            
            If oStiffenedPlate Is oRootPlate Then
                'this is the connection between stiffener and stiffened plate
                'need to compute split angles base on attachment method and weld category
                
                eAttachmentMethod = oProfilePart.AttachmentMethod
                webthickness = oProfilePart.webthickness
        
                If eAttachmentMethod = JS_TOPOLOGY_STIFF_NoOverlap Then
                    If strCategory = gsNormal Then
                        ' we will have two split angle values - EQ1 and EQ2 from analysis
                        ReDim splitAngles(1)
                        splitAngles(0) = Atn(ERROR_TOLERANCE / webthickness) ' EQ2
                        splitAngles(1) = Atn(2# * (ERROR_TOLERANCE / webthickness)) ' EQ1
                        bSplitAnglesComputed = True
        
                    ElseIf strCategory = gsFull Then
                        ' we will have one split angle value - EQ1 from analysis
                        ReDim splitAngles(0)
                        splitAngles(0) = Atn(2# * (ERROR_TOLERANCE / webthickness)) ' EQ1
                        bSplitAnglesComputed = True
                    ElseIf strCategory = gsDeep Then
                        ' we will have one split angle value - EQ3 from analysis
                        ' iterative solution, but we only solve twice -- assume it will converge
                        ' well enough in two passes
                        ReDim splitAngles(0)
                        ' for first pass, use NOSE_SIZE
                        firstAngle = Atn(2# * ERROR_TOLERANCE / (webthickness + NOSE_SIZE))  'EQ3a
                        ' for second pass, use NOSE_SIZE*cos(firstAngle)
                        splitAngles(0) = Atn(2# * ERROR_TOLERANCE / (webthickness + NOSE_SIZE * Cos(firstAngle)))  'EQ3b
                        bSplitAnglesComputed = True
                    Else  'Chain or Zigzag -- could we end up here???
        
                        'we have a problem???
        
                    End If
        
                ElseIf eAttachmentMethod = JS_TOPOLOGY_STIFF_MidPoint Then
        
                    If strCategory = gsNormal Or strCategory = gsFull Then
                         ' we will have one split angle value - EQ1 from analysis
                        ReDim splitAngles(0)
                        splitAngles(0) = Atn(2# * (ERROR_TOLERANCE / webthickness)) ' EQ1
                        bSplitAnglesComputed = True
                    ElseIf strCategory = gsDeep Then
                        If (((webthickness / 2#) - NOSE_SIZE) > 0.00001) Then
                            ' we will have one split angle value - EQ4 from analysis
                            ' iterative solution, but we only solve twice -- assume it will converge
                            ' well enough in two passes
                            ReDim splitAngles(0)
                            ' for first pass, use NOSE_SIZE
                            firstAngle = Atn(ERROR_TOLERANCE / ((webthickness / 2#) - NOSE_SIZE)) 'EQ4a
                            ' for second pass, use NOSE_SIZE*cos(firstAngle)
                            splitAngles(0) = Atn(ERROR_TOLERANCE / ((webthickness / 2#) - NOSE_SIZE * Cos(firstAngle)))  'EQ4b
                            bSplitAnglesComputed = True
                        Else
                            'this is just a safety check.  The user would never want a deep
                            ' penetration weld where the nose size is nearly equal to the
                            ' member thickness.  Just need to prevent a divide by zero if
                            ' this would occur.
                            bSplitAnglesComputed = False
                        End If
                    Else  'Chain or Zigzag -- invalid here
                        Err.Raise LogError(Err, MODULE, "SelectorLogic", "Invalid Category value encountered.").Number
                    End If
        
                ElseIf eAttachmentMethod = JS_TOPOLOGY_STIFF_MaxOverlap Then
        
                    If strCategory = gsNormal Then
                        ' we will have two split angle values - EQ1 and EQ2 from analysis
                        ReDim splitAngles(1)
                        splitAngles(0) = Atn(ERROR_TOLERANCE / webthickness) ' EQ2
                        splitAngles(1) = Atn(2# * (ERROR_TOLERANCE / webthickness)) ' EQ1
                        bSplitAnglesComputed = True
        
                    ElseIf strCategory = gsFull Then
                        ' we will have one split angle value - EQ2 from analysis
                        ReDim splitAngles(0)
                        splitAngles(0) = Atn(ERROR_TOLERANCE / webthickness) ' EQ2
                        bSplitAnglesComputed = True
        
                    ElseIf strCategory = gsDeep Then
                        If ((webthickness - NOSE_SIZE) > 0.00001) Then
                            ' we will have one split angle value - EQ5 from analysis
                            ' iterative solution, but we only solve twice -- assume it will converge
                            ' well enough in two passes
                            ReDim splitAngles(0)
                            ' for first pass, use NOSE_SIZE
                            firstAngle = Atn(ERROR_TOLERANCE / (webthickness - NOSE_SIZE))  'EQ5a
                            ' for second pass, use NOSE_SIZE*cos(firstAngle)
                            splitAngles(0) = Atn(ERROR_TOLERANCE / (webthickness - NOSE_SIZE * Cos(firstAngle)))  'EQ5b
                            bSplitAnglesComputed = True
                        Else
                            'this is just a safety check.  The user would never want a deep
                            ' penetration weld where the nose size is nearly equal to the
                            ' member thickness.  Just need to prevent a divide by zero if
                            ' this would occur.
                            bSplitAnglesComputed = False
                        End If
        
                    Else  'Chain or Zigzag -- invalid here
                        Err.Raise LogError(Err, MODULE, "SelectorLogic", "Invalid Category value encountered.").Number
                    End If
        
                Else  'Attachment method is UNDEFINED
                    Err.Raise LogError(Err, MODULE, "SelectorLogic", "Profile has undefined attachment method").Number
                End If
            End If
        End If
        
    ElseIf oPhysConn.ConnectedObject1Type = SDOBJECT_PLATE Then
    
        'The bounded part is a plate, only need to split based on angles if weld category is normal
        If strCategory = gsNormal Then
        
            'first criteria occurs when the root opening hits 3 mm (.003 M)
            ' this is dependent on the plate thickness
            Dim rootGapAngle As Double
            Dim plateThickness As Double
            plateThickness = oPhysConn.Object1Thickness
            ReDim splitAngles(1)
            rootGapAngle = Atn(MAX_ROOT_GAP / plateThickness)
            If (rootGapAngle < DegToRad(45#)) Then
                splitAngles(0) = rootGapAngle
                splitAngles(1) = DegToRad(45#)
            Else
                ' this does not seem likely
                splitAngles(0) = DegToRad(45#)
                splitAngles(1) = rootGapAngle
            End If
            bSplitAnglesComputed = True
        End If
    End If
   
    ComputeSplitAngles = bSplitAnglesComputed
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ComputeSplitAngles").Number
    ComputeSplitAngles = bSplitAnglesComputed

End Function
Public Sub SetClassSociety(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class

    On Error GoTo ErrorHandler

    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Set oInputDG = pInput

    'Pass the Parent PC answer value to the leaf PC
    SetAnswerFromParentPhysicalConnection oInputDG.Definition, "ClassSociety"

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetClassSociety").Number
End Sub

Public Sub SetCategory(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Set oInputDG = pInput
    
    'Pass the Parent PC answer value to the leaf PC
    SetAnswerFromParentPhysicalConnection oInputDG.Definition, "Category"
        
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetCategory").Number
End Sub

Public Sub SetBevelAngleMethod(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler

    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Set oInputDG = pInput

    'Pass the Parent PC answer value to the leaf PC
    SetAnswerFromParentPhysicalConnection oInputDG.Definition, "BevelAngleMethod"

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetBevelAngleMethod").Number
End Sub


