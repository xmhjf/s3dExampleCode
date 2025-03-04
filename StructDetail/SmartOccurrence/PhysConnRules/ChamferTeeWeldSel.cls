VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChamferTeeWeldSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_SelectorProgid As String = "PhysConnRules.ChamferTeeWeldSel"
Const m_SelectorName As String = "PhysConnRules.ChamferTeeWeldSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\PhysConnRules\ChamferTeeWeldSel.cls"

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
    
    pQH.SetQuestion "Category", gsNormal, "TeeWeldCategory"
    pQH.SetQuestion "ClassSociety", gsLloyds, "ClassSocietyCol"
    pQH.SetQuestion "BevelAngleMethod", gsConstant, "BevelMethod"
    pQH.SetQuestion "ChamferThickness", "0.004" 'from the chamfer rules
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)

    On Error GoTo ErrorHandler
    
    ' Get Class Arguments
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    
    Set oPhysConn.object = pSLH.SmartOccurrence
    
    ' Get the Question answers
    Dim strCategory As String
    
    strCategory = pSLH.Answer("Category")
    
    '************************************  Selection code - TO DO  ************************************
    '  Note: Convert thickness from meters to millimeters, angles from degrees to radians
    
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Dim oPlatePart As New StructDetailObjects.PlatePart
    Dim oStiffenedPlate As Object
    Dim oParentSystem As Object
    Dim oRootPlate As Object
    Dim oStructDetailUtilHelper As New GSCADStructDetailUtil.StructDetailHelper
    Dim oStructDetailObjectsHelper As New StructDetailObjects.Helper
    
    Dim bNeedToExamineStiffenerAttachment As Boolean
    Dim bIsPlateSystem As Boolean
    
    Dim eAttachmentMethod As JS_TOPOLOGY_STIFF_MOUNTING_POINT
    
    Dim dMountingAngle As Double
    Dim dMountingAngleComplement As Double
    Dim dOpening As Double
    Dim dErrorTeeWeld1 As Double
    Dim dErrorTeeWeld2 As Double
    Dim dErrorTeeWeld3 As Double
    Dim dErrorTeeWeld4 As Double
    Dim dErrorTeeWeld5 As Double
    Dim dErrorTeeWeld6 As Double
    Dim dErrorTeeWeld7 As Double
    
    Const OFF_NORMAL_ANGLE_TOLERANCE As Double = 2#  'Greater than +/- 2 degrees is considered well off normal
    Const ERROR_TOLERANCE As Double = 0.0012                 '1.2 mm
    Const NOSE_SIZE As Double = 0.003                               '3 mm
    
    'If the returned Mounting Angle is greater than 90 degrees, subtract from 180 to get the smaller angle
    dMountingAngle = oPhysConn.MountingAngle
    If (dMountingAngle - TOL) >= DegToRad(90#) Then
        dMountingAngle = Pi - dMountingAngle
    End If
    dMountingAngleComplement = DegToRad(90#) - dMountingAngle
    
    'If this is a physical connection between a stiffener and the plate it is stiffening, and if the mounting
    'angle is well off normal (beyond OFF_NORMAL_ANGLE_TOLERANCE), then stiffener attachment method
    'must be examined in selecting the proper tee weld.
    bNeedToExamineStiffenerAttachment = False
    If oPhysConn.ConnectedObject1Type = SDOBJECT_STIFFENER And _
        oPhysConn.ConnectedObject2Type = SDOBJECT_PLATE Then 'Possibly need to consider long point /
                                                                                                  'short point issue (stiffener attachment)

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
        
        If oStiffenedPlate Is oRootPlate And _
             dMountingAngleComplement > DegToRad(OFF_NORMAL_ANGLE_TOLERANCE) Then  'Mounting angle well off normal - welds dependent
                                                                                                                                       'on attachment method and tee weld category
            bNeedToExamineStiffenerAttachment = True
        End If
    End If
    
    If bNeedToExamineStiffenerAttachment Then
        eAttachmentMethod = oProfilePart.AttachmentMethod

        If eAttachmentMethod = JS_TOPOLOGY_STIFF_NoOverlap Then

            If strCategory = gsNormal Then

                dErrorTeeWeld2 = oProfilePart.WebThickness * Tan(dMountingAngleComplement)
                dErrorTeeWeld3 = dErrorTeeWeld2 / 2#

                'TeeWeld1 allowed in all cases, but must check size of error for TeeWeld2 and TeeWeld3
                If dErrorTeeWeld3 > ERROR_TOLERANCE Then  'Both TeeWeld2 and TeeWeld3 errors > tolerance, so only TeeWeld1 allowed

                    pSLH.Add "TeeWeldChamfer1"

                Else  'TeeWeld3 error <= tolerance, so TeeWeld3 allowed

                    If dErrorTeeWeld2 > ERROR_TOLERANCE Then  'TeeWeld2 not allowed

                        pSLH.Add "TeeWeldChamfer1"
                        pSLH.Add "TeeWeldChamfer3"

                    Else  'TeeWeld2 error <= tolerance, so all allowed

                        pSLH.Add "TeeWeldChamfer1"
                        pSLH.Add "TeeWeldChamfer2"
                        pSLH.Add "TeeWeldChamfer3"

                    End If

                End If

            ElseIf strCategory = gsFull Then

                dErrorTeeWeld5 = oProfilePart.WebThickness / 2# * Tan(dMountingAngleComplement)

                If dErrorTeeWeld5 > ERROR_TOLERANCE Then  'Only TeeWeld4 allowed

                    pSLH.Add "TeeWeldChamfer4"

                Else  'TeeWeld5 error <= tolerance, so both allowed

                    pSLH.Add "TeeWeldChamfer4"
                    pSLH.Add "TeeWeldChamfer5"

                End If

            ElseIf strCategory = gsDeep Then

                dErrorTeeWeld6 = (((oProfilePart.WebThickness - NOSE_SIZE * Cos(dMountingAngleComplement)) / 2#) + _
                                               NOSE_SIZE * Cos(dMountingAngleComplement)) * Tan(dMountingAngleComplement)

                If dErrorTeeWeld6 > ERROR_TOLERANCE Then  'Only TeeWeld7 allowed

                    pSLH.Add "TeeWeldChamfer7"

                Else  'TeeWeld6 error <= tolerance, so both allowed

                    pSLH.Add "TeeWeldChamfer6"
                    pSLH.Add "TeeWeldChamfer7"

                End If

            Else  'Chain or Zigzag -- could we end up here???

                'we have a problem???

            End If

        ElseIf eAttachmentMethod = JS_TOPOLOGY_STIFF_MidPoint Then

            If strCategory = gsNormal Then

                dErrorTeeWeld1 = oProfilePart.WebThickness / 2# * Tan(dMountingAngleComplement)

                If dErrorTeeWeld1 > ERROR_TOLERANCE Then  'Only TeeWeld3 allowed

                    pSLH.Add "TeeWeldChamfer3"

                Else  'TeeWeld1 error <= tolerance, so all allowed

                    pSLH.Add "TeeWeldChamfer1"
                    pSLH.Add "TeeWeldChamfer2"
                    pSLH.Add "TeeWeldChamfer3"

                End If

            ElseIf strCategory = gsFull Then

                dErrorTeeWeld4 = oProfilePart.WebThickness / 2# * Tan(dMountingAngleComplement)

                If dErrorTeeWeld4 > ERROR_TOLERANCE Then  'Only TeeWeld5 allowed

                    pSLH.Add "TeeWeldChamfer5"

                Else  'TeeWeld4 error <= tolerance, so both allowed

                    pSLH.Add "TeeWeldChamfer4"
                    pSLH.Add "TeeWeldChamfer5"

                End If

            ElseIf strCategory = gsDeep Then

                dErrorTeeWeld7 = (oProfilePart.WebThickness / 2# - NOSE_SIZE * Cos(dMountingAngleComplement)) * _
                                            Tan(dMountingAngleComplement)

                If dErrorTeeWeld7 > ERROR_TOLERANCE Then  'Only TeeWeld6 allowed

                    pSLH.Add "TeeWeldChamfer6"

                Else  'TeeWeld7 error <= tolerance, so both allowed

                    pSLH.Add "TeeWeldChamfer6"
                    pSLH.Add "TeeWeldChamfer7"

                End If

            Else  'Chain or Zigzag -- invalid here
                Err.Raise LogError(Err, MODULE, "SelectorLogic", "Invalid Category value encountered.").Number
            End If

        ElseIf eAttachmentMethod = JS_TOPOLOGY_STIFF_MaxOverlap Then

            If strCategory = gsNormal Then

                dErrorTeeWeld1 = oProfilePart.WebThickness * Tan(dMountingAngleComplement)
                dErrorTeeWeld3 = dErrorTeeWeld1 / 2#

                'TeeWeld2 allowed in all cases, but must check size of error for TeeWeld1 and TeeWeld3
                If dErrorTeeWeld3 > ERROR_TOLERANCE Then  'Both TeeWeld1 and TeeWeld3 errors > tolerance, so only TeeWeld2 allowed

                    pSLH.Add "TeeWeldChamfer2"

                Else  'TeeWeld3 error <= tolerance, so TeeWeld3 allowed

                    If dErrorTeeWeld1 > ERROR_TOLERANCE Then  'TeeWeld1 not allowed

                        pSLH.Add "TeeWeldChamfer2"
                        pSLH.Add "TeeWeldChamfer3"

                    Else  'TeeWeld1 error <= tolerance, so all allowed

                        pSLH.Add "TeeWeldChamfer1"
                        pSLH.Add "TeeWeldChamfer2"
                        pSLH.Add "TeeWeldChamfer3"

                    End If

                End If

            ElseIf strCategory = gsFull Then

                dErrorTeeWeld4 = oProfilePart.WebThickness * Tan(dMountingAngleComplement)

                If dErrorTeeWeld4 > ERROR_TOLERANCE Then  'Only TeeWeld5 allowed

                    pSLH.Add "TeeWeldChamfer5"

                Else  'TeeWeld4 error <= tolerance, so both allowed

                    pSLH.Add "TeeWeldChamfer4"
                    pSLH.Add "TeeWeldChamfer5"

                End If

            ElseIf strCategory = gsDeep Then

                dErrorTeeWeld7 = (oProfilePart.WebThickness - NOSE_SIZE * Cos(dMountingAngleComplement)) * _
                                            Tan(dMountingAngleComplement)

                If dErrorTeeWeld7 > ERROR_TOLERANCE Then  'Only TeeWeld6 allowed

                    pSLH.Add "TeeWeldChamfer6"

                Else  'TeeWeld7 error <= tolerance, so both allowed

                    pSLH.Add "TeeWeldChamfer6"
                    pSLH.Add "TeeWeldChamfer7"

                End If

            Else  'Chain or Zigzag -- invalid here
                Err.Raise LogError(Err, MODULE, "SelectorLogic", "Invalid Category value encountered.").Number
            End If

        Else  'Attachment method is UNDEFINED
            Err.Raise LogError(Err, MODULE, "SelectorLogic", "Profile has undefined attachment method").Number
        End If

    Else
        'Don't need to examine stiffener attachment.
        'Either this
        '   physical connection is not between a stiffener and its stiffened plate
        'OR
        '   the mounting angle between stiffener and it's stiffened plate is close to normal
        
        Dim dChamferThickness As Double
        dChamferThickness = pSLH.Answer("ChamferThickness")
        Select Case strCategory
        
            Case gsNormal
                'dOpening = oPhysConn.Object1Thickness * Abs(Tan(dMountingAngleComplement))
                dOpening = dChamferThickness * Abs(Tan(dMountingAngleComplement))
                
                If dOpening < 0.003 Then
                    pSLH.Add "TeeWeldChamfer1"
                ElseIf dOpening >= 0.003 And (dMountingAngle - TOL) > DegToRad(45#) Then
                    pSLH.Add "TeeWeldChamfer2"
                ElseIf dOpening >= 0.003 And (dMountingAngle - TOL) <= DegToRad(45#) Then
                    pSLH.Add "TeeWeldChamfer3"
                End If
                
            Case gsFull
                If (dChamferThickness) < 0.004 Then
                    pSLH.Add "TeeWeldChamfer4"
                Else
                    pSLH.Add "TeeWeldChamfer5"
                End If
                
            Case gsDeep
                If (dChamferThickness) > 0.0025 Then
                    pSLH.Add "TeeWeldChamfer6"
                Else
                    pSLH.Add "TeeWeldChamfer7"
                End If
                
            Case gsChain
                pSLH.Add "ChainWeldChamfer"
                
            Case gsZigZag
                pSLH.Add "ZigZagWeldChamfer"
                
        End Select
    
    End If
        
    Set oProfilePart = Nothing
    Set oPlatePart = Nothing
    Set oStiffenedPlate = Nothing
    Set oParentSystem = Nothing
    Set oRootPlate = Nothing
    Set oStructDetailUtilHelper = Nothing
    Set oStructDetailObjectsHelper = Nothing
    
    '************************************  End Selection code - TO DO  ************************************
    
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



