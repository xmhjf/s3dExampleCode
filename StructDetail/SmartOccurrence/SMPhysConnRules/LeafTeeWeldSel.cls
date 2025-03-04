VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LeafTeeWeldSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2013-16 Intergraph Corporation. All rights reserved.
'
'File : LeafTeeWeldSel.cls
'
'Author : Alligators
'
'Description :
'
'
'History :
'    21/Nov/2013 - svsmylav
'          TR-234194 Earlier selector logic used to avoid PC creation by checking
'                    distance between bounded part and the AC. Fix skips this portion of code for
'                    stiffener-face to plate-face AC.
'
'         DI-259156 Modified selector rule to delete geometry checks
'         which were added to avoid missing PC TDR.
'*****************************************************************************
Option Explicit

Const m_SelectorProgid As String = CUSTOMERID + "PhysConnRules.LeafTeeWeldSel"
Const m_SelectorName As String = CUSTOMERID + "PhysConnRules.LeafTeeWeldSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "PhysConnRules\LeafTeeWeldSel.cls"

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
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)

    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------
    ' Use filter in place of selection logic, if set
    ' ----------------------------------------------
    Dim oParmCM As New PhysConnParmCM
    
    If oParmCM.UsePreSelection(pSLH) Then
        Exit Sub
    End If
    
    ' Get Class Arguments
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    
    Set oPhysConn.object = pSLH.SmartOccurrence

    
    ' Check if this Physical Connection is "Square Trim"
    ' Currently, it is NOT expected that "Square Trim" cases are mixed
    ' ie: One side "Square Trim" but the other is NOT "Square Trim"
    Dim bSquareTrim As Boolean
    Dim bSquareTrim1 As Boolean
    Dim bSquareTrim2 As Boolean
    oPhysConn.HasSquareTrim bSquareTrim1, bSquareTrim2
    If bSquareTrim1 Or bSquareTrim2 Then
        bSquareTrim = True
    Else
        bSquareTrim = False
    End If

    ' Get the Question answers
    Dim strCategory As String
    
    strCategory = pSLH.SelectorAnswer(CUSTOMERID + "PhysConnRules.TeeWeldSel", "Category")
    
    '  Note: Convert thickness from meters to millimeters, angles from degrees to radians
    Dim dObject1Thickness As Double
    Dim dObject2Thickness As Double
    
    dObject1Thickness = 0
    dObject2Thickness = 0
    GetPhysicalConnPartsThickness _
                           pSLH.SmartOccurrence, _
                           dObject1Thickness, _
                           dObject2Thickness

    If dObject1Thickness < 0.000001 Then
       dObject1Thickness = oPhysConn.Object1Thickness
    End If
    If dObject2Thickness < 0.000001 Then
       dObject2Thickness = oPhysConn.Object2Thickness
    End If
    
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
    Dim dGAPFilletWeld1 As Double
    Dim dGAPFilletWeld2 As Double
    Dim dGAPTeeWeldY As Double
    Dim dGAPTeeWeldV As Double
    Dim dGAPTeeWeldX As Double
    Dim dGAPTeeWeldK As Double
       
    Const OFF_NORMAL_ANGLE_TOLERANCE As Double = 2#  'Greater than +/- 2 degrees is considered well off normal
    Const GAP_TOLERANCE As Double = 0.0012                 '1.2 mm
    Const NOSE_SIZE As Double = 0.003                               '3 mm
    
    Dim oHelper As New StructDetailObjects.Helper
    Dim dPortValue As Long

    'If the returned Mounting Angle is greater than 90 degrees, subtract from 180 to get the smaller angle
    dMountingAngle = oPhysConn.MountingAngle
    If (dMountingAngle - TOL) >= DegToRad(90#) Then
        dMountingAngle = PI - dMountingAngle
    End If
    
    'this is the theta angle in the requirements
    dMountingAngleComplement = DegToRad(90#) - dMountingAngle
    
    Dim oStiffener As New StructDetailObjects.ProfilePart
    Dim dProfileThick As Double
    
    'Start with the Plate by Plate cases
    If oPhysConn.ConnectedObject1Type = SDOBJECT_PLATE And _
        oPhysConn.ConnectedObject2Type = SDOBJECT_PLATE Then 'this is a plate to plate case
        
        '************************************************
        'check to see if the welded part is a collar
        'if the feature is a collar, pick the fillet weld or the single bevel
        If TypeOf oPhysConn.ConnectedObject1 Is IJCollarPart Then
            'check if the penetrated plate is tight or watertight
            Dim oCollar As New StructDetailObjects.Collar
            Set oCollar.object = oPhysConn.ConnectedObject1
            
            Dim oCollarPenetrated As New StructDetailObjects.PlatePart
            If TypeOf oCollar.Penetrated Is IJPlatePart Then
                Set oCollarPenetrated.object = oCollar.Penetrated
                If oCollarPenetrated.Tightness = NonTight Then
                    pSLH.Add "FilletWeld1"
                    Exit Sub
                Else 'it is tight
                    pSLH.Add "TeeWeldV"
                    Exit Sub
                End If
            Else
                pSLH.Add "FilletWeld1"
                Exit Sub
            End If
            
        ElseIf bSquareTrim Then
            ' Physical Connection for Material Handling Squrare Trim cases
            ' Note:
            '   For the Bevel data Property Page data to be displayed
            '   The SmartItem must contain "TeeWled" in its name
            '   see \StructDetail\Data\Controls\PhysConnPropPageBevel
            pSLH.Add "TeeWeldSquare"
            
        Else 'the welded plate is not a collar, so continue
             Select Case strCategory
                Case "OneSidedBevel"
                     pSLH.Add "TeeWeldY"
                 
                     'also add the special cases
                     If dObject1Thickness <= 0.025 Then
                         pSLH.Add "TeeWeldV"
                     ElseIf dObject1Thickness > 0.025 Then
                         pSLH.Add "TeeWeldX"
                     End If
                                         
                Case "TwoSidedBevel"
                     pSLH.Add "TeeWeldK"
                     
                Case gsChain
                     pSLH.Add "ChainWeld"
                     
                Case "Staggered"
                     pSLH.Add "StaggeredWeld"
                     
                Case "Chill"
                    pSLH.Add "TeeWeldChill"
                
                Case Else
                    pSLH.Add "TeeWeldV"
                
            End Select
        End If
        
    ElseIf oPhysConn.ConnectedObject1Type = SDOBJECT_PLATE And _
        oPhysConn.ConnectedObject2Type = SDOBJECT_STIFFENER Then 'this is a plate to profile case
        
        'we can assume that this is a bracket that is bounded to the top of a profile
        
        ' Physical Connection for Material Handling should be "TeeWeldSquare"
        ' if the corner Trim is set to open Square Trim cases. Implemented for TR#210195
        If bSquareTrim Then
            pSLH.Add "TeeWeldSquare"
        'check the thickness of the tripping bracket
        ElseIf dObject1Thickness > 0.015 Then  'greater than 15 mm might get a Y bevel
            'check the mounting angle
            If dMountingAngleComplement > DegToRad(30) Then
                pSLH.Add "TeeWeldY"
            Else
                'give both options to let the user choose
                pSLH.Add "FilletWeld1"
                pSLH.Add "TeeWeldY"
            End If
        Else
            'default selection is a fillet weld
            pSLH.Add "FilletWeld1"
        End If
        
    ElseIf oPhysConn.ConnectedObject1Type = SDOBJECT_STIFFENER And _
        oPhysConn.ConnectedObject2Type = SDOBJECT_STIFFENER Then 'this is a profile to profile case
        
        'could either be bounded to web or to flange; check to find out
        dPortValue = oHelper.GetBoundingProfileFace(oPhysConn.Port1, oPhysConn.Port2)
        
        If dPortValue = JXSEC_TOP Or dPortValue = JXSEC_BOTTOM Then
            'default selection is a fillet weld
            pSLH.Add "FilletWeld1"
            'optional selection is a TeeWeldY
            pSLH.Add "TeeWeldY"
        Else 'this profile end is connected to the web of another profile
            Set oStiffener.object = oPhysConn.ConnectedObject1
            
            dProfileThick = oStiffener.webthickness
            
            If dProfileThick < 0.016 Then
                If dMountingAngleComplement <= 0.785398 Then '45 degrees
                    pSLH.Add "FilletWeld2" 'this is the angled version of fillet weld 1
                ElseIf dMountingAngleComplement > 0.785398 And dMountingAngleComplement <= 1.047198 Then '45 to 60 degrees
                    pSLH.Add "TeeWeldV"
                Else
                    pSLH.Add "FilletWeld1"
                End If
            Else
                If dMountingAngleComplement <= 0.261799 Then '< 15 degrees
                    pSLH.Add "FilletWeld2" 'this is the angled version of fillet weld 1
                ElseIf dMountingAngleComplement > 0.261799 And dMountingAngleComplement <= 0.785398 Then '15 to 45 degrees
                    pSLH.Add "TeeWeldY"
                ElseIf dMountingAngleComplement > 0.785398 And dMountingAngleComplement <= 1.047198 Then '45 to 60 degrees
                    pSLH.Add "TeeWeldY"
                End If
            End If
                
        End If

    ElseIf oPhysConn.ConnectedObject1Type = SDOBJECT_STIFFENER And _
        oPhysConn.ConnectedObject2Type = SDOBJECT_PLATE Then 'Possibly need to consider long point /
                                                              'short point issue (stiffener attachment)
        'If this is a physical connection between a stiffener and the plate it is stiffening, and if the mounting
        'angle is well off normal (beyond OFF_NORMAL_ANGLE_TOLERANCE), then stiffener attachment method
        'must be examined in selecting the proper tee weld.
        bNeedToExamineStiffenerAttachment = False

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

    
        If bNeedToExamineStiffenerAttachment Then
            eAttachmentMethod = oProfilePart.AttachmentMethod
    
            ' Physical Connection for Material Handling should be "TeeWeldSquare"
            ' if the corner Trim is set to open Square Trim cases. Implemented for TR#210195
            If bSquareTrim Then
                pSLH.Add "TeeWeldSquare"
            ElseIf eAttachmentMethod = JS_TOPOLOGY_STIFF_NoOverlap Then
    
                If strCategory = gsNormal Or strCategory = "OneSidedBevel" Then
    
                    dGAPFilletWeld2 = oProfilePart.webthickness * Tan(dMountingAngleComplement)
                    dGAPTeeWeldY = dGAPFilletWeld2 / 2#
    
                    'FilletWeld1 allowed in all cases, but must check size of GAP for FilletWeld2 and TeeWeldY
                    If dGAPTeeWeldY > GAP_TOLERANCE Then  'Both FilletWeld2 and TeeWeldY GAPs > tolerance, so only FilletWeld1 allowed
    
                        pSLH.Add "FilletWeld1"
    
                    Else  'TeeWeldY GAP <= tolerance, so TeeWeldY allowed
    
                        If dGAPFilletWeld2 > GAP_TOLERANCE Then  'FilletWeld2 not allowed
    
                            pSLH.Add "FilletWeld1"
                            pSLH.Add "TeeWeldY"
    
                        Else  'FilletWeld2 GAP <= tolerance, so all allowed
    
                            pSLH.Add "FilletWeld1"
                            pSLH.Add "FilletWeld2"
                            pSLH.Add "TeeWeldY"
    
                        End If
    
                    End If
    
                ElseIf strCategory = "TwoSidedBevel" Then
    
                    dGAPTeeWeldX = oProfilePart.webthickness / 2# * Tan(dMountingAngleComplement)
    
                    If dGAPTeeWeldX > GAP_TOLERANCE Then  'Only TeeWeldV allowed
    
                        pSLH.Add "TeeWeldV"
    
                    Else  'TeeWeldX GAP <= tolerance, so both allowed
    
                        pSLH.Add "TeeWeldV"
                        pSLH.Add "TeeWeldX"
    
                    End If
    
                ElseIf strCategory = "OneSidedBevel" Then
    
                    dGAPTeeWeldK = (((oProfilePart.webthickness - NOSE_SIZE * Cos(dMountingAngleComplement)) / 2#) + _
                                                   NOSE_SIZE * Cos(dMountingAngleComplement)) * Tan(dMountingAngleComplement)
    
                    If dGAPTeeWeldK > GAP_TOLERANCE Then  'Only TeeWeldY allowed
    
                        pSLH.Add "TeeWeldY"
    
                    Else  'TeeWeldK GAP <= tolerance, so both allowed
    
                        pSLH.Add "TeeWeldK"
                        pSLH.Add "TeeWeldY"
    
                    End If
    
                Else  'Chain or Zigzag -- could we end up here???
    
                    'we have a problem???
    
                End If
    
            ElseIf eAttachmentMethod = JS_TOPOLOGY_STIFF_MidPoint Then
    
                If strCategory = gsNormal Or strCategory = "OneSidedBevel" Then
    
                    dGAPFilletWeld1 = oProfilePart.webthickness / 2# * Tan(dMountingAngleComplement)
    
                    If dGAPFilletWeld1 > GAP_TOLERANCE Then  'Only TeeWeldY allowed
    
                        pSLH.Add "TeeWeldY"
    
                    Else  'FilletWeld1 GAP <= tolerance, so all allowed
    
                        pSLH.Add "FilletWeld1"
                        pSLH.Add "FilletWeld2"
                        pSLH.Add "TeeWeldY"
    
                    End If
    
                ElseIf strCategory = "TwoSidedBevel" Then
    
                    dGAPTeeWeldV = oProfilePart.webthickness / 2# * Tan(dMountingAngleComplement)
    
                    If dGAPTeeWeldV > GAP_TOLERANCE Then  'Only TeeWeldX allowed
    
                        pSLH.Add "TeeWeldX"
    
                    Else  'TeeWeldV GAP <= tolerance, so both allowed
    
                        pSLH.Add "TeeWeldV"
                        pSLH.Add "TeeWeldX"
    
                    End If
    
                ElseIf strCategory = "OneSidedBevel" Then
    
                    dGAPTeeWeldY = (oProfilePart.webthickness / 2# - NOSE_SIZE * Cos(dMountingAngleComplement)) * _
                                                Tan(dMountingAngleComplement)
    
                    If dGAPTeeWeldY > GAP_TOLERANCE Then  'Only TeeWeldK allowed
    
                        pSLH.Add "TeeWeldK"
    
                    Else  'TeeWeldY GAP <= tolerance, so both allowed
    
                        pSLH.Add "TeeWeldK"
                        pSLH.Add "TeeWeldY"
    
                    End If
    
                Else  'Chain or Zigzag -- invalid here
                    Err.Raise LogError(Err, MODULE, "SelectorLogic", "Invalid Category value encountered.").Number
                End If
    
            ElseIf eAttachmentMethod = JS_TOPOLOGY_STIFF_MaxOverlap Then
    
                If strCategory = gsNormal Or strCategory = "OneSidedBevel" Then
    
                    dGAPFilletWeld1 = oProfilePart.webthickness * Tan(dMountingAngleComplement)
                    dGAPTeeWeldY = dGAPFilletWeld1 / 2#
    
                    'FilletWeld2 allowed in all cases, but must check size of GAP for FilletWeld1 and TeeWeldY
                    If dGAPTeeWeldY > GAP_TOLERANCE Then  'Both FilletWeld1 and TeeWeldY GAPs > tolerance, so only FilletWeld2 allowed
    
                        pSLH.Add "FilletWeld2"
    
                    Else  'TeeWeldY GAP <= tolerance, so TeeWeldY allowed
    
                        If dGAPFilletWeld1 > GAP_TOLERANCE Then  'FilletWeld1 not allowed
    
                            pSLH.Add "FilletWeld2"
                            pSLH.Add "TeeWeldY"
    
                        Else  'FilletWeld1 GAP <= tolerance, so all allowed
    
                            pSLH.Add "FilletWeld1"
                            pSLH.Add "FilletWeld2"
                            pSLH.Add "TeeWeldY"
    
                        End If
    
                    End If
    
                ElseIf strCategory = "TwoSidedBevel" Then
    
                    dGAPTeeWeldV = oProfilePart.webthickness * Tan(dMountingAngleComplement)
    
                    If dGAPTeeWeldV > GAP_TOLERANCE Then  'Only TeeWeldX allowed
    
                        pSLH.Add "TeeWeldX"
    
                    Else  'TeeWeldV GAP <= tolerance, so both allowed
    
                        pSLH.Add "TeeWeldV"
                        pSLH.Add "TeeWeldX"
    
                    End If
    
                ElseIf strCategory = "OneSidedBevel" Then
    
                    dGAPTeeWeldY = (oProfilePart.webthickness - NOSE_SIZE * Cos(dMountingAngleComplement)) * _
                                                Tan(dMountingAngleComplement)
    
                    If dGAPTeeWeldY > GAP_TOLERANCE Then  'Only TeeWeldK allowed
    
                        pSLH.Add "TeeWeldK"
    
                    Else  'TeeWeldY GAP <= tolerance, so both allowed
    
                        pSLH.Add "TeeWeldK"
                        pSLH.Add "TeeWeldY"
    
                    End If
    
                Else  'Chain or Zigzag -- invalid here
                    Err.Raise LogError(Err, MODULE, "SelectorLogic", "Invalid Category value encountered.").Number
                End If
    
            Else  'Attachment method is UNDEFINED
                Err.Raise LogError(Err, MODULE, "SelectorLogic", "Profile has undefined attachment method").Number
            End If
    
        Else  'Don't need to examine stiffener attachment.  Either this physical connection is not between a stiffener and
                'its stiffened plate OR the mounting angle between a stiffener and it's stiffened plate is close to normal
            
            Set oStiffener.object = oPhysConn.ConnectedObject1
            
            dProfileThick = oStiffener.webthickness
            dOpening = dObject1Thickness * Abs(Tan(dMountingAngleComplement))
            
            ' Physical Connection for Material Handling should be "TeeWeldSquare"
            ' if the corner Trim is set to open Square Trim cases. Implemented for TR#210195
            If bSquareTrim Then
                pSLH.Add "TeeWeldSquare"
            ElseIf dOpening < 0.003 Then
                pSLH.Add "FilletWeld1"
            ElseIf dProfileThick < 0.016 Then
                If dMountingAngleComplement <= 0.785398 Then '45 degrees
                    pSLH.Add "FilletWeld2" 'this is the angled version of fillet weld 1
                ElseIf dMountingAngleComplement > 0.785398 And dMountingAngleComplement <= 1.047198 Then '45 to 60 degrees
                    pSLH.Add "TeeWeldV"
                Else
                    pSLH.Add "FilletWeld1"
                End If
            Else
                If dMountingAngleComplement <= 0.261799 Then '< 15 degrees
                    pSLH.Add "FilletWeld2" 'this is the angled version of fillet weld 1
                ElseIf dMountingAngleComplement > 0.261799 And dMountingAngleComplement <= 0.785398 Then '15 to 45 degrees
                    pSLH.Add "TeeWeldY"
                ElseIf dMountingAngleComplement > 0.785398 And dMountingAngleComplement <= 1.047198 Then '45 to 60 degrees
                    pSLH.Add "TeeWeldY"
                End If
            End If
       
        End If
    ElseIf oPhysConn.ConnectedObject1Type = SDOBJECT_BEAM Then
        Dim sSelectedItem As String
        
        sSelectedItem = "FilletWeld1"
        If oPhysConn.ConnectedObject2Type = SDOBJECT_PLATE Then
            ' Similar to profile end bounded by plate,no attachement considered
            Dim oBeamPartWrapper As New StructDetailObjects.BeamPart
            
            Set oBeamPartWrapper.object = oPhysConn.ConnectedObject1
            dProfileThick = oBeamPartWrapper.webthickness
            
            Dim oDesignChild As IJDesignChild
            Dim oPCParent As Object
            Dim oPCParentSF As IJStructFeature
            
            Set oDesignChild = pSLH.SmartOccurrence
            Set oPCParent = oDesignChild.GetParent
            Set oDesignChild = Nothing
            If TypeOf oPCParent Is IJStructFeature Then
                Set oPCParentSF = oPCParent
                If oPCParentSF.get_StructFeatureType = SF_FlangeCut Then
                   dProfileThick = oBeamPartWrapper.flangeThickness
                End If
                Set oPCParentSF = Nothing
            End If
            Set oPCParent = Nothing
            Set oBeamPartWrapper = Nothing
            
            dOpening = dProfileThick * Abs(Tan(dMountingAngleComplement))
            
            If dMountingAngleComplement <= 0.261799 Then ' 15 degrees
                If dOpening < 0.003 Then
                    sSelectedItem = "FilletWeld1"
                Else
                    sSelectedItem = "FilletWeld2"
                End If
            ElseIf dMountingAngleComplement <= 0.785398 Then  ' 15 ~ 45 degrees
                If dProfileThick < 0.016 Then
                    sSelectedItem = "FilletWeld2"
                Else
                    sSelectedItem = "TeeWeldY"
                End If
            ElseIf dMountingAngleComplement <= 1.047198 Then  ' 45 ~ 60 degrees
                If dProfileThick < 0.016 Then
                    sSelectedItem = "TeeWeldV"
                Else
                    sSelectedItem = "TeeWeldY"
                End If
            End If
            
            pSLH.Add sSelectedItem
            
        Else
           ' Add other cases here
           pSLH.Add "FilletWeld1"
           pSLH.Add "FilletWeld2"
           
        End If
        
    Else
        Select Case strCategory
            Case gsNormal
                dOpening = dObject1Thickness * Abs(Tan(dMountingAngleComplement))
                
                If dOpening < 0.003 Then
                   pSLH.Add "FilletWeld1"
                ElseIf dOpening >= 0.003 Then
                    pSLH.Add "FilletWeld2"
                End If
                
            Case gsFull
                If (dObject1Thickness) < 0.004 Then
                    pSLH.Add "TeeWeldChill"
                Else
                    pSLH.Add "TeeWeldV"
                End If
                
            Case gsDeep
                If (dObject1Thickness) > 0.0025 Then
                    pSLH.Add "TeeWeldK"
                Else
                    pSLH.Add "TeeWeldY"
                End If
                
            Case gsChain
                pSLH.Add "ChainWeld"
                
            Case "Staggered"
                pSLH.Add "StaggeredWeld"
                
            Case "OneSidedBevel"
                pSLH.Add "TeeWeldY"
                If dObject1Thickness <= 0.025 Then
                    pSLH.Add "TeeWeldV"
                End If
                
            Case "TwoSidedBevel"
                pSLH.Add "TeeWeldK"
                If dObject1Thickness > 0.025 Then
                    pSLH.Add "TeeWeldX"
                End If
                
            Case "Chill"
                pSLH.Add "TeeWeldChill"
                
        End Select
        
    End If
    
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
    pIH.definition = pSelector
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


