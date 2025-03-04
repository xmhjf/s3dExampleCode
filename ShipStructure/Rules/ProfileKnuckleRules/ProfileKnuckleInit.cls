VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileKnuckleInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------------
'  Copyright (C) 2004, Intergraph Corporation.  All rights reserved.
'
'  FILE:  ProfileKnuckleInit.cls
'
'  DESCRIPTION:
'  Rule to set the manufacturing mrthod of a Profile Knuckle and if the
'  Profile Knuckle is a 'bend' one, also set the 'bend' properties.
'
'  AUTHOR:  Alain Falasse
'
'  HISTORY
'  04/17/2005    Alain Falasse  Creation
'  04/19/2005    Maria White    Modification after interface change
'-------------------------------------------------------------------------------

Option Explicit

Implements IJProfileKnuckleRule
Implements IJProfileKnuckleRuleEx

'-------------------------------------------------------------------------------
' Instance data
'-------------------------------------------------------------------------------

Private m_eManufacturingMethod As ProfileKnuckleManufacturingMethod
Private m_dInnerRadius As Double

'-------------------------------------------------------------------------------
Private Function ManufacturingMethod(oProfile As Object, oKnuckle As Object) As ProfileKnuckleManufacturingMethod

    ' Decide of manufacturing method here.
    ' By default we set it to 'Split'
    m_eManufacturingMethod = pkmmSplit
    
'*** The code below is comented out because of TR 87231
'*** It must be uncommented when bent plates are supported and DI 87237
'*** is implemented.
'    ' Determine if cross-section is "Flat Bar"
'    Dim oSection As IJDProfileSection
'    Set oSection = oProfile
'    If oSection.crossSection.Type = "FB" Then
'        ' Determine if there is a related plate knuckle
'        Dim oKnuckleRefs As IJProfileKnuckleRefs
'        Dim oPlateKnuckle As IJKnuckle
'
'        Set oKnuckleRefs = oKnuckle
'        Set oPlateKnuckle = oKnuckleRefs.PlateKnuckle
'        If Not oPlateKnuckle Is Nothing Then
'            If oPlateKnuckle.KnuckleType = eKT_Bent Then
'                m_eManufacturingMethod = pkmmBend
'            End If
'        End If
'    End If
    
    ManufacturingMethod = m_eManufacturingMethod
    
End Function
'-------------------------------------------------------------------------------
Private Function InnerRadius(oKnuckle As Object) As Double

    ' Decide of inner radius of a Profile Knuckle of manufacturing method
    ' 'pkmmBend'. For 'pkmmSplit' or 'pkmmIgnore', this value has no
    ' importance. However, it should be >= 0 or the Profile Knuckle will
    ' generate an error.
    m_dInnerRadius = 0#   'initialize to 0
    'InnerRadius = 0.001
    
    
'*** The code below is comented out because of TR 87231
'*** It must be uncommented when bent plates are supported and DI 87237
'*** is implemented.
'    If m_eManufacturingMethod = pkmmBend Then
'        Dim oKnuckleRefs As IJProfileKnuckleRefs
'        Dim oPlateKnuckle As IJKnuckle
'
'        Set oKnuckleRefs = oKnuckle
'        Set oPlateKnuckle = oKnuckleRefs.PlateKnuckle
'        If Not oPlateKnuckle Is Nothing Then
'            InnerRadius = oPlateKnuckle.InnerRadius
'        End If
'
'    End If

    
' ***   The innerRadius radius will always be determined for every manufacturing type
' ***   Skip MfgMethod check

'    If TypeOf oKnuckle Is IJProfileKnuckleMfg Then
        Dim oProfileKnuckleMfg As IJProfileKnuckleMfg
        Dim eExistingMfgMethod As ProfileKnuckleManufacturingMethod
        eExistingMfgMethod = pkmmSplit
        
        On Error Resume Next
        Set oProfileKnuckleMfg = oKnuckle
        
        If Not oProfileKnuckleMfg Is Nothing Then
            eExistingMfgMethod = oProfileKnuckleMfg.ManufacturingMethod
            Set oProfileKnuckleMfg = Nothing
        End If
  
        'If eExistingMfgMethod = pkmmBend Or eExistingMfgMethod = pkmmBendPlusFeature Then
            Dim dProfileHeight As Double
            Dim sCrossSectionType As String
            
            Dim oKnuckledProfilePart As Object
            Dim bIsConvexKnuckle As Boolean
            
            'bIsConvexKnuckle should be always false unless some cases for pkmmBendPlusFeature
            bIsConvexKnuckle = False
            
            GetKnuckleProfilePart oKnuckle, eExistingMfgMethod, oKnuckledProfilePart, bIsConvexKnuckle
            If Not oKnuckledProfilePart Is Nothing Then
                Dim oSDOProfilePart As IJSDOProfilePart
                Set oSDOProfilePart = New StructDetailObjectsEx.SDOProfilePart
                If oSDOProfilePart Is Nothing Then Exit Function
                
                Set oSDOProfilePart.Object = oKnuckledProfilePart
    
                dProfileHeight = oSDOProfilePart.Height
                sCrossSectionType = oSDOProfilePart.SectionType
                       
                'Only bend at Web currently.  May need to add conditions in the future for flange.
                If sCrossSectionType = "FB" Or bIsConvexKnuckle = False Then 'bend at web
                    Dim dWebThickness As Double
                    dWebThickness = oSDOProfilePart.WebThickness
                    If dWebThickness < 0.01 Then
                    m_dInnerRadius = 0.04
                    ElseIf dWebThickness < 0.02 Then
                    m_dInnerRadius = 0.08
                    ElseIf dWebThickness < 0.03 Then
                    m_dInnerRadius = 0.12
                    Else
                    m_dInnerRadius = 0.16
                    End If
                Else
                    'At flange
                      m_dInnerRadius = dProfileHeight * 4
                End If
                Set oSDOProfilePart = Nothing
                Set oKnuckledProfilePart = Nothing
            Else
                'when no profile part is available yet (happens for some ATPs), set to 0.001
                m_dInnerRadius = 0.001
            End If
            
        'End If
'    End If
    InnerRadius = m_dInnerRadius

    
End Function
' Function:
    
Private Sub GetKnuckleProfilePart(ByVal oKnuckleDisp As Object, ByVal eKnuckleType As ProfileKnuckleManufacturingMethod, oKnuckledProfilePart As Object, bIsConvexKnuckle As Boolean)
    If Not TypeOf oKnuckleDisp Is IJProfileKnuckle Then
       Exit Sub
    End If

    Dim oProfileKnuckle As IJProfileKnuckle
    Dim oProfileKnuckleHelper As IJProfileKnuckleHelper
    Dim oKnuckledProfileLeafSys As Object
    Dim oKnuckledProfileRootSys As Object
    
    Set oProfileKnuckle = oKnuckleDisp
    If oProfileKnuckle Is Nothing Then Exit Sub
        
    Set oProfileKnuckleHelper = New ProfileKnuckleHelper
    If oProfileKnuckleHelper Is Nothing Then Exit Sub
        
    oProfileKnuckleHelper.GetKnuckledProfilePartAndSystem _
                             oProfileKnuckle, oKnuckledProfilePart, oKnuckledProfileLeafSys, oKnuckledProfileRootSys
    
    bIsConvexKnuckle = False
    If eKnuckleType = pkmmBendPlusFeature Then
        On Error Resume Next
        bIsConvexKnuckle = oProfileKnuckleHelper.IsConvexKnuckle(oProfileKnuckle)
    End If
    
    Set oProfileKnuckleHelper = Nothing
    Set oKnuckledProfileLeafSys = Nothing
    Set oKnuckledProfileRootSys = Nothing
End Sub
'-------------------------------------------------------------------------------

Private Sub IJProfileKnuckleRule_ComputeKnuckleValues(ByVal pProfile As Object, ByVal pProfileKnuckle As Object, pManufacturingMethod As IMSProfileKnuckleEntity.ProfileKnuckleManufacturingMethod, pInnerRadius As Double)
    
    'Since we only have one override button for profile knuckle, the rule returns default bend radius for
    'all cases (bend or not bend) We need to set the inner radius to 0 for non bend case in profile knuckle property.
    'In the future, we could modify the rule to return 0 for split/ignore case if we have
    'two override buttons for profile knuckle.  //TR210376/210386
    pManufacturingMethod = ManufacturingMethod(pProfile, pProfileKnuckle)
    pInnerRadius = InnerRadius(pProfileKnuckle)
End Sub
' Interface IJProfileKnuckleRuleEx
' Method:   ValidateKnuckleType
' Description:
'   Evaluate if kunckle type is still valid, partticully when cross section is changed
'
Private Function IJProfileKnuckleRuleEx_ValidateKnuckleType(ByVal oKnuckle As Object) As Boolean

    IJProfileKnuckleRuleEx_ValidateKnuckleType = False
    If Not TypeOf oKnuckle Is IJProfileKnuckleMfg Then
        Exit Function
    End If

    Dim oProfileKnuckleMfg As IJProfileKnuckleMfg
    Dim eExistingMfgMethod As ProfileKnuckleManufacturingMethod

    Set oProfileKnuckleMfg = oKnuckle
    eExistingMfgMethod = oProfileKnuckleMfg.ManufacturingMethod
    Set oProfileKnuckleMfg = Nothing
    
    IJProfileKnuckleRuleEx_ValidateKnuckleType = IJProfileKnuckleRuleEx_IsKnuckleTypeAllowed(oKnuckle, eExistingMfgMethod)
End Function

'
' Check is input knuckle type appropriate for input knuckle
'    Split
'       Allowed for knuckles
'    Bend
'       Allowed for "FB", "EA", "UA", or "B"
'
'    Ignore
'       Allowed for all knuckles
'
'    Extend
'       Allowed for a knuckle realted to plate knuckle or
'       First/last profile knuckle
'
'    BendPlusFeature
'       Allowed for particular cross section types and sizes/knuckle angle
'       Cross section types:
'          "FB"
'          "EA"
'          "UA"
'          "B"
'
'       Size/knuckle angle
'          If profile size is in [200,300)
'             angle is greater than 1 degree
'          If profile size is in [300,)
'             angle is greater than 2 degrees
'
Private Function IJProfileKnuckleRuleEx_IsKnuckleTypeAllowed( _
                        ByVal oKnuckleDisp As Object, _
                        ByVal eKnuckleType As ProfileKnuckleManufacturingMethod) As Boolean
    IJProfileKnuckleRuleEx_IsKnuckleTypeAllowed = False

    
    If Not TypeOf oKnuckleDisp Is IJProfileKnuckle Then
       Exit Function
    End If

    Dim oProfileKnuckle As IJProfileKnuckle
    Dim dAngle As Double
    Dim oProfileKnuckleHelper As IJProfileKnuckleHelper
    Dim oKnuckledProfilePart As Object
    Dim oKnuckledProfileLeafSys As Object
    Dim oKnuckledProfileRootSys As Object
    
    Set oProfileKnuckle = oKnuckleDisp
    If oProfileKnuckle Is Nothing Then Exit Function

    dAngle = oProfileKnuckle.Angle
    dAngle = (dAngle * 180) / 3.14159265
        
    Set oProfileKnuckleHelper = New ProfileKnuckleHelper
    If oProfileKnuckleHelper Is Nothing Then Exit Function
        
    oProfileKnuckleHelper.GetKnuckledProfilePartAndSystem _
                             oProfileKnuckle, oKnuckledProfilePart, oKnuckledProfileLeafSys, oKnuckledProfileRootSys
    If oKnuckledProfilePart Is Nothing Then
       ' Knuckle is not related to any part, inactive
       Exit Function
    End If
    
    Dim oPlateKnuckle As Object
    Dim oProfKnuckleRefs As IJProfileKnuckleRefs
    Dim bAssociatedWithPlateKnuckle As Boolean
    Dim bIsConvexKnuckle As Boolean
    Dim bIsManualKnuckle As Boolean
    
    On Error Resume Next
    Set oProfKnuckleRefs = oKnuckleDisp
    If oProfKnuckleRefs Is Nothing Then Exit Function
    
    Set oPlateKnuckle = oProfKnuckleRefs.PlateKnuckle
    If Not oPlateKnuckle Is Nothing Then
        bAssociatedWithPlateKnuckle = True
    Else
        bAssociatedWithPlateKnuckle = False
    End If
    
    bIsManualKnuckle = oProfileKnuckle.IsPseudo
    
    Dim dProfileHeight As Double
    Dim sCrossSectionType As String
    Dim oSDOProfilePart As IJSDOProfilePart
    Dim bIsKnuckleTypeAllowed As Boolean
    
    bIsKnuckleTypeAllowed = False
    Set oSDOProfilePart = New StructDetailObjectsEx.SDOProfilePart
    If oSDOProfilePart Is Nothing Then Exit Function
    
    Set oSDOProfilePart.Object = oKnuckledProfilePart
    dProfileHeight = oSDOProfilePart.Height
    sCrossSectionType = oSDOProfilePart.SectionType
    
    Select Case eKnuckleType
        Case pkmmSplit
            bIsKnuckleTypeAllowed = True

        Case pkmmBend
            If sCrossSectionType = "FB" Then
                bIsKnuckleTypeAllowed = True
            End If
            
        Case pkmmIgnore
            bIsKnuckleTypeAllowed = True

        Case pkmmExtend
            Dim bIsFirstOrLastKnuckle As Boolean

            bIsFirstOrLastKnuckle = False
            bIsFirstOrLastKnuckle = oProfileKnuckleHelper.IsFirstOrLastKnuckleOnLeafProfile(oKnuckleDisp)
                        
            If (bAssociatedWithPlateKnuckle = True Or bIsManualKnuckle = True) And _
               bIsFirstOrLastKnuckle = True Then
                bIsKnuckleTypeAllowed = True
            Else
                bIsKnuckleTypeAllowed = False
            End If

        Case pkmmBendPlusFeature
            If oProfileKnuckle.IsPseudo = True Then
                bIsKnuckleTypeAllowed = False
            Else
                bIsConvexKnuckle = oProfileKnuckleHelper.IsConvexKnuckle(oProfileKnuckle)
                If bAssociatedWithPlateKnuckle = True Then
                    If bIsConvexKnuckle = False Then
                        ' Bend and insert at web
                        
                    Else
                        ' Bend and cut at web
                    End If
                    
                    'Not supported yet
                    bIsKnuckleTypeAllowed = False
                    
                Else
                    If bIsConvexKnuckle = False Then
                        ' Bend and insert at flange
                        Select Case sCrossSectionType
                            Case "UA", "EA"
                                If dAngle > 90# Then
                                    If dProfileHeight >= 0.2 And dProfileHeight < 0.3 Then
                                        If dAngle <= 179# Then
                                            bIsKnuckleTypeAllowed = True
                                        End If
                                    ElseIf dProfileHeight >= 0.3 Then
                                        If dAngle <= 178# Then
                                            bIsKnuckleTypeAllowed = True
                                        End If
                                    End If
                                End If
                        End Select
                    
                    ElseIf bIsConvexKnuckle = True Then
                        ' Bend and cut at flange
                        
                        ' Not supported yet
                        bIsKnuckleTypeAllowed = False
                    End If
                    
                End If
            End If

        Case pkmmSplitAndExtend
            If oProfileKnuckle.IsPseudo = True Then
                bIsKnuckleTypeAllowed = True
            Else
                bIsKnuckleTypeAllowed = False
            End If
    End Select

    IJProfileKnuckleRuleEx_IsKnuckleTypeAllowed = bIsKnuckleTypeAllowed
    
End Function
'
' This method returns root class for knuckle edge feature
' It should match the class name in bulkload file
'
Private Function IJProfileKnuckleRuleEx_GetKnuckleEdgeFeatureRootClass() As String
    IJProfileKnuckleRuleEx_GetKnuckleEdgeFeatureRootClass = "RootEdge"
End Function
'
' This method returns edge port and a point to create kunckle edge feature
' The point defaults to knuckle point. It may be changed as needed
' The edge port is on profile trim geometry
'
Private Sub IJProfileKnuckleRuleEx_GetInputsForKnuckleEdgeFeature( _
                        ByVal oKnuckleDisp As Object, _
                        ByRef oEdgePort As Object, _
                        ByRef oPointOnEdge As Object)
    On Error GoTo ErrorHandler
    
    ' Return an edge port and a point the the edge
    Dim oProfileKnuckleHelper As IJProfileKnuckleHelper
    Dim oKnuckledProfilePart As Object
    Dim oKnuckledLeafProfile As Object
    Dim oKnuckledRootProfile As Object
    
    Set oProfileKnuckleHelper = New ProfileKnuckleHelper
    If oProfileKnuckleHelper Is Nothing Then Exit Sub
        
    oProfileKnuckleHelper.GetKnuckledProfilePartAndSystem oKnuckleDisp, _
                        oKnuckledProfilePart, _
                        oKnuckledLeafProfile, _
                        oKnuckledRootProfile
    If oKnuckledProfilePart Is Nothing Then
       Set oProfileKnuckleHelper = Nothing
       Exit Sub
    End If
    
    Dim oSDOProfilePart As IJSDOProfilePart
    Dim sCrossSectionType As String
   
    Set oSDOProfilePart = New StructDetailObjectsEx.SDOProfilePart
    If oSDOProfilePart Is Nothing Then
        Set oProfileKnuckleHelper = Nothing
        Exit Sub
    End If
    
    
    Set oSDOProfilePart.Object = oKnuckledProfilePart
    sCrossSectionType = oSDOProfilePart.SectionType
   
    Dim oProfKnuckleRefs As IJProfileKnuckleRefs
    Dim bAssociatedWithPlateKnuckle As Boolean
    Dim nFace1Xid As Long
    Dim nFace2Xid As Long
    
    nFace1Xid = JXSEC_UNKNOWN
    nFace2Xid = JXSEC_UNKNOWN
    Set oProfKnuckleRefs = oKnuckleDisp
    If Not oProfKnuckleRefs.PlateKnuckle Is Nothing Then
        ' Profile knuckle is due to plate knuckle
        ' Corner feature is on Web
        nFace1Xid = JXSEC_WEB_LEFT
        nFace2Xid = JXSEC_BOTTOM
    Else
        Select Case sCrossSectionType
           Case "UA", "EA"
               nFace1Xid = JXSEC_TOP
               nFace2Xid = JXSEC_TOP_FLANGE_RIGHT
               
           Case "B"
               nFace1Xid = JXSEC_TOP
               nFace2Xid = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER
        End Select
    End If
    If nFace1Xid <> JXSEC_UNKNOWN And _
       nFace2Xid <> JXSEC_UNKNOWN Then
        Set oEdgePort = oSDOProfilePart.GetEdgePortFromFacesXids( _
                               "ProfilePartActiveEntities.ProfileTrim_AE.1", _
                               nFace1Xid, _
                               nFace2Xid, _
                               False)
    End If
   
    Set oPointOnEdge = oKnuckleDisp
    Set oProfileKnuckleHelper = Nothing
    Set oSDOProfilePart = Nothing
   
ErrorHandler:
    Exit Sub
End Sub
