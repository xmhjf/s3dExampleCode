Attribute VB_Name = "Common"
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  File        : Common.bas
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    20/Feb/2012 - svsmylav
'           CR-174918: 'FlangeCutNoFlange' method is modified to handle profile-pseudo
'           knuckle cases viz. 'SplitAndExtend' and 'Extend' Mfg. options.
'
'    20/Oct/2012 - svsmylav
'           TR-221920: Root cause of the issue is found to be that fix for TR-179462
'           needed an extra check in 'FlangeCutNoFlange' method to ensure that the plate part
'           stiffened by the ER(bounding) and the stiffener(bounded) is the same (in current
'           TR case they stiffen different plates).
'    15/Mar/2013 - skcheeka
'           TR-226309: Proper checks were provided to select approprioate flange cut items created newly for bulb cross sections.
'*********************************************************************************************
Private Const MODULE As String = PROJECT_PATH + "Common.bas"

'  ------------------------------------------------------------
' This selection rule selects a flangecut based on the two connected
' object properties and the Bounding face, EndCondition of the bounded object.
'  Selects a flangecut when the bounding profile's flange is not in the
'   the way.
'
' Inputs:
'   Bounded Object
'   bIsPlate
'   EndCondition = W,C,F,FV,S,R,or RV
'   SelectorLogicHelper
'
' Automation Function
'   BoundingFace = returns the name of the bounding object face
'                  that the base of the bounded profile's web is bounded by.
'
Public Sub FlangeCutNoFlange(oBoundedPart As Object, bIsPlate As Boolean, _
                             strEndCutType As String, pSLH As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    Dim oBoundingObject As Object
    Set oBoundingObject = pSLH.InputObject(INPUT_BOUNDING)
    If (oBoundingObject Is Nothing) Then
        strError = " pSLH.InputObject(INPUT_BOUNDING) is NOTHING"
        GoTo ErrorHandler
    End If
    
    '$$$ Free End functionality
    'Get FlangeCut smart occurence to determine if this is a FreeEnd WebCut
    'is the profile end a free end
    Dim bFreeEndCut As Boolean
    Dim oFlangeCut As StructDetailObjects.FlangeCut
    Set oFlangeCut = New StructDetailObjects.FlangeCut
    Set oFlangeCut.object = pSLH.SmartOccurrence
    bFreeEndCut = oFlangeCut.IsFreeEndCut
    
    'Determine if current flange is because of Convex knuckle with 'SplitAndExtend' Mfg. option
    Dim oProfileKnuckle As Object
    Dim iKnuckleMfgMethod As Integer
    Dim bIsConvex As Boolean
    Dim bCnvxSplitAndExtend As Boolean

    bIsConvex = False
    bCnvxSplitAndExtend = False
    
    GetProfileKnuckleType pSLH.SmartOccurrence, iKnuckleMfgMethod, oProfileKnuckle, bIsConvex
    If iKnuckleMfgMethod = pkmmSplitAndExtend And bIsConvex Then
        bCnvxSplitAndExtend = True
    Else
        'Determine if this flange cut is because of 'Extend' Mfg. option
        Dim oBoundedObject As Object
        Set oBoundedObject = pSLH.InputObject(INPUT_BOUNDED)
        If (oBoundedObject Is Nothing) Then
            strError = " pSLH.InputObject(INPUT_BOUNDED) is NOTHING"
            GoTo ErrorHandler
        End If
        
        'Currently 'oFlangeCut.IsFreeEndCut' returns False for 'Extend' Mfg. option.
        ' However, we do not want to depend on that behavior, so set it explicitly.
        'Need check for Free EndCut this cut is Not because of 'Extend'
        'Noticed SMLProfileATP:KnuckledProfilesExtended.cls had failures to extend surface body method
        ' fails (not content issue but extraneous split in profile part is the reason).
        'If IsFreeEndCutFromExtend_ConvexKnuckle(oBoundedObject) Then bFreeEndCut = False
        Set oBoundedObject = Nothing
    End If
    
    ' get Bounding Part if not a FreeEnd FlangeCut
    Dim oBoundingPart As Object
    Dim oBoundingPort As IJPort
    If TypeOf oBoundingObject Is IJPort Then
        Set oBoundingPort = oBoundingObject
        Set oBoundingPart = oBoundingPort.Connectable
        
    ElseIf Not bFreeEndCut Then
        ' error, BoundingObject MUST be a IJPort object if not FreeEnd WebCut
        strError = "BoundingObject MUST be a IJPort object"
        GoTo ErrorHandler
    End If
    
    Dim phelper As New StructDetailObjects.Helper

    ' Determine the Height of the BoundingObject
    Dim bWebCut_C2Spline As Boolean
    Dim dBoundingHeight As Double
    Dim oBoundingParentSys As Object
    Dim oBdgBottomPort As IJPort
    Dim oBdgTopPort As IJPort
    
    bWebCut_C2Spline = False
    If bFreeEndCut Then
        ' for FreeEnd FlangeCut, do not need Bounding Object Height
        ' might not have a Bounding IJPort
    ElseIf (bIsPlate) Then
        dBoundingHeight = 10    'meters
        
        'Need to check for Special case:
        'where the WebCut selected was the WebCut_C2Spline
        'Due to the WebClearance being .LE. 25mm
        '(see WebCutNoFlange called from WebCutSel::SelectorLogic)
        'WebCut_C2Spline snipes back the Web
        'Therefore the Flange Cut selected should be a Free Flange Cut
        '   Instead of re-calcualting the WebCut.BoundingPlateFlangeClearance
        '   check if the Web Cut selected is the WebCut_C2Spline
        Dim oSDO_WebCut As StructDetailObjects.WebCut
        Set oSDO_WebCut = New StructDetailObjects.WebCut
        Set oSDO_WebCut.object = oFlangeCut.WebCut
        If LCase(Trim(oSDO_WebCut.ItemName)) = LCase("WebCut_C2Spline") Then
            bWebCut_C2Spline = True
        End If
        
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
        Dim oBoundingProfile As New StructDetailObjects.ProfilePart
        Set oBoundingProfile.object = oBoundingPart
        dBoundingHeight = oBoundingProfile.Height
        Set oBoundingParentSys = oBoundingProfile.ParentSystem
        Set oBdgTopPort = oBoundingProfile.SubPort(JXSEC_TOP)
        Set oBdgBottomPort = oBoundingProfile.SubPort(JXSEC_BOTTOM)
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_BEAM Then
        Dim oBoundingBeam As New StructDetailObjects.BeamPart
        Set oBoundingBeam.object = oBoundingPart
        dBoundingHeight = oBoundingBeam.Height
        Set oBoundingParentSys = oBoundingBeam.ParentSystem
        Set oBdgTopPort = oBoundingBeam.SubPort(JXSEC_TOP)
        Set oBdgBottomPort = oBoundingBeam.SubPort(JXSEC_BOTTOM)
    Else
        ' Bounding Part type is unknown, use Default value
        dBoundingHeight = 10    'meters
    End If
    
    'Earlier approach used bounding height calulated above as is. However,
    'when bounding is a ER or else when bounded is inclined we need to compute
    ' bounding length measured in bounded-web plane. Geometrical approach is used for this.
    
    'get bounded SectionType and Height
    Dim sSectionType As String
    Dim dBoundedHeight As Double

    If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
        Dim oBoundedProfile As New StructDetailObjects.ProfilePart
        Set oBoundedProfile.object = oBoundedPart
        sSectionType = oBoundedProfile.sectionType
        dBoundedHeight = oBoundedProfile.Height
    ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        Dim oBoundedBeam As New StructDetailObjects.BeamPart
        Set oBoundedBeam.object = oBoundedPart
        sSectionType = oBoundedBeam.sectionType
        dBoundedHeight = oBoundedBeam.Height
    Else
        'Bounded Part must be a Stiffener(or Edge Reinforcment) or Beam
        Exit Sub
    End If

If (Not (oBoundedPart Is oBoundingPart)) And _
    bFreeEndCut = False Then
    'Check if bounding is ER system
    If Not oBoundingParentSys Is Nothing Then
        If TypeOf oBoundingParentSys Is IJERSystem Then
            Dim oBoundingER As New StructDetailObjects.EdgeReinforcement
            Set oBoundingER.object = oBoundingPart
            
            'Following logic works for 'OnEdge...' placement of ER
            If InStr(1, oBoundingER.ReinforcementPosition, "OnFace", vbTextCompare) < 1 Then
                'TR-179462 test case needs bIsStiffenedPlateSameto be 'True' i.e. Plate Stiffened by the Bounding-ER
                'and Plate Stiffened by the bounded stiffener to be the same part
                Dim bIsPltSystem As Boolean
                
                Dim oPltStiffenedByBounding As Object
                Set oPltStiffenedByBounding = Nothing
                Dim bIsStiffenedPlateSame As Boolean
                Dim oPltStiffenedByBounded As Object
                bIsStiffenedPlateSame = False
                
                If Not oBoundingProfile Is Nothing And _
                   Not oBoundedProfile Is Nothing Then
                    'Get stiffened plate for the bounding ER
                    Set oPltStiffenedByBounding = Nothing 'Initialize
                    oBoundingProfile.GetStiffenedPlate oPltStiffenedByBounding, bIsPltSystem
                
                    'Get stiffened plate for the bounded profile
                    Set oPltStiffenedByBounded = Nothing  'Initialize
                    oBoundedProfile.GetStiffenedPlate oPltStiffenedByBounded, bIsPltSystem
                    
                    'Check if Bounding-ER and Bounded stiffener reinforces same plate
                    If Not oPltStiffenedByBounding Is Nothing And _
                        Not oPltStiffenedByBounded Is Nothing Then
                        If oPltStiffenedByBounding Is oPltStiffenedByBounded Then
                            bIsStiffenedPlateSame = True
                        End If
                    End If
                End If
                     
                If bIsStiffenedPlateSame Then
                    Dim dPlateToTop As Double
                    Dim dPlateToBottom As Double
                    Dim oBoundingPrimary As IJDVector
                    Dim oBoundingSecondary As IJDVector
                    Dim oHelper As New StructDetailObjects.Helper
                    Dim oRootStiffBounding As Object
                    Dim oRootStiffBounded As Object
                    Dim oProfileAttributes As IJProfileAttributes
                    Set oProfileAttributes = New ProfileUtils
        
                    oBoundingER.GetPlateToTopBottom oBoundingPort, dPlateToTop, dPlateToBottom
                    Set oRootStiffBounding = oHelper.Object_ParentSystem(oBoundingPart)
                    Set oRootStiffBounded = oHelper.Object_ParentSystem(oBoundedPart)
                    
                    Dim oSBTop As IJSurface
                    Dim oSBBottom As IJSurface
                    Dim oTopCenter As IJDPosition
                    Dim oBotCenter As IJDPosition
                    Dim dDotProduct As Double
                    Dim dDistance As Double
    '               Following statement returned wrong value for primary vector: for
    '               ER this seem to need fix:
    '               oProfileAttributes.GetProfileOrientation oRootStiffBounding, oPos, _
    '                        oBoundingSecondary, oBoundingPrimary
    '               So use different approach to compute primary vector
                    'Get center of top port
                    Set oSBTop = oBdgTopPort
                    Dim inObj As Object
                    Dim numCent As Long
                    numCent = 1
                    Dim dCent() As Double
                    oSBTop.Centroid inObj, numCent, dCent(), False
                    Set oTopCenter = New DPosition
                    oTopCenter.Set dCent(0), dCent(1), dCent(2)
                    
                    'Get center of bottom port
                    Set oSBBottom = oBdgBottomPort
                    oSBBottom.Centroid inObj, numCent, dCent(), False
                    Set oBotCenter = New DPosition
                    oBotCenter.Set dCent(0), dCent(1), dCent(2)
                    
                    'Get vector joining the bottom port center to the top port
                    Set oBoundingPrimary = oTopCenter.Subtract(oBotCenter)
                    oBoundingPrimary.Length = 1 'make it a unit vector
                    
                    Dim oAC As New StructDetailObjects.AssemblyConn
                    Dim oChild As IJDesignChild
                    Set oChild = oFlangeCut.object
                    Set oAC.object = oChild.GetParent
                    Dim oEndCutPosition As IJDPosition
                    Set oEndCutPosition = New DPosition
                    Set oEndCutPosition = oAC.BoundGlobalShipLocation
        
                    Dim oBoundedPrimary As IJDVector
                    Dim oBoundedSecondary As IJDVector
                    oProfileAttributes.GetProfileOrientation oRootStiffBounded, oEndCutPosition, _
                            oBoundedSecondary, oBoundedPrimary
                    oBoundedPrimary.Length = 1

                    dDotProduct = oBoundedPrimary.Dot(oBoundingPrimary)
                    'Choose between two distances obtained earlier
                    dDistance = 0
                    If dDotProduct < 0 Then
                        dDistance = dPlateToBottom
                    Else
                        dDistance = dPlateToTop
                    End If
                    'Get projected length in bounded web plane
                    dBoundingHeight = dDistance / Abs(dDotProduct)
                End If
            End If
        End If
    End If
End If
        
    ' Check if Bounded Object has only Right Flange
    Dim bRightFlangeOnly As Boolean
    If (sSectionType = "EA") Or _
       (sSectionType = "UA") Or _
       (sSectionType = gsB) Then
       bRightFlangeOnly = True
    Else
       bRightFlangeOnly = False
    End If
        
        
    Dim eSCType As SmartClassType
    Dim eSCSubType As SmartClassSubType
    Dim sClassName As String
        
    ' the following is valid for both
    ' SDOBJECT_STIFFENER and SDOBJECT_BEAM bounded Part types
    If bFreeEndCut Then
        'For a FreeEnd FlangeCut,
        Select Case strEndCutType
                
            Case "Snip"
                If bRightFlangeOnly Then
                    If sSectionType = gsB Then
                    
                        eSCType = SMARTTYPE_FLANGECUT
                        eSCSubType = 1
                        sClassName = "FlangeCuts"
                    
                        If CheckSmartItemExists(eSCType, eSCSubType, sClassName, "FlangeCut_BF1") Then
                            pSLH.Add "FlangeCut_BF1"
                        End If
                        
                        pSLH.Add "FlangeCut_F1"
                    Else
                        '------- Removed as required for TR168651
                        'pSLH.Add "FlangeCutBackAngle"
                        '------- Added the following for TR 168651
                        pSLH.Add "FlangeCutSniped"
                        '------- Done modifcation for TR 168651
                    End If
                Else
                    pSLH.Add "FlangeCutSniped"
                End If
                
            Case "Cutback"
                If bRightFlangeOnly Then
                    If sSectionType = gsB Then
                        pSLH.Add "FlangeCutBackBulb"
                    Else
                        pSLH.Add "FlangeCutBackAngle"
                    End If
                Else
                    pSLH.Add "FlangeCutBackBUT"
                End If
                
             Case "Clip"
                 If bRightFlangeOnly Then
                     If sSectionType = "UA" Or sSectionType = "EA" Then
                        pSLH.Add "FlangeCutSnipWeld"
                     End If
                  Else
                      pSLH.Add "FlangeCutSniped"
                  End If
                
            Case "Welded", "Bracketed"
                If bRightFlangeOnly Then
                    If sSectionType = gsB Then
                        pSLH.Add "FlangeCutBackBulb"
                    Else
                        pSLH.Add "FlangeCutBackAngle"
                    End If
                Else
                    pSLH.Add "FlangeCutBackBUT"
                End If
        End Select
        
        Exit Sub
    End If
               
    
'    If phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
'        'Profile bounded by profile case.
'        'If the landing curves are not intersecting, we would go with a simple flange cut.
'        If Not AreLandingCurvesIntersecting(oBoundedPart, oBoundingPart) Then
'            If bRightFlangeOnly Then
'                pSLH.Add "FlangeCut_F1"
'            Else
'                pSLH.Add "FlangeCut_F4"
'            End If
'
'            Exit Sub
'        End If
'    End If

    ' the following is valid for both
    ' SDOBJECT_STIFFENER and SDOBJECT_BEAM bounded Part types
    Select Case strEndCutType
        Case "Snip"
            If bRightFlangeOnly Then
                If sSectionType = gsB Then
                    eSCType = SMARTTYPE_FLANGECUT
                    eSCSubType = 1
                    sClassName = "FlangeCuts"
                    
                    If CheckSmartItemExists(eSCType, eSCSubType, sClassName, "FlangeCut_BF1") Then
                        pSLH.Add "FlangeCut_BF1"
                    End If
                    
                    pSLH.Add "FlangeCut_F1"
                Else
                    pSLH.Add "FlangeCutBackAngle"
                End If
            Else
                pSLH.Add "FlangeCutSniped"
            End If
        
        Case "Cutback"
                If bRightFlangeOnly Then
                    If sSectionType = gsB Then
                        pSLH.Add "FlangeCutBackBulb"
                    Else
                        pSLH.Add "FlangeCutBackAngle"
                    End If
                Else
                    pSLH.Add "FlangeCutBackBUT"
                End If
        
        Case "Clip"
                 If bRightFlangeOnly Then
                     If sSectionType = "UA" Or sSectionType = "EA" Then
                        pSLH.Add "FlangeCutSnipWeld"
                     End If
                  Else
                      pSLH.Add "FlangeCutSniped"
                  End If
                  
        Case "Welded", "Bracketed"
            Dim bBoundaryLCAlongBeamV As Boolean
            
            bBoundaryLCAlongBeamV = False
            bBoundaryLCAlongBeamV = IsBoundaryLandindCurveAlongBeamVDirection( _
                                         oBoundedPart, _
                                         oBoundingPart, _
                                         oFlangeCut.WebCut)
            If (bBoundaryLCAlongBeamV = True) Then
               ' This this the case when a beam is bounded by ER Web
               pSLH.Add "FlangeCutStraight"
               
            ElseIf GreaterThanOrEqualTo((dBoundingHeight - 0.015), dBoundedHeight) Then
                ' Check for Special Case: WebCut_C2Spline
                ' Web is Sniped Back, select a Free Flange Cut
                If bWebCut_C2Spline Then
                    If bRightFlangeOnly Then
                        pSLH.Add "FlangeCut_F1"
                    Else
                        pSLH.Add "FlangeCut_F4"
                    End If
                
                'weld needed
                ElseIf bRightFlangeOnly Then
                    If sSectionType = gsB Then
                        pSLH.Add "FlangeCutStraightBulb"
                        pSLH.Add "FlangeCutSnipWeld"
                    Else
                        pSLH.Add "FlangeCutStraightAngle"
                        pSLH.Add "FlangeCutSnipWeld"

                    End If
                Else
                    pSLH.Add "FlangeCutStraight"
                End If
            ElseIf Equal(dBoundingHeight, dBoundedHeight) Then '=
                If bRightFlangeOnly Then
                    If bCnvxSplitAndExtend Then
                        'Need welded flange cut items
                        If sSectionType = gsB Then
                            pSLH.Add "FlangeCutStraightBulb"
                            pSLH.Add "FlangeCutSnipWeld"
                        Else
                            pSLH.Add "FlangeCutStraightAngle"
                            pSLH.Add "FlangeCutSnipWeld"
                        End If
                    Else
                        If sSectionType = gsB Then
                            pSLH.Add "FlangeCutBackBulb"
                        Else
                            pSLH.Add "FlangeCutBackAngle"
                        End If
                    End If
                Else
                    pSLH.Add "FlangeCutBackBUT"
                End If
            Else
                If sSectionType = gsB Then
                    eSCType = SMARTTYPE_FLANGECUT
                    eSCSubType = 1
                    sClassName = "FlangeCuts"
                    
                    If CheckSmartItemExists(eSCType, eSCSubType, sClassName, "FlangeCut_BF1") Then
                        pSLH.Add "FlangeCut_BF1"
                    End If
                    pSLH.Add "FlangeCut_F1"
                    
                ElseIf sSectionType = "BUT" Then
                    pSLH.Add "FlangeCutBackBUT"
                Else
                        pSLH.Add "FlangeCutSnipedAngle"
                        pSLH.Add "FlangeCutBackAngle"
                End If
            End If


    End Select
    
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "FlangeCutNoFlange", strError).Number
End Sub
    
' Inputs:
'   Bounded Object
'   Bounding Object
'   EndCondition
'   Fixity
'   Alpha = Angle between the bounding and bounded in the plane of the bounded web
'
'   Based on the EndCondition and projected heights of the profiles select webcut,
'   selects a webcut when the bounding profile's flange is potentially in the
'   the way.

Public Sub FlangeCutFlanged(oBoundedPart As Object, _
                            strEndCutType As String, pSLH As IJDSelectorLogic)

    Dim strError As String
    
    ' Verify the BoundingObject exists
    Dim oBoundingObject As Object
    Set oBoundingObject = pSLH.InputObject(INPUT_BOUNDING)
    If (oBoundingObject Is Nothing) Then
        strError = " pSLH.InputObject(INPUT_BOUNDING) is NOTHING"
        GoTo ErrorHandler
    End If
    
    ' Verify the BoundingObject is an IJPort (not a FreeEndCut)
    If Not TypeOf oBoundingObject Is IJPort Then
        Exit Sub
    End If
    
    Dim oBoundingPart  As Object
    Dim oBoundingPort As IJPort
    Set oBoundingPort = oBoundingObject
    Set oBoundingPart = oBoundingPort.Connectable
        
    Dim dBoundingHeight As Double
    Dim dFlangeThickness As Double
    Dim phelper As New StructDetailObjects.Helper
    Dim dist As Double, alpha As Double
        
    If phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
        Dim oBoundingProfile As New StructDetailObjects.ProfilePart
        Set oBoundingProfile.object = oBoundingPart
        dBoundingHeight = oBoundingProfile.Height
        dFlangeThickness = oBoundingProfile.flangeThickness
        If oBoundingProfile.sectionType = gsB Then
            dist = 0.01
        Else
            dist = 0.035
        End If
    
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_BEAM Then
        Dim oBoundingBeam As New StructDetailObjects.BeamPart
        Set oBoundingBeam.object = oBoundingPart
        dBoundingHeight = oBoundingBeam.Height
        dFlangeThickness = oBoundingBeam.flangeThickness
        If oBoundingBeam.sectionType = gsB Then
            dist = 0.01
        Else
            dist = 0.035
        End If
       
    Else
        dist = 0.035
   End If
    
    ' Verify the BoundedObject exists
    Dim oBoundedObject As Object
    Set oBoundedObject = pSLH.InputObject(INPUT_BOUNDED)
    If (oBoundedObject Is Nothing) Then
        strError = " pSLH.InputObject(INPUT_BOUNDED) is NOTHING"
        GoTo ErrorHandler
    End If
   
    ' Verify the BoundedObject is an IJPort
    If Not TypeOf oBoundingObject Is IJPort Then
        Exit Sub
    End If
    
    Dim strSectionType As String
    Dim oBoundedProfile As StructDetailObjects.ProfilePart
    Dim oBoundedBeam As StructDetailObjects.BeamPart
    
    If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
        Set oBoundedProfile = New StructDetailObjects.ProfilePart
        Set oBoundedProfile.object = oBoundedPart
        strSectionType = oBoundedProfile.sectionType
        
    ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        Set oBoundedBeam = New StructDetailObjects.BeamPart
        Set oBoundedBeam.object = oBoundedPart
        strSectionType = oBoundedBeam.sectionType
    End If
    
'    If phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
'
'        'Profile bounded by profile case.
'        'If the landing curves are not intersecting, we would go with a simple flange cut.
'        If Not AreLandingCurvesIntersecting(oBoundedPart, oBoundingPart) Then
'            If strSectionType = "EA" Or _
'               strSectionType = "UA" Or _
'               strSectionType = gsB Then
'
'                pSLH.Add "FlangeCut_F1"
'            Else
'                pSLH.Add "FlangeCut_F4"
'            End If
'
'            Exit Sub
'        End If
'    End If
    
    Dim oBoundedPort As IJPort
    Set oBoundedPort = oBoundedObject
    
    'Get the name of the port, used to determine if part is connected
    'to the web or the flange
    Dim dPortValue As Long
    Dim oPartInfo As New StructDetailObjects.Helper
    dPortValue = oPartInfo.GetBoundingProfileFace(oBoundedPort, oBoundingPort)
    
    'Hardcode the angle to 90 deg
    alpha = 90
    
    'get bounded height
    Dim dBoundedHeight As Double
    Dim dBoundedWebLength As Double
    
    Dim eSCType As SmartClassType
    Dim eSCSubType As SmartClassSubType
    Dim sClassName As String
    
    If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
        dBoundedHeight = oBoundedProfile.Height
        dBoundedWebLength = oBoundedProfile.WebLength
        If oBoundedProfile.sectionType = "EA" Or oBoundedProfile.sectionType = "UA" Then
            dBoundedWebLength = dBoundedWebLength - oBoundedProfile.flangeThickness
        End If
        
            Select Case strEndCutType
                Case "Snip"
                    If strSectionType = "EA" Or _
                       strSectionType = "UA" Then
                        pSLH.Add "FlangeCutBackAngle"
                        pSLH.Add "FlangeCutSnipedAngle"
                    ElseIf strSectionType = gsB Then
                    
                        eSCType = SMARTTYPE_FLANGECUT
                        eSCSubType = 1
                        sClassName = "FlangeCuts"
                    
                        If CheckSmartItemExists(eSCType, eSCSubType, sClassName, "FlangeCut_BF1") Then
                            pSLH.Add "FlangeCut_BF1"
                        End If
                        
                        pSLH.Add "FlangeCut_F1"
                    Else
                        pSLH.Add "FlangeCutSniped"
                    End If
                    
                Case "Cutback"
                    If strSectionType = gsB Then
                        pSLH.Add "FlangeCutBackBulb"
                    ElseIf strSectionType = "UA" Or strSectionType = "EA" Then
                        pSLH.Add "FlangeCutBackAngle"
                    Else
                        pSLH.Add "FlangeCutBackBUT"
                    End If
                    
                Case "Clip"
                    If strSectionType = "UA" Or strSectionType = "EA" Then
                        pSLH.Add "FlangeCutSnipWeld"
                    Else
                        pSLH.Add "FlangeCutSniped"
                    
                    End If
                  
                        
                Case "Welded", "Bracketed"
                    If dPortValue = JXSEC_TOP Or _
                        dPortValue = JXSEC_BOTTOM Or _
                        dPortValue = JXSEC_FLANGE_TOP Or _
                        dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                        dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                        dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                        dPortValue = JXSEC_WEB_LEFT_TOP Or _
                        dPortValue = JXSEC_FLANGE_BOTTOM Or _
                        dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                        dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                        dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                        dPortValue = JXSEC_WEB_LEFT_BOTTOM Then
                        
                        If strSectionType = "EA" Or _
                           strSectionType = "UA" Then
                            pSLH.Add "FlangeCutStraightAngle"
                            pSLH.Add "FlangeCutSnipWeld"
                        ElseIf strSectionType = gsB Then
                            pSLH.Add "FlangeCutStraightBulb"
                            pSLH.Add "FlangeCutSnipWeld"
                        Else
                            pSLH.Add "FlangeCutStraight"
                        End If
                    Else
                        If (dBoundingHeight - dFlangeThickness - dist) >= dBoundedHeight Then
                            If strSectionType = "EA" Or _
                               strSectionType = "UA" Then
                                pSLH.Add "FlangeCutStraightAngle"
                                pSLH.Add "FlangeCutSnipWeld"
                            ElseIf strSectionType = gsB Then
                                pSLH.Add "FlangeCutStraightBulb"
                                pSLH.Add "FlangeCutSnipWeld"
                            Else
                                pSLH.Add "FlangeCutStraight"
                            End If
                        Else 'the bounding profile is not taller
                            If dBoundedWebLength < dBoundingHeight Then
                                If (oBoundedProfile.sectionType = "EA") Or _
                                    (oBoundedProfile.sectionType = "UA") Then
                                    pSLH.Add "FlangeCut_W1"
                                    pSLH.Add "FlangeCutSnipWeld"
                                ElseIf (oBoundedProfile.sectionType = gsB) Then
                                    'no need for a flange cut on the bulb - the web cut gets everything
    '                                pSLH.Add "FlangeCutStraightBCope"
                                Else
                                    pSLH.Add "FlangeCut_W4"
                                End If
                            Else
                            
                                If (oBoundedProfile.sectionType = "EA") Or _
                                    (oBoundedProfile.sectionType = "UA") Then
                                pSLH.Add "FlangeCutStraightACope"
                                pSLH.Add "FlangeCutSnipWeld"
                            ElseIf (oBoundedProfile.sectionType = gsB) Then
                                'no need for a flange cut on the bulb - the web cut gets everything
'                                pSLH.Add "FlangeCutStraightBCope"
                            Else
                                    pSLH.Add "FlangeCutStraightCope"
                            End If
                        End If
                    End If
                    End If
        
            End Select
    
    ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        dBoundedHeight = oBoundedBeam.Height
        dBoundedWebLength = oBoundedBeam.WebLength
        If oBoundedBeam.sectionType = "EA" Or oBoundedBeam.sectionType = "UA" Then
            dBoundedWebLength = dBoundedWebLength - oBoundedBeam.flangeThickness
        End If
            Select Case strEndCutType
                Case "Snip"
                    If strSectionType = "EA" Or _
                       strSectionType = "UA" Then
                        pSLH.Add "FlangeCutBackAngle"
                        pSLH.Add "FlangeCutSnipedAngle"
                    ElseIf strSectionType = gsB Then
                    
                        eSCType = SMARTTYPE_FLANGECUT
                        eSCSubType = 1
                        sClassName = "FlangeCuts"
                    
                        If CheckSmartItemExists(eSCType, eSCSubType, sClassName, "FlangeCut_BF1") Then
                            pSLH.Add "FlangeCut_BF1"
                        End If
                        
                        pSLH.Add "FlangeCut_F1"
                    Else
                        pSLH.Add "FlangeCutSniped"
                    End If

                    
                Case "Cutback"
                    If strSectionType = gsB Then
                        pSLH.Add "FlangeCutBackBulb"
                    ElseIf strSectionType = "UA" Or strSectionType = "EA" Then
                        pSLH.Add "FlangeCutBackAngle"
                    Else
                        pSLH.Add "FlangeCutBackBUT"
                    End If
                    
                Case "Clip"
                    If strSectionType = "UA" Or strSectionType = "EA" Then
                        pSLH.Add "FlangeCutSnipWeld"
                    Else
                        pSLH.Add "FlangeCutSniped"
                    End If
                    
                Case "Welded", "Bracketed"
                    If dPortValue = JXSEC_TOP Or _
                        dPortValue = JXSEC_BOTTOM Or _
                        dPortValue = JXSEC_FLANGE_TOP Or _
                        dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                        dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                        dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                        dPortValue = JXSEC_WEB_LEFT_TOP Or _
                        dPortValue = JXSEC_FLANGE_BOTTOM Or _
                        dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                        dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                        dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                        dPortValue = JXSEC_WEB_LEFT_BOTTOM Then
                        
                        If strSectionType = "EA" Or _
                           strSectionType = "UA" Then
                            pSLH.Add "FlangeCutStraightAngle"
                        ElseIf strSectionType = gsB Then
                            pSLH.Add "FlangeCutStraightBulb"
                        Else
                            pSLH.Add "FlangeCutStraight"
                        End If
                    Else
                        If (dBoundingHeight - dFlangeThickness - dist) >= dBoundedHeight Then
                            If strSectionType = "EA" Or _
                               strSectionType = "UA" Then
                                pSLH.Add "FlangeCutStraightAngle"
                            ElseIf strSectionType = gsB Then
                                pSLH.Add "FlangeCutStraightBulb"
                            Else
                                pSLH.Add "FlangeCutStraight"
                            End If
                        Else 'the bounding profile is not taller
                            If dBoundedWebLength < dBoundingHeight Then
                                If (oBoundedProfile.sectionType = "EA") Or _
                                    (oBoundedProfile.sectionType = "UA") Then
                                    pSLH.Add "FlangeCut_W1"
                                    pSLH.Add "FlangeCutSnipWeld"
                                ElseIf (oBoundedProfile.sectionType = gsB) Then
                                    'no need for a flange cut on the bulb - the web cut gets everything
    '                                pSLH.Add "FlangeCutStraightBCope"
                                Else
                                        pSLH.Add "FlangeCut_W4"
                                End If
                            Else
                                If (oBoundedProfile.sectionType = "EA") Or _
                                    (oBoundedProfile.sectionType = "UA") Then
                                pSLH.Add "FlangeCutStraightACope"
                                pSLH.Add "FlangeCutSnipWeld"
                            ElseIf (oBoundedProfile.sectionType = gsB) Then
                                'no need for a flange cut on the bulb - the web cut gets everything
'                                pSLH.Add "FlangeCutStraightBCope"
                            Else
                                    pSLH.Add "FlangeCutStraightCope"
                            End If
                        End If
                    End If
                    End If
        
            End Select
    End If
  
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "FlangeCutFlanged", strError).Number
End Sub
Public Sub FlangeCutNoFlangeEndToEnd(oBoundedPart As Object, _
                                     bIsPlate As Boolean, _
                                     strEndCutType As String, _
                                     strWeldPartNumber As String, _
                                     pSLH As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    Dim phelper As New StructDetailObjects.Helper
    
    Dim strError As String
    
    ' Verify the BoundingObject exists
    Dim oBoundingObject As Object
    Set oBoundingObject = pSLH.InputObject(INPUT_BOUNDING)
    If (oBoundingObject Is Nothing) Then
        strError = " pSLH.InputObject(INPUT_BOUNDING) is NOTHING"
        GoTo ErrorHandler
    End If
    
zMsgBox "::FlangeCutNoFlangeEndToEnd" & vbCrLf & _
        "   bIsPlate.........: " & bIsPlate & vbCrLf & _
        "   strEndCutType....: " & strEndCutType & vbCrLf & _
        "   strWeldPartNumber: " & strWeldPartNumber

    Dim oFlangeCutWrapper As New StructDetailObjects.FlangeCut
    Dim oWebCut As Object
    Dim bIsAlligatorCase As Boolean
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim sItemName As String
    
    bIsAlligatorCase = False
    Set oFlangeCutWrapper.object = pSLH.SmartOccurrence
    Set oSmartOccurrence = oFlangeCutWrapper.WebCut
    Set oSmartItem = oSmartOccurrence.SmartItemObject
    sItemName = oSmartItem.Name
    If LCase(sItemName) = LCase("WebCut_C1AlligatorPCEnd") Or _
       LCase(sItemName) = LCase("WebCut_C1AlligatorEnd") Then
       bIsAlligatorCase = True
    End If
    
    If LCase(sItemName) = LCase("WebCut_OffsetInside") Or _
       LCase(sItemName) = LCase("WebCut_OffsetOutside") Then
       ' Straight Offset case
       If strEndCutType = "Welded" Then
          If strWeldPartNumber = "First" Then
             pSLH.Add "FlangeCut_OffsetInside" ' This cut always has PhysConn
          Else
             pSLH.Add "FlangeCut_OffsetOutside"
          End If
       End If
       
       Exit Sub
    End If
    
    'get bounded height
    Dim dBoundedHeight As Double
    Dim sSectionType As String
    Dim eSCType As SmartClassType
    Dim eSCSubType As SmartClassSubType
    Dim sClassName As String
    
    If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
        Dim oBoundedProfile As New StructDetailObjects.ProfilePart
        Set oBoundedProfile.object = oBoundedPart
        dBoundedHeight = oBoundedProfile.Height
        sSectionType = oBoundedProfile.sectionType
        
        Select Case strEndCutType
            Case "Welded"
                If strWeldPartNumber = "First" Then
                    If sSectionType = "EA" Or _
                       sSectionType = "UA" Or _
                       sSectionType = gsB Then
                            If sSectionType = gsB Then
                                eSCType = SMARTTYPE_FLANGECUT
                                eSCSubType = 2
                                sClassName = "FlangeCutsEndToEnd"
                            
                                If CheckSmartItemExists(eSCType, eSCSubType, sClassName, "FlangeCut_BulbPCEnd") Then
                                    pSLH.Add "FlangeCut_BulbPCEnd"
                                End If
                            End If
                       pSLH.Add "FlangeCut_W1PCEnd"
zMsgBox "::FlangeCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add FlangeCut_W1PCEnd"
        
                    Else
                        If bIsAlligatorCase = False Then
                            pSLH.Add "FlangeCut_W4PCEnd"
zMsgBox "::FlangeCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add FlangeCut_W4PCEnd"
                        Else
                            pSLH.Add "FlangeCut_W4AlligatorPCEnd"
                        End If
                    End If
                Else
                    If sSectionType = "EA" Or _
                      sSectionType = "UA" Or _
                       sSectionType = gsB Then
                       
                        If sSectionType = gsB Then
                            eSCType = SMARTTYPE_FLANGECUT
                            eSCSubType = 2
                            sClassName = "FlangeCutsEndToEnd"
                        
                            If CheckSmartItemExists(eSCType, eSCSubType, sClassName, "FlangeCut_BulbEnd") Then
                                pSLH.Add "FlangeCut_BulbEnd"
                            End If
                        End If
                        
                        pSLH.Add "FlangeCut_W1End"
zMsgBox "::FlangeCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add FlangeCut_W1End"
                    Else
                        If bIsAlligatorCase = False Then
                            pSLH.Add "FlangeCut_W4End"
zMsgBox "::FlangeCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add FlangeCut_W4End"
                        Else
                            pSLH.Add "FlangeCut_W4AlligatorEnd"
                        End If
                    End If
                End If
            
            Case Else
zMsgBox "::FlangeCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add (Nothing)"
        End Select
        
    ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        Dim oBoundedBeam As New StructDetailObjects.BeamPart
        Set oBoundedBeam.object = oBoundedPart
        dBoundedHeight = oBoundedBeam.Height
        sSectionType = oBoundedBeam.sectionType
        
        Select Case strEndCutType
            Case "Welded"
                If strWeldPartNumber = "First" Then
                    If sSectionType = "EA" Or _
                       sSectionType = "UA" Or _
                        sSectionType = gsB Then
                            If sSectionType = gsB Then
                                eSCType = SMARTTYPE_FLANGECUT
                                eSCSubType = 2
                                sClassName = "FlangeCutsEndToEnd"
                            
                                If CheckSmartItemExists(eSCType, eSCSubType, sClassName, "FlangeCut_BulbPCEnd") Then
                                    pSLH.Add "FlangeCut_BulbPCEnd"
                                End If
                            End If
                        pSLH.Add "FlangeCut_W1PCEnd"
                    Else
                        If bIsAlligatorCase = False Then
                            pSLH.Add "FlangeCut_W4PCEnd"
                        Else
                            pSLH.Add "FlangeCut_W4AlligatorPCEnd"
                        End If
                    End If
                Else
                    If sSectionType = "EA" Or _
                       sSectionType = "UA" Or _
                        sSectionType = gsB Then
                        If sSectionType = gsB Then
                            eSCType = SMARTTYPE_FLANGECUT
                            eSCSubType = 2
                            sClassName = "FlangeCutsEndToEnd"
                        
                            If CheckSmartItemExists(eSCType, eSCSubType, sClassName, "FlangeCut_BulbEnd") Then
                                pSLH.Add "FlangeCut_BulbEnd"
                            End If
                        End If
                        pSLH.Add "FlangeCut_W1End"
                    Else
                        If bIsAlligatorCase = False Then
                            pSLH.Add "FlangeCut_W4End"
                        Else
                            pSLH.Add "FlangeCut_W4AlligatorEnd"
                        End If
                    End If
                End If
        End Select
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "FlangeCutNoFlangeEndToEnd", strError).Number
End Sub

'  ------------------------------------------------------------
' This selection rule selects a webcut based on the two connected
' object properties and the Bounding face, EndCondition of the bounded object.
'  Selects a webcut when the bounding profile's flange is not in the
'   the way.
'
' Inputs:
'   Bounded Object
'   bIsPlate
'   EndCondition = W,C,F,FV,S,R,or RV
'   SelectorLogicHelper
'
' Automation Function
'   BoundingFace = returns the name of the bounding object face
'                  that the base of the bounded profile's web is bounded by.
'
Public Sub WebCutNoFlange(oBoundedPart As Object, bIsPlate As Boolean, _
                          strEndCutType As String, pSLH As IJDSelectorLogic)
'
'   Base on the EndCondition and projected heights of the profiles select webcut.
'
    On Error GoTo ErrorHandler
    Dim sTypeObject As String
    Dim sObjectType As String
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oEdgeReinforcement As StructDetailObjects.EdgeReinforcement
    Dim dPortValue As Long
    Dim strError As String
    
    Dim oBoundingObject As Object
    Set oBoundingObject = pSLH.InputObject(INPUT_BOUNDING)
    If (oBoundingObject Is Nothing) Then
        strError = " pSLH.InputObject(INPUT_BOUNDING) is NOTHING"
        GoTo ErrorHandler
    End If
    
    Dim oBoundedObject As Object
    Set oBoundedObject = pSLH.InputObject(INPUT_BOUNDED)
    If (oBoundedObject Is Nothing) Then
        strError = " pSLH.InputObject(INPUT_BOUNDED) is NOTHING"
        GoTo ErrorHandler
    End If
   
    ' get Bounded Port
    Dim oBoundedPort As IJPort
    If TypeOf oBoundedObject Is IJPort Then
        Set oBoundedPort = oBoundedObject
    Else
        ' error, BoundedObject MUST be a IJPort object
        strError = "BoundedObject MUST be a IJPort object"
        GoTo ErrorHandler
    End If

    '$$$ Free End functionality
    ' no check is needed here because all FreeEnd WebCuts are
    ' handled in the WebCutSel logic
    
    ' get Bounding Part if not a FreeEnd FlangeCut
    Dim oBoundingPart As Object
    Dim oBoundingPort As IJPort
    If TypeOf oBoundingObject Is IJPort Then
        Set oBoundingPort = oBoundingObject
        Set oBoundingPart = oBoundingPort.Connectable
    Else
        ' error, BoundingObject MUST be a IJPort object if not FreeEnd WebCut
        strError = "BoundingObject MUST be a IJPort object"
        GoTo ErrorHandler
    End If
    
    Dim dBoundingHeight As Double
    dBoundingHeight = 0#
    
    'get the answer for theta
    Dim sTheta As String
    sTheta = pSLH.SelectorAnswer(PROJECT_NAME + ".WebCutSel", "Theta")
    Dim dTheta As Double
    dTheta = Val(sTheta)

    Set oSDO_Helper = New StructDetailObjects.Helper
    If (bIsPlate) Then
        Dim oWebCut As New StructDetailObjects.WebCut
        Set oWebCut.object = pSLH.SmartOccurrence
                
        Dim dFlangeClearance As Double
        dFlangeClearance = oWebCut.BoundingPlateFlangeClearance
        
        Set oWebCut = Nothing
        
        If dFlangeClearance > 0.025 Then
            dBoundingHeight = 10  'meters
        Else
            pSLH.Add "WebCut_C2Spline"
            Exit Sub
        End If
    Else
'        If oSDO_Helper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
'            'Profile bounded by profile case.
'            'If the landing curves are not intersecting, we would go with a simple web cut.
'
'            If Not AreLandingCurvesIntersecting(oBoundedPart, oBoundingPart) Then
'                pSLH.Add "WebAndFlangeWelded"
'                Exit Sub
'            End If
'        End If
        
        ' Retrieve the Bounding Profile's Height
        dBoundingHeight = GetBoundingStiffenerHeight(oBoundedObject, _
                                                     oBoundingObject)
    End If
               
    ' Default the Height to 10.0 meters
    If Abs(dBoundingHeight) < 0.001 Then
        dBoundingHeight = 10#
    End If
    
    'get bounded height
    Dim dBoundedHeight As Double
    If oSDO_Helper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
        Dim oBoundedProfile As New StructDetailObjects.ProfilePart
        Set oBoundedProfile.object = oBoundedPart
        dBoundedHeight = oBoundedProfile.Height
        
        '**** Code added to handle connected edge reinforcement,
        '**** on edge with F,FV or S endcuttype
        ' Will Need to add additional logic to handle all cases
        ' of edge reinforcment bounded by profile.
        '
        'Is the stiffener an edge reinforcement
        oSDO_Helper.GetObjectTypeData oBoundedPart, sTypeObject, sObjectType
        If sTypeObject = "EdgeReinforcement" Then
            'Is the Edge Reinforcement mounted On Face
            'then treat it as a normal stiffener
            Set oEdgeReinforcement = New StructDetailObjects.EdgeReinforcement
            Set oEdgeReinforcement.object = oBoundedPart
            
            'check if Edge Reinforcement is on edge
            If Not oEdgeReinforcement.ReinforcementPosition = "OnFace" Then
                'if the edge reinforcement a Flat bar.
                'If not a flat bar simply cut neat to bounding object regardless of end cut type.
                If oBoundedProfile.sectionType = "FB" Then
                    Select Case strEndCutType
                        Case "Snip"
                            'Edge Reinforcement
                            'Select web cut that looks like a sniped Tee flange cut
                            pSLH.Add "FreeEndWebCut_F1_ER"
                        
                        Case Else
                            pSLH.Add "WebAndFlangeWelded"
                        End Select
                Else
                    pSLH.Add "WebAndFlangeWelded"
                End If
            End If
            
            ' if the edge reiforcement is on face, treat it like a stiffener
        End If
        Select Case strEndCutType
            Case "Snip"
                'for bounded profiles, do not check the theta
                If ((oBoundedProfile.sectionType = gsB) Or _
                   (oBoundedProfile.sectionType = "FB")) Then
                    pSLH.Add "Snip"
                Else
                    pSLH.Add "SnipWithFlange"
                End If
            
                
            Case "Cutback"
                pSLH.Add "StraightSnip"
                
            Case "Snip"
                pSLH.Add "Clip"
            
            Case "Welded", "Bracketed"
                If (dBoundingHeight - 0.015) >= dBoundedHeight Then
                    If (bIsPlate) Then
                        'This end cut symbol uses a spline for the bounding curve.
                        'To minimize the impact on existing catalogs, the old WebCut_C1
                        'symbol is still used for cases where the bounding object is not
                        'a plate.
                        pSLH.Add "WebAndFlangeWelded"
                    Else
                        pSLH.Add "WebAndFlangeWelded"
                    End If
                    
                Else
                    'Get the name of the port, used to determine side for the symbol
                    dPortValue = oSDO_Helper.GetBoundingProfileFace(oBoundedPort, oBoundingPort)
                    oSDO_Helper.GetObjectTypeData oBoundingPart, sTypeObject, sObjectType
                    If dPortValue = JXSEC_WEB_LEFT Or _
                        dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                        If sTypeObject = "EdgeReinforcement" Then
                            pSLH.Add "SnipTopCornerLeft_ER"
                        Else
                            pSLH.Add "SnipTopCornerLeft"
                            pSLH.Add "CopeTopCornerSnipe"
                        End If
                            
                        
                    ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                        If sTypeObject = "EdgeReinforcement" Then
                            pSLH.Add "SnipTopCornerRight_ER"
                        Else
                            pSLH.Add "SnipTopCornerRight"
                            pSLH.Add "CopeTopCornerSnipe"
                        End If
                        
                        
                    ElseIf dPortValue = JXSEC_TOP Or _
                        dPortValue = JXSEC_BOTTOM Or _
                        dPortValue = JXSEC_FLANGE_TOP Or _
                        dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                        dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                        dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                        dPortValue = JXSEC_WEB_LEFT_TOP Or _
                        dPortValue = JXSEC_FLANGE_BOTTOM Or _
                        dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                        dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                        dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                        dPortValue = JXSEC_WEB_LEFT_BOTTOM Then

                        pSLH.Add "Clip"
                        
                    End If
                End If
            
        End Select
        
    ElseIf oSDO_Helper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        Dim oBoundedBeam As New StructDetailObjects.BeamPart
        Set oBoundedBeam.object = oBoundedPart
        dBoundedHeight = oBoundedBeam.Height
        
        Select Case strEndCutType
            Case "Snip"
                'for bounded profiles, do not check the theta
                If ((oBoundedProfile.sectionType = gsB) Or _
                   (oBoundedProfile.sectionType = "FB")) Then
                    pSLH.Add "Snip"
                Else
                    pSLH.Add "SnipWithFlange"
                End If
             
            Case "Cutback"
                pSLH.Add "StraightSnip"
                
            Case "Clip"
                pSLH.Add "Clip"
                
            Case "Welded", "Bracketed"
                Dim bIsBoundaryCurveAlongBeamVDir As Boolean
                
                bIsBoundaryCurveAlongBeamVDir = False
                bIsBoundaryCurveAlongBeamVDir = IsBoundaryLandindCurveAlongBeamVDirection( _
                                                oBoundedPart, oBoundingPart, pSLH.SmartOccurrence)
                If bIsBoundaryCurveAlongBeamVDir = True Then
                    pSLH.Add "WebAndFlangeWelded"
                    
                ElseIf (dBoundingHeight - 0.015) >= dBoundedHeight Then
                    If (bIsPlate) Then
                        'This end cut symbol uses a spline for the bounding curve.
                        'To minimize the impact on existing catalogs, the old WebCut_C1
                        'symbol is still used for cases where the bounding object is not
                        'a plate.
                        pSLH.Add "WebAndFlangeWelded"
                    Else
                        pSLH.Add "WebAndFlangeWelded"
                    End If
                    
                Else
                    'Get the name of the port, used to determine side for the symbol
                    dPortValue = oSDO_Helper.GetBoundingProfileFace(oBoundedPort, oBoundingPort)
                                    
                    If dPortValue = JXSEC_WEB_LEFT Or _
                        dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                        
                        pSLH.Add "CopeTopCornerSnipe"
                        pSLH.Add "SnipTopCornerLeft"
                        
                    ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                        
                        pSLH.Add "CopeTopCornerSnipe"
                        pSLH.Add "SnipTopCornerRight"
                        
                    ElseIf dPortValue = JXSEC_TOP Or _
                        dPortValue = JXSEC_BOTTOM Or _
                        dPortValue = JXSEC_FLANGE_TOP Or _
                        dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                        dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                        dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                        dPortValue = JXSEC_WEB_LEFT_TOP Or _
                        dPortValue = JXSEC_FLANGE_BOTTOM Or _
                        dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                        dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                        dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                        dPortValue = JXSEC_WEB_LEFT_BOTTOM Then

                        pSLH.Add "Clip"
                    End If
                End If
            
        End Select
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "WebCutNoFlange", strError).Number
End Sub

    
' Inputs:
'   Bounded Object
'   Bounding Object
'   EndCondition
'   Fixity
'   Alpha = Angle between the bounding and bounded in the plane of the bounded web
'
'   Based on the EndCondition and projected heights of the profiles select webcut,
'   selects a webcut when the bounding profile's flange is potentially in the
'   the way.
Public Sub WebCutFlanged(oBoundedPart As Object, _
                         strEndCutType As String, pSLH As IJDSelectorLogic)

    Dim strError As String
    
    ' Verify the BoundingObject exists
    Dim oBoundingObject As Object
    Set oBoundingObject = pSLH.InputObject(INPUT_BOUNDING)
    If (oBoundingObject Is Nothing) Then
        strError = " pSLH.InputObject(INPUT_BOUNDING) is NOTHING"
        GoTo ErrorHandler
    End If
    
    ' Verify the BoundingObject is an IJPort (not a FreeEndCut)
    If Not TypeOf oBoundingObject Is IJPort Then
        Exit Sub
    End If
    
    Dim oBoundingPart As Object
    Dim oBoundingPort As IJPort
    Set oBoundingPort = oBoundingObject
    Set oBoundingPart = oBoundingPort.Connectable
    
    Dim oSDO_Bounding  As Object
    Dim dBoundingHeight As Double
    Dim dFlangeThickness As Double
    Dim alpha As Double
    Dim dist As Double
    Dim phelper As New StructDetailObjects.Helper

    'get the answer for theta
    Dim sTheta As String
    sTheta = pSLH.SelectorAnswer(PROJECT_NAME + ".WebCutSel", "Theta")
    Dim dTheta As Double
    dTheta = Val(sTheta)

    If phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
        Dim oBoundingProfile As New StructDetailObjects.ProfilePart
        Set oBoundingProfile.object = oBoundingPart
        dBoundingHeight = oBoundingProfile.Height
        dFlangeThickness = oBoundingProfile.flangeThickness
        If oBoundingProfile.sectionType = gsB Then
            dist = 0.01
        Else
            dist = 0.035
        End If
        Set oSDO_Bounding = oBoundingProfile
        
     ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_BEAM Then
        Dim oBoundingBeam As New StructDetailObjects.BeamPart
        Set oBoundingBeam.object = oBoundingPart
        dBoundingHeight = oBoundingBeam.Height
        dFlangeThickness = oBoundingBeam.flangeThickness
        If oBoundingBeam.sectionType = gsB Then
            dist = 0.01
        Else
            dist = 0.035
        End If
        Set oSDO_Bounding = oBoundingBeam
    
    Else
            dist = 0.035
    End If
    
    Dim oBoundedObject As Object
    Set oBoundedObject = pSLH.InputObject(INPUT_BOUNDED)
    If (oBoundedObject Is Nothing) Then
        strError = " pSLH.InputObject(INPUT_BOUNDED) is NOTHING"
        GoTo ErrorHandler
    End If
    
    ' Verify the BoundingObject is an IJPort (not a FreeEndCut)
    If Not TypeOf oBoundedObject Is IJPort Then
        Exit Sub
    End If
    
'    If phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
'        'Profile bounded by profile case.
'        'If the landing curves are not intersecting, we would go with a simple web cut.
'        If Not AreLandingCurvesIntersecting(oBoundedPart, oBoundingPart) Then
'            pSLH.Add "WebAndFlangeWelded"
'            Exit Sub
'        End If
'    End If
        
    Dim oBoundedPort As IJPort
    Set oBoundedPort = oBoundedObject
    
    'Get the name of the port, used to determine side for the symbol
    Dim dPortValue As Long
    dPortValue = phelper.GetBoundingProfileFace(oBoundedPort, oBoundingPort)
   
    'check to see if the bounded part is a beam or a stiffener
    Dim dBoundedHeight As Double
    Dim strBoundedSectionType As String
    
    If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
        Dim oBoundedProfile As New StructDetailObjects.ProfilePart
        Set oBoundedProfile.object = oBoundedPart
        dBoundedHeight = oBoundedProfile.Height
    
        Call GetMeasurementSymbolData(pSLH, oBoundedProfile, oSDO_Bounding, dPortValue, alpha)
        
        Select Case strEndCutType
            Case "Snip"
                'for bounded profiles, do not check the theta
                If ((oBoundedProfile.sectionType = gsB) Or _
                   (oBoundedProfile.sectionType = "FB")) Then
                    pSLH.Add "Snip"
                Else
                    pSLH.Add "SnipWithFlange"
                End If
                
            Case "Cutback"
                pSLH.Add "StraightSnip"
                
             Case "Clip"
                pSLH.Add "Clip"
                
            Case "Welded", "Bracketed"
                If ((oBoundedProfile.sectionType = "BUT") Or _
                    (oBoundedProfile.sectionType = "BUTL3") Or _
                    (oBoundedProfile.sectionType = "T_XType") Or _
                    (oBoundedProfile.sectionType = "TSType") Or _
                    (oBoundedProfile.sectionType = "BUTL2")) And _
                    (dBoundingHeight = dBoundedHeight) Then
                    If 1.588249619 >= alpha And alpha >= 1.553343034 Then
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Left"
                            ElseIf (oSDO_Bounding.sectionType = "T_XType") Or _
                                    (oSDO_Bounding.sectionType = "TSType") Then
                                pSLH.Add "WebCutT_W1Left"
                            Else
                                pSLH.Add "CopeBUTLeft"
                                pSLH.Add "CopeTopCornerSnipe"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.sectionType = gsB) Then
                                pSLH.Add "CopeBulb"
                            ElseIf (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Right"
                            ElseIf (oSDO_Bounding.sectionType = "T_XType") Or _
                                   (oSDO_Bounding.sectionType = "TSType") Then
                                pSLH.Add "WebCutT_W1Right"
                            Else
                                pSLH.Add "CopeBUTRight"
                                pSLH.Add "CopeTopCornerSnipe"
                            End If
                        ElseIf dPortValue = JXSEC_TOP Or _
                            dPortValue = JXSEC_BOTTOM Or _
                            dPortValue = JXSEC_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                            dPortValue = JXSEC_WEB_LEFT_TOP Or _
                            dPortValue = JXSEC_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_LEFT_BOTTOM Then
                            pSLH.Add "Clip"
                        End If
                    ElseIf alpha < 1.553343034 Then
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Left"
                            Else
                                pSLH.Add "WebCut_C3Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.sectionType = gsB) Then
                                pSLH.Add "CopeBulb"
                            ElseIf (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Right"
                            Else
                                pSLH.Add "WebCut_C3Right"
                            End If
                        ElseIf dPortValue = JXSEC_TOP Or _
                            dPortValue = JXSEC_BOTTOM Or _
                            dPortValue = JXSEC_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                            dPortValue = JXSEC_WEB_LEFT_TOP Or _
                            dPortValue = JXSEC_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_LEFT_BOTTOM Then
                            pSLH.Add "Clip"
                        End If
                    ElseIf alpha > 1.588249619 Then
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Left"
                            Else
                                pSLH.Add "WebCut_C3Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Right"
                            Else
                                pSLH.Add "WebCut_C3Right"
                            End If
                        ElseIf dPortValue = JXSEC_TOP Or _
                            dPortValue = JXSEC_BOTTOM Or _
                            dPortValue = JXSEC_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                            dPortValue = JXSEC_WEB_LEFT_TOP Or _
                            dPortValue = JXSEC_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_LEFT_BOTTOM Then
                            pSLH.Add "Clip"
                        End If
                    End If
                Else 'the bounded object is not a built-up or tee section
                    If (dBoundingHeight - dFlangeThickness - dist) >= dBoundedHeight Then
                        pSLH.Add "Clip"
                    Else
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Left"
                            Else
                                pSLH.Add "WebCut_C3Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.sectionType = gsB) Then
                                pSLH.Add "CopeBulb"
                            Else
                                If (oSDO_Bounding.sectionType = "I") Or _
                                    (oSDO_Bounding.sectionType = "ISType") Or _
                                    (oSDO_Bounding.sectionType = "H") Then
                                    pSLH.Add "WebCutI_C3Right"
                                Else
                                    pSLH.Add "Cope_AngleRight"
                                End If
                            End If
                        ElseIf dPortValue = JXSEC_TOP Or _
                            dPortValue = JXSEC_BOTTOM Or _
                            dPortValue = JXSEC_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                            dPortValue = JXSEC_WEB_LEFT_TOP Or _
                            dPortValue = JXSEC_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_LEFT_BOTTOM Then
                            pSLH.Add "Clip"
                        End If
                    End If
                End If
                
            Set phelper = Nothing
            
        End Select
        
    ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        Dim oBoundedBeam As New StructDetailObjects.BeamPart
        Set oBoundedBeam.object = oBoundedPart
        dBoundedHeight = oBoundedBeam.Height
        strBoundedSectionType = oBoundedBeam.sectionType
        
        Call GetMeasurementSymbolData(pSLH, oBoundedBeam, oSDO_Bounding, dPortValue, alpha)
        
        Select Case strEndCutType
            
            Case "Snip"
                'for bounded profiles, do not check the theta
                If strBoundedSectionType = gsB Or _
                   strBoundedSectionType = "FB" Then
                    pSLH.Add "Snip"
                Else
                    pSLH.Add "SnipWithFlange"
                End If
                
            Case "Cutback"
                pSLH.Add "StraightSnip"
                
             Case "Clip"
                pSLH.Add "Clip"
                
            Case "Welded", "Bracketed"
                If ((strBoundedSectionType = "BUT") Or _
                    (strBoundedSectionType = "BUTL3") Or _
                    (strBoundedSectionType = "T_XType") Or _
                    (strBoundedSectionType = "TSType") Or _
                    (strBoundedSectionType = "BUTL2")) And _
                    (dBoundingHeight = dBoundedHeight) Then
                    If 1.588249619 >= alpha And alpha >= 1.553343034 Then
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Left"
                            ElseIf (oSDO_Bounding.sectionType = "T_XType") Or _
                                    (oSDO_Bounding.sectionType = "TSType") Then
                                pSLH.Add "WebCutT_W1Left"
                            Else
                                pSLH.Add "CopeBUTLeft"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.sectionType = gsB) Then
                                pSLH.Add "CopeBulb"
                            ElseIf (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Right"
                            ElseIf (oSDO_Bounding.sectionType = "T_XType") Or _
                                   (oSDO_Bounding.sectionType = "TSType") Then
                                pSLH.Add "WebCutT_W1Right"
                            Else
                                pSLH.Add "CopeBUTRight"
                            End If
                        ElseIf dPortValue = JXSEC_TOP Or _
                            dPortValue = JXSEC_BOTTOM Or _
                            dPortValue = JXSEC_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                            dPortValue = JXSEC_WEB_LEFT_TOP Or _
                            dPortValue = JXSEC_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_LEFT_BOTTOM Then
                            pSLH.Add "Clip"
                        End If
                    ElseIf alpha < 1.553343034 Then
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Left"
                            Else
                                pSLH.Add "WebCut_C3Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.sectionType = gsB) Then
                                pSLH.Add "CopeBulb"
                            ElseIf (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Right"
                            Else
                                pSLH.Add "WebCut_C3Right"
                            End If
                        ElseIf dPortValue = JXSEC_TOP Or _
                            dPortValue = JXSEC_BOTTOM Or _
                            dPortValue = JXSEC_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                            dPortValue = JXSEC_WEB_LEFT_TOP Or _
                            dPortValue = JXSEC_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_LEFT_BOTTOM Then
                            pSLH.Add "Clip"
                        End If
                    ElseIf alpha > 1.588249619 Then
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Left"
                            Else
                                pSLH.Add "WebCut_C3Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Right"
                            Else
                                pSLH.Add "WebCut_C3Right"
                            End If
                        ElseIf dPortValue = JXSEC_TOP Or _
                            dPortValue = JXSEC_BOTTOM Or _
                            dPortValue = JXSEC_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                            dPortValue = JXSEC_WEB_LEFT_TOP Or _
                            dPortValue = JXSEC_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_LEFT_BOTTOM Then
                            pSLH.Add "Clip"
                        End If
                    End If
                Else 'the bounded object is not a built-up or tee section
                    If (dBoundingHeight - dFlangeThickness - dist) >= dBoundedHeight Then
                        pSLH.Add "Clip"
                    Else
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.sectionType = "I") Or _
                                (oSDO_Bounding.sectionType = "ISType") Or _
                                (oSDO_Bounding.sectionType = "H") Then
                                pSLH.Add "WebCutI_C3Left"
                            Else
                                pSLH.Add "WebCut_C3Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.sectionType = gsB) Then
                                pSLH.Add "CopeBulb"
                            Else
                                If (oSDO_Bounding.sectionType = "I") Or _
                                    (oSDO_Bounding.sectionType = "ISType") Or _
                                    (oSDO_Bounding.sectionType = "H") Then
                                    pSLH.Add "WebCutI_C3Right"
                                Else
                                    pSLH.Add "Cope_AngleRight"
                                End If
                            End If
                        ElseIf dPortValue = JXSEC_TOP Or _
                            dPortValue = JXSEC_BOTTOM Or _
                            dPortValue = JXSEC_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                            dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                            dPortValue = JXSEC_WEB_LEFT_TOP Or _
                            dPortValue = JXSEC_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                            dPortValue = JXSEC_WEB_LEFT_BOTTOM Then
                            pSLH.Add "Clip"
                        End If
                    End If
                End If
            
            
            Set phelper = Nothing
            
        End Select
  
    End If
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "WebCutFlanged", strError).Number
End Sub


Public Sub WebCutNoFlangeEndToEnd(oBounded As Object, bIsPlate As Boolean, _
                           strEndCutType As String, strWeldPartNumber As String, _
                           pSLH As IJDSelectorLogic)
'
'   Base on the EndCondition and projected heights of the profiles select webcut.
'
    Dim strError As String
    
    On Error GoTo ErrorHandler
    
    ' Verify the BoundingObject exists
    Dim oBoundingObject As Object
    Set oBoundingObject = pSLH.InputObject(INPUT_BOUNDING)
    If (oBoundingObject Is Nothing) Then
        strError = " pSLH.InputObject(INPUT_BOUNDING) is NOTHING"
        GoTo ErrorHandler
    End If
    
zMsgBox "EndToEndWebCutSel::WebCutNoFlangeEndToEnd" & vbCrLf & _
        "   strEndCutType: " & strEndCutType
        
    ' Verify the BoundingObject is an IJPort (not a FreeEndCut)
    If TypeOf oBoundingObject Is IJPort Then
       ' Check if it is tapered case:
       ' 1. Bounded and bounding parts are from the same root Edge centered FB ER
       ' 2. Web lengths are different
       Dim oBoundingPort As IJPort
      
       Dim dBoundedWebLen As Double
       Dim dBoundingWebLen As Double

       Dim sBoundedXSectionType As String
       Dim sBoundingXSectionType As String
       Dim oBounding As Object
       Dim bIsTaperedCase As Boolean
       Dim bIsAlligatorCase As Boolean
       Dim dBoundedHeight As Double
       Dim dBoundingHeight As Double
       Dim dBoundedFT As Double
       Dim dBoundingFT As Double
       
       bIsAlligatorCase = False
       bIsTaperedCase = False
       If TypeOf oBounded Is IJStiffenerPart Then
          Dim oProfilePartWrapper As New StructDetailObjects.ProfilePart
          
          Set oProfilePartWrapper.object = oBounded
          dBoundedWebLen = oProfilePartWrapper.WebLength
          dBoundedWebLen = Int(dBoundedWebLen * 1000)
          dBoundedHeight = Int(oProfilePartWrapper.Height * 1000)
          dBoundedFT = Int(oProfilePartWrapper.flangeThickness * 1000)
          
          sBoundedXSectionType = oProfilePartWrapper.sectionType
         
          Set oBoundingPort = oBoundingObject
          Set oBounding = oBoundingPort.Connectable
          Set oProfilePartWrapper.object = oBounding
          
          dBoundingWebLen = oProfilePartWrapper.WebLength
          dBoundingWebLen = Int(dBoundingWebLen * 1000)
          dBoundingHeight = Int(oProfilePartWrapper.Height * 1000)
          dBoundingFT = Int(oProfilePartWrapper.flangeThickness * 1000)
          
          sBoundingXSectionType = oProfilePartWrapper.sectionType
       Else
          Dim oBeamPartWrapper As New StructDetailObjects.BeamPart
          
          Set oBeamPartWrapper.object = oBounded
          dBoundedWebLen = oBeamPartWrapper.WebLength
          dBoundedWebLen = Int(dBoundedWebLen * 1000)
          dBoundedHeight = Int(oBeamPartWrapper.Height * 1000)
          dBoundedFT = Int(oBeamPartWrapper.flangeThickness * 1000)
          
          sBoundedXSectionType = oBeamPartWrapper.sectionType
         
          Set oBoundingPort = oBoundingObject
          Set oBounding = oBoundingPort.Connectable
          Set oBeamPartWrapper.object = oBounding
          
          dBoundingWebLen = oBeamPartWrapper.WebLength
          dBoundingWebLen = Int(dBoundingWebLen * 1000)
          dBoundingHeight = Int(oBeamPartWrapper.Height * 1000)
          dBoundingFT = Int(oBeamPartWrapper.flangeThickness * 1000)
          
          sBoundingXSectionType = oBeamPartWrapper.sectionType
       End If
       
       If sBoundedXSectionType = "FB" And sBoundingXSectionType = "FB" And _
          TypeOf oBounded Is IJProfileER And _
          TypeOf oBoundingPort.Connectable Is IJProfileER Then
           Dim oERWrapper As New StructDetailObjects.EdgeReinforcement
           Dim sBoundedPosition As String
           Dim sBoundingPosition As String
           Dim oBoundedRootSystem As Object
           Dim oBoundingRootSystem As Object
           Dim oSDHelper As New StructDetailObjects.Helper
           
           Set oERWrapper.object = oBounded
           sBoundedPosition = oERWrapper.ReinforcementPosition
           sBoundedPosition = Trim(sBoundedPosition)
           Set oBoundedRootSystem = oSDHelper.Object_RootParentSystem(oBounded)
           
           Set oERWrapper.object = oBounding
           sBoundingPosition = oERWrapper.ReinforcementPosition
           sBoundingPosition = Trim(sBoundingPosition)
           Set oBoundingRootSystem = oSDHelper.Object_RootParentSystem(oBounding)
   
           If LCase(sBoundedPosition) = LCase("On Edge - Centered") And _
              LCase(sBoundingPosition) = LCase("On Edge - Centered") And _
              oBoundedRootSystem Is oBoundingRootSystem And _
              dBoundedWebLen <> dBoundingWebLen Then
   
              bIsTaperedCase = True
           End If
       End If
   
       If sBoundedXSectionType = "BUT" Or sBoundingXSectionType = "BUTL2" Then
          If dBoundedHeight <> dBoundingHeight And _
             dBoundedFT < dBoundingHeight Then
              bIsAlligatorCase = True
          End If
       End If

       Dim oChild As IJDesignChild
       Dim oACWrapper As New StructDetailObjects.AssemblyConn
       Dim oACBounded As Object
       Dim dLengthV As Double
       Dim dSlope As Double
       Dim dAngle As Double
       
       ' Default alligator cut angle to 30 degrees
       dAngle = (3.1415926 * 30) / 180
   
       ' Default cut slope near flange to 0.01 degrees
       dSlope = (3.1415916 * 0.01) / 180
       
       Set oChild = pSLH.SmartOccurrence
       Set oACWrapper.object = oChild.GetParent
       Set oACBounded = oACWrapper.ConnectedObject1
    
        Select Case strEndCutType
            Case "Welded"
                If strWeldPartNumber = "First" Then
                    If bIsTaperedCase = True Then
                        If oBounded Is oACBounded Then
                            If dBoundedWebLen > dBoundingWebLen Then
                                ' Make tapered web cut avaliable
                                pSLH.Add "WebCut_C1TaperedPCEnd"
                            End If
                        End If
                    End If
                
                    
zMsgBox "EndToEndWebCutSel::WebCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add WebCut_C1PCEnd"
                    If bIsAlligatorCase = True Then
                        If oBounded Is oACBounded Then
                            If dBoundedWebLen > dBoundingWebLen Then
                                pSLH.Add "WebCut_C1AlligatorPCEnd"
                            End If
                        End If
                    End If
                ' add standard case as option
                pSLH.Add "WebCut_C1PCEnd"
                Else
                    If bIsTaperedCase = True Then
                        If oBounding Is oACBounded Then
                            If dBoundedWebLen > dBoundingWebLen Then
                                pSLH.Add "WebCut_C1TaperedEnd"
                            End If
                        End If
                    End If
                                    
                    
zMsgBox "EndToEndWebCutSel::WebCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add WebCut_C1End"
        
                    If bIsAlligatorCase = True Then
                        If oBounding Is oACBounded Then
                            If dBoundedWebLen > dBoundingWebLen Then
                                pSLH.Add "WebCut_C1AlligatorEnd"
                            End If
                        End If
                    End If
                ' add standard case as option
                pSLH.Add "WebCut_C1End"
                End If
                
            Case Else
zMsgBox "EndToEndWebCutSel::WebCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add (Nothing)"
            End Select
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "WebCutNoFlangeEndToEnd", strError).Number
End Sub

Public Sub WebCutEndToEndAngle(oBounded As Object, _
                               bIsPlate As Boolean, _
                               strEndCutType As String, _
                               strWeldPartNumber As String, _
                               sSplitAngleCase As String, _
                               sFlipAngle As String, _
                               pSLH As IJDSelectorLogic)
'
'   Base on the EndCondition and projected heights of the profiles select webcut.
'
    Dim strError As String
    
    On Error GoTo ErrorHandler
    
    Dim oPortBounding As IJPort
    Set oPortBounding = pSLH.InputObject(INPUT_BOUNDING)
    If (oPortBounding Is Nothing) Then
        strError = " pSLH.InputObject(INPUT_BOUNDING) is NOTHING"
        GoTo ErrorHandler
    End If
    
    On Error GoTo ErrorHandler
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oBounded)
    
zMsgBox "::WebCutEndToEndAngle" & vbCrLf & _
        "   strEndCutType: " & strEndCutType & vbCrLf & _
        "   sSplitAngleCase: " & sSplitAngleCase & vbCrLf & _
        "   sFlipAngle: " & sFlipAngle

    Dim sAddWebCutId As String
    sAddWebCutId = ""
    
    Dim sPartNumber As String
    sPartNumber = strWeldPartNumber
    If sFlipAngle = "Flip" Then
        If strWeldPartNumber = "First" Then
            sPartNumber = "Second"
        Else
            sPartNumber = "First"
        End If
        
    End If
    
    Select Case strEndCutType
        Case "Welded"
        
            ' for FB :
            '           AngleWebSquareFlange(Case1)
            '           AngleWebBevelFlange(Case2)
            '           AngleWebAngleFlange(Case3)
            '   WebCut Only
            '   based on WebCut Angle
            If sSectionType = "FB" Then
                If LCase(Trim(sSplitAngleCase)) = LCase("AngleWebSquareFlange") Or _
                   LCase(Trim(sSplitAngleCase)) = ("AngleWebBevelFlange") Or _
                   LCase(Trim(sSplitAngleCase)) = LCase("AngleWebAngleFlange") Then
                    If sPartNumber = "First" Then
                        sAddWebCutId = "WebCut_PC_SeamAngleCase2_FB"
                    Else
                        sAddWebCutId = "WebCut_SeamAngleCase2_FB_2"
                    End If
            
                ' for FB :
                '           DistanceWebDistanceFlange(Case4)
                '           OffsetWebOffsetFlange(Case5)
                '   WebCut Only
                '   based on WebCut and FlangeCut Offset distances
                ElseIf LCase(Trim(sSplitAngleCase)) = LCase("DistanceWebDistanceFlange") Or _
                       LCase(Trim(sSplitAngleCase)) = LCase("OffsetWebOffsetFlange") Then
                    If sPartNumber = "First" Then
                        sAddWebCutId = "WebCut_PC_SeamAngleCase4_FB"
                    Else
                        sAddWebCutId = "WebCut_SeamAngleCase4_FB_2"
                    End If
                End If
            
            ' for all other Cross Section Types:
            Else
                ' AngleWebSquareFlange (Case1) WebCut only
                '   based on WebCut and FlangeCut Angles
                '   cuts both Web and Flange
                If LCase(Trim(sSplitAngleCase)) = LCase("AngleWebSquareFlange") Then
                    If sPartNumber = "First" Then
                        sAddWebCutId = "WebCut_PC_SeamAngleCase1"
                    Else
                        sAddWebCutId = "WebCut_SeamAngleCase1_2"
                    End If
            
                ' AngleWebBevelFlange (Case2) WebCut Only
                '   based on WebCut Angle
                '   cuts both Web and Flange
                ElseIf LCase(Trim(sSplitAngleCase)) = LCase("AngleWebBevelFlange") Then
                    If sPartNumber = "First" Then
                        sAddWebCutId = "WebCut_PC_SeamAngleCase2"
                    Else
                        sAddWebCutId = "WebCut_SeamAngleCase2_2"
                    End If
            
                ' AngleWebAngleFlange (Case3) WebCut and FlangeCut
                '   based on WebCut Angle
                '   cuts both Web only
                ' Note: Case 3 is currently not implemented
                ElseIf LCase(Trim(sSplitAngleCase)) = LCase("AngleWebAngleFlange") Then
                    If sPartNumber = "First" Then
                        sAddWebCutId = "WebCut_PC_SeamAngleCase3"
                    Else
                        sAddWebCutId = "WebCut_SeamAngleCase3_2"
                    End If
            
                ' DistanceWebDistanceFlange (Case4)
                '   based on WebCut and FlangeCut Offset distances
                '   cuts both Web and Flange
                ElseIf LCase(Trim(sSplitAngleCase)) = LCase("DistanceWebDistanceFlange") Then
                    If sPartNumber = "First" Then
                        sAddWebCutId = "WebCut_PC_SeamAngleCase4"
                    Else
                        sAddWebCutId = "WebCut_SeamAngleCase4_2"
                    End If
                
                ' OffsetWebOffsetFlange (Case5) WebCut Only
                '   based on WebCut angle and FlangeCut Offset distance
                '   cuts both Web and Flange
                ElseIf LCase(Trim(sSplitAngleCase)) = LCase("OffsetWebOffsetFlange") Then
                    If sPartNumber = "First" Then
                        sAddWebCutId = "WebCut_PC_SeamAngleCase5"
                    Else
                        sAddWebCutId = "WebCut_SeamAngleCase5_2"
                    End If
                End If
            End If
            
            ' Offset web cut,a special case of OffsetWebOffsetFlange with OffsetWeb = OffsetFlange
            ' It does not care about cross section type. Offset inside cut always has PC
            If LCase(Trim(sSplitAngleCase)) = LCase("NoAngleOffset") Then
                If sPartNumber = "First" Then
                    sAddWebCutId = "WebCut_OffsetInside"
                Else
                    sAddWebCutId = "WebCut_OffsetOutside"
                End If
            End If
            
        Case Else
        
               
        End Select
    
    If Len(Trim(sAddWebCutId)) > 0 Then
        pSLH.Add sAddWebCutId
        zMsgBox "::WebCutEndToEndAngle... pSL.Add " & sAddWebCutId
    Else
        zMsgBox "EndToEndWebCutSel::WebCutEndToEndAngle... pSL.Add (Nothing)"
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "WebCutNoFlangeEndToEnd", strError).Number
End Sub

Public Sub GetMeasurementSymbolData(pSLH As IJDSelectorLogic, _
                                    oSDO_Bounded As Object, _
                                    oSDO_Bounding As Object, _
                                    dPortValue As Long, ByRef dAngle As Double)
                           
'Use the input objects and the port name to get the measurement and compute
'necessary values

    On Error GoTo ErrorHandler
    
    If dPortValue = JXSEC_WEB_LEFT Or dPortValue = JXSEC_TOP_FLANGE_LEFT Or dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Or _
       dPortValue = JXSEC_WEB_RIGHT Or dPortValue = JXSEC_TOP_FLANGE_RIGHT Or dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
        'depending on the cross section type, use measurement symbol to get bounding profile angle
        If (oSDO_Bounding.sectionType = gsB) Or _
            (oSDO_Bounding.sectionType = "UA") Or _
            (oSDO_Bounding.sectionType = "EA") Or _
            (oSDO_Bounding.sectionType = "BUT") Or _
            (oSDO_Bounding.sectionType = "BUTL2") Or _
            (oSDO_Bounding.sectionType = "TSType") Or _
            (oSDO_Bounding.sectionType = "T_XType") Or _
            (oSDO_Bounding.sectionType = "BUTL3") Then
                       
            '**********************************************************************************
            ' Measurement Symbol for getting angle of the bounding profile
            '**********************************************************************************
            'dim the base plate, which will be input to the symbol file
            Dim oBasePlate As IJPlate
            Dim bIsSystem As Boolean
            oSDO_Bounded.GetStiffenedPlate oBasePlate, bIsSystem
        
            'Get web cut object
            Dim oWebCut As StructDetailObjects.WebCut
            Set oWebCut = New WebCut
            Set oWebCut.object = pSLH.SmartOccurrence
        
            Dim MSSymbol As New StructDetailObjects.Measurement
            'Initialize MS Symbol by getting the resource Manager of a Persistent Object and then using it
            Dim pIJDObject As IJDObject
            Set pIJDObject = pSLH.SmartOccurrence
            MSSymbol.Initialize pIJDObject.ResourceManager
        
            ' Get SketchPlane
            ' using the Web_Right and Top edge planes
            ' calculate the Cross Section Sketch Plane
            Dim oTopPlane As IJPlane
            Dim oRightPlane As IJPlane
            Dim oSketchPlane As IJPlane
            Set oTopPlane = oSDO_Bounding.EdgePlane(JXSEC_TOP, oWebCut.BoundedLocation, pIJDObject.ResourceManager)
            Set oRightPlane = oSDO_Bounding.EdgePlane(dPortValue, oWebCut.BoundedLocation, pIJDObject.ResourceManager)
        
            Dim dXtop As Double, dYtop As Double, dZtop As Double
            Dim dXright As Double, dYright As Double, dZright As Double
            Dim dXsketch As Double, dYsketch As Double, dZsketch As Double
            oTopPlane.GetNormal dXtop, dYtop, dZtop
            oRightPlane.GetNormal dXright, dYright, dZright
            
            Dim oTopNormal As AutoMath.DVector
            Dim oRightNormal As AutoMath.DVector
            Dim oSketchNormal As AutoMath.DVector
            Set oTopNormal = New AutoMath.DVector
            Set oRightNormal = New AutoMath.DVector
            
            oTopNormal.Set dXtop, dYtop, dZtop
            oRightNormal.Set dXright, dYright, dZright
            Set oSketchNormal = oRightNormal.Cross(oTopNormal)
            oSketchNormal.Get dXsketch, dYsketch, dZsketch
            
            Set oSketchPlane = oRightPlane
            oSketchPlane.SetNormal dXsketch, dYsketch, dZsketch
            oSketchPlane.SetUDirection dXright, dYright, dZright
            '===================================================================
            ' TR65588
            '   BUT measurement sym value returned wrong due to angle of base plate
            ' The Sketch Plane calculated above is based on the Bounding Profile
            ' such that the Profile Edges will be mainly vertical and horizontal
            ' while the Base Plate may be angled
            '
            ' If the users desires that the Profile Edges are angled and
            ' the Base Plate be mainly horizontal
            ' Need to calculate the Sketching Plane U direction based on the
            ' Stiffened Plate geometry by:
            '   Retreiving the Base (Stiffened) Plate (System or Part)
            '   Retreiving the Base (Stiffened) Plate Geometry
            '   Retreiving a Normal vector from the Base Plate Geometry
            '   Cross the Base Plate Normal with the Sketching Plane Normal
            '   Adjust for direction using the Web Right Face Normal
            Dim oBasePoint As IJDPosition
            Dim oBaseVector As IJDVector
            Dim oBaseGeometry As Object
            Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
            Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
            Set oBaseGeometry = oTopologyLocate.GetPlateParentBodyModel(oBasePlate)
            oTopologyLocate.GetProjectedPointOnModelBody oBaseGeometry, _
                                                         oWebCut.BoundedLocation, _
                                                         oBasePoint, _
                                                         oBaseVector
            
            Dim dDot As Double
            Dim oBaseNormal As AutoMath.DVector
            Set oBaseNormal = New AutoMath.DVector
            Set oBaseNormal = oBaseVector.Cross(oSketchNormal)
            oBaseNormal.Get dXright, dYright, dZright
            dDot = oBaseNormal.Dot(oRightNormal)
            If dDot < 0 Then
                dXright = -dXright
                dYright = -dYright
                dZright = -dZright
            End If
            
'
'          oSketchPlane.SetUDirection dXright, dYright, dZright
        'pblm - can't set SketchPlane.Udirection if BaseVec X SketchNormal = 0
'solution is to leave it as is
            Dim oGeometryServices As IGeometryServices
            Set oGeometryServices = New IngrGeom3D.GeometryFactory

            dDot = (dXright * dXright) + (dYright * dYright) + (dZright * dZright)
            If dDot > oGeometryServices.DistTolerance Then 'okay to set UDir
                oSketchPlane.SetUDirection dXright, dYright, dZright 'this simply "adjusts" the UDir
            Else ' what if set obasenormal=otopnormal
'do nothing
            End If
    
                    
            Dim ssymbolname As String
            Select Case oSDO_Bounding.sectionType
                Case gsB
                    ssymbolname = "Measurement\CrossSectionB.sym"
                    'Set the input parameter for the symbol file
                    MSSymbol.AddInputParameter "WebOffset", 0.02
                Case "UA", "EA"
                    ssymbolname = "MarineLibrary\Measurement\CrossSectionUA.sym"
                    'Set the input parameter for the symbol file
                    MSSymbol.AddInputParameter "WebOffset", 0.02
                Case "BUT", "BUTL2"
                    ssymbolname = "Measurement\CrossSectionBUT.sym"
                Case "TSType", "T_XType"
                    ssymbolname = "Measurement\CrossSectionT.sym"
                Case "BUTL3"
                    ssymbolname = "Measurement\CrossSectionBUTL3.sym"
            End Select
            
            ' Call the Measurement Symbol
            MSSymbol.ComputeCrossSection ssymbolname, oWebCut.Bounding, oBasePlate, _
                                        oSketchPlane
            
            ' Delete Persistent Edge Planes created
            Set pIJDObject = oTopPlane
            pIJDObject.Remove

            Set pIJDObject = oRightPlane
            pIJDObject.Remove

            ' Get Data out of Measurement Symbol
            dAngle = MSSymbol.GetOutputParameter("WebRightAngle")
            
        Else
            dAngle = 1.5707963267949
        End If
    End If
            
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetMeasurementSymbolData").Number
End Sub

Public Function GetCrossSectionType(ByRef oProfilePartObject As Object) As String
    On Error GoTo ErrorHandler

    Dim sSectionType As String
    sSectionType = ""
    
    If TypeOf oProfilePartObject Is IJStiffener Then
        Dim oSDO_Profile As StructDetailObjects.ProfilePart
        Set oSDO_Profile = New StructDetailObjects.ProfilePart
        Set oSDO_Profile.object = oProfilePartObject
        
        sSectionType = oSDO_Profile.sectionType
        
        Set oSDO_Profile = Nothing
        Set oSDO_Profile = Nothing
        
    ElseIf TypeOf oProfilePartObject Is IJBeam Then
        Dim oSDO_Beam As StructDetailObjects.BeamPart
        Set oSDO_Beam = New StructDetailObjects.BeamPart
        Set oSDO_Beam.object = oProfilePartObject
        
        sSectionType = oSDO_Beam.sectionType
        
        Set oSDO_Beam = Nothing
        Set oSDO_Beam = Nothing
    End If
        
    GetCrossSectionType = sSectionType
zMsgBox "Common::GetCrossSectionType ...:" & sSectionType
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetCrossSectionType").Number
End Function

Public Function GetBoundedCutDepth(oBoundedObject As Object, _
                                   bUseFlangeData As Boolean) As Double
    On Error GoTo ErrorHandler

    Dim dCuttingDepth As Double
    
    Dim oBoundedPort As IJPort
    Dim oBoundedConnectable As Object
    
    Dim oSDO_Helper As StructDetailObjects.Helper
    Set oSDO_Helper = New StructDetailObjects.Helper
     
    dCuttingDepth = 0.01
    If TypeOf oBoundedObject Is IJPort Then
        Set oBoundedPort = oBoundedObject
        Set oBoundedConnectable = oBoundedPort.Connectable
        Set oBoundedPort = Nothing
    ElseIf TypeOf oBoundedObject Is IJConnectable Then
        Set oBoundedConnectable = oBoundedObject
    Else
        Set oBoundedConnectable = oBoundedObject
    End If
    
    Dim sObjectType As sdwObjectType
    sObjectType = oSDO_Helper.ObjectType(oBoundedConnectable)
    If sObjectType = SDOBJECT_STIFFENER Then
        Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
        Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
        Set oSDO_ProfilePart.object = oBoundedConnectable
        If bUseFlangeData Then
            dCuttingDepth = oSDO_ProfilePart.Width
            If dCuttingDepth < 0.001 Then
                dCuttingDepth = oSDO_ProfilePart.WebThickness
            End If
        Else
            If oSDO_ProfilePart.sectionType = "P" Or _
               oSDO_ProfilePart.sectionType = "HalfR" Or _
               oSDO_ProfilePart.sectionType = "R" Then
                dCuttingDepth = 0.6 '600 mm
            Else
                dCuttingDepth = oSDO_ProfilePart.WebThickness
            End If
        End If
        Set oSDO_ProfilePart = Nothing
    
    ElseIf sObjectType = SDOBJECT_BEAM Then
        Dim oSDO_BoundedBeam As StructDetailObjects.BeamPart
        Set oSDO_BoundedBeam = New StructDetailObjects.BeamPart
        Set oSDO_BoundedBeam.object = oBoundedConnectable
        If bUseFlangeData Then
            dCuttingDepth = oSDO_BoundedBeam.Width
            If dCuttingDepth < 0.001 Then
                dCuttingDepth = oSDO_BoundedBeam.WebThickness
            End If
        Else
            dCuttingDepth = oSDO_BoundedBeam.WebThickness
        End If
            
        Set oSDO_BoundedBeam = Nothing

    End If

    GetBoundedCutDepth = dCuttingDepth
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetBoundedCutDepth").Number
End Function

Public Function GetBoundingAngle(oWebCut As Object) As Double
    ' Get landing curves
    Dim oSDOWebCut As New StructDetailObjects.WebCut
    Set oSDOWebCut.object = oWebCut
    
    Dim oProfileUtil As IJProfileAttributes
    Set oProfileUtil = New ProfileUtils
    
    Dim oBoundedLC As IJWireBody
    Dim oBoundingLC As IJWireBody
    
    Dim oBoundedPort As IJPort
    Set oBoundedPort = oSDOWebCut.BoundedPort
    
    oProfileUtil.GetLandingCurveFromProfile oSDOWebCut.Bounded, oBoundedLC
    oProfileUtil.GetLandingCurveFromProfile oSDOWebCut.Bounding, oBoundingLC
    
    ' Get the bounded port normal
    Dim oSurfaceUtil As IJSurfaceBody
    Dim oModelUtil As IJSGOModelBodyUtilities
    Dim oPointOnPort As IJDPosition
    Dim oBoundedNormal As IJDVector
    Dim dist As Double
    
    Set oModelUtil = New SGOModelBodyUtilities
    oModelUtil.GetClosestPointOnBody oBoundedPort.Geometry, _
                                     oSDOWebCut.BoundedLocation, _
                                     oPointOnPort, _
                                     dist
    
    Set oSurfaceUtil = oBoundedPort.Geometry
    oSurfaceUtil.GetNormalFromPosition oPointOnPort, oBoundedNormal
    
    ' Get bounded primary orientation vector
    Dim oPartUtil As IJPartSupport
    Dim oProfilePartUtil As IJProfilePartSupport
    Set oProfilePartUtil = New ProfilePartSupport
    Set oPartUtil = oProfilePartUtil
    
    Set oPartUtil.Part = oSDOWebCut.Bounded
    
    Dim oPrimaryDir As IJDVector
    Dim oSecondaryDir As IJDVector
    Dim oOrigin As IJDPosition
    oProfilePartUtil.GetOrientation oSDOWebCut.BoundedLocation, oSecondaryDir, oPrimaryDir, oOrigin
            
    ' Get tangents at bounded position
    Dim oWireUtil As IJSGOWireBodyUtilities
    Set oWireUtil = New SGOWireBodyUtilities
    
    Dim oClosestPos As IJDPosition
    Dim oBoundedVec As IJDVector
    Dim oBoundingVec As IJDVector
    
    oWireUtil.GetClosestPointOnWire oBoundedLC, oSDOWebCut.BoundedLocation, oClosestPos, oBoundedVec
    oWireUtil.GetClosestPointOnWire oBoundingLC, oSDOWebCut.BoundedLocation, oClosestPos, oBoundingVec

    ' Reverse bounded tangent if not in same direction as bounded normal
    If oBoundedVec.Dot(oBoundedNormal) < 0# Then
        oBoundedVec.Length = -1#
    End If

    ' Reverse bounding tangenet if not in same direction as bounded primary orientation
    If oPrimaryDir.Dot(oBoundingVec) < 0# Then
        oBoundingVec.Length = -1#
    End If
    
    ' Measure angle between
    Dim angle As Double
    GetBoundingAngle = oBoundedVec.angle(oBoundingVec, oSecondaryDir)
    
    Dim dPI As Double
    dPI = Atn(1) * 4
    
    If GetBoundingAngle > dPI Then
        GetBoundingAngle = 2 * dPI - GetBoundingAngle
    End If

End Function

' ********************************************************************************
' Method: GetBoundingStiffenerHeight
'
' Description:
'   Retruens the Bounding Profile (Stiffener, Edge Reinforcement) Bounding height
'
' Inputs:
'   oBoundedPortObject:     Bounded Port (from Stiffener Part)
'   oBoundingPortObject:    Bounding Port (from Stiffener Part,Edge Reinforcement)
'
' Outputs:
'   GetBoundingStiffenerHeight:     Bounding Height
'
' For Stiffener Parts: Bounding Height is Stiffener Height
' For Edge Reinforcement Parts:
'   if Edge Reinforcement Part's Position is "On Edge",
'       Bounding Height is Edge Reinforcement Part Height
'   else
'       determine bounding height distance based on distances returned from
'       Edge Reinforcement wrapper method: GetPlateToTopBottom
' ********************************************************************************
Public Function GetBoundingStiffenerHeight(oBoundedPortObject As Object, _
                                           oBoundingPortObject As Object) As Double
    On Error GoTo ErrorHandler

    Dim sObjectType As String
    Dim sTypeObject As String
        
    Dim dDot As Double
    Dim dTopSetBack As Double
    Dim dBottomSetBack As Double
    Dim dBoundingHeight As Double
    
    Dim oPort As IJPort
    Dim oNormal As IJDVector
    Dim oEndPosition As IJDPosition
    
    Dim oPoint_MountingFace As IJDPosition
    Dim oNormal_MountingFace As IJDVector
    
    Dim oPoint_TopFace As IJDPosition
    Dim oNormal_TopFace As IJDVector
    
    Dim oBoundedProfile As Object
    Dim oBoundingStiffener As Object
    Dim oBoundedPortGeometry As Object
    
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oSDO_BeamPart As StructDetailObjects.BeamPart
    Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
    Dim oSDO_EdgeReinforcement As StructDetailObjects.EdgeReinforcement
    
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    
    ' Verify Objects passed in are valid IJPort objects
    GetBoundingStiffenerHeight = 0#
    
    If TypeOf oBoundedPortObject Is IJPort Then
        Set oPort = oBoundedPortObject
        Set oBoundedProfile = oPort.Connectable
        Set oBoundedPortGeometry = oPort.Geometry
        Set oPort = Nothing
    Else
        Exit Function
    End If
    
    If TypeOf oBoundingPortObject Is IJPort Then
        Set oPort = oBoundingPortObject
        Set oBoundingStiffener = oPort.Connectable
        Set oPort = Nothing
    Else
        Exit Function
    End If
    
    ' Check if Bounding Stiffener is an Edge Reinforcement
    Set oSDO_Helper = New StructDetailObjects.Helper
    oSDO_Helper.GetObjectTypeData oBoundingStiffener, sTypeObject, sObjectType
    If LCase(Trim(sTypeObject)) <> LCase("EdgeReinforcement") Then
    
        If InStr(LCase(Trim(sTypeObject)), LCase("Beam")) > 0 Then
            ' Retrieve the Bounding Beam's Height
            Set oSDO_BeamPart = New StructDetailObjects.BeamPart
            Set oSDO_BeamPart.object = oBoundingStiffener
            dBoundingHeight = oSDO_BeamPart.Height
        
        ElseIf InStr(LCase(Trim(sTypeObject)), LCase("Stiffener")) > 0 Then
            ' Retrieve the Bounding Stiffener's Height
            Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
            Set oSDO_ProfilePart.object = oBoundingStiffener
            dBoundingHeight = oSDO_ProfilePart.Height
            
        Else
            ' Bounding Object is not valid type
            dBoundingHeight = 0#
        End If
    
        GetBoundingStiffenerHeight = dBoundingHeight
        Exit Function
    End If
        
    ' Bounding Stiffener is an Edge Reinforcement
    ' Check Edge Reinforcement orientation is "OnCentered"
    Set oSDO_EdgeReinforcement = New StructDetailObjects.EdgeReinforcement
    Set oSDO_EdgeReinforcement.object = oBoundingStiffener
    
    If LCase(Trim(oSDO_EdgeReinforcement.ReinforcementPosition)) = LCase("OnFace") Then
        ' Edge Reinforcement orientation is "OnFace"
        ' use the Edge Reinforcement's Height
        Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
        Set oSDO_ProfilePart.object = oBoundingStiffener
        dBoundingHeight = oSDO_ProfilePart.Height
        GetBoundingStiffenerHeight = dBoundingHeight
        Exit Function
    End If
     
    ' Edge Reinforcement orientation is NOT "OnFace"
    ' Calculate Edge Reinforcement TopSetBack and BottomSetBack distances
    ' Use the given Bounded Port as the Poistion to calculate the distances
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
    oTopologyLocate.FindApproxCenterAndNormal oBoundedPortGeometry, _
                                              oEndPosition, oNormal

    oSDO_EdgeReinforcement.GetPlateToTopBottom oEndPosition, _
                                               dTopSetBack, dBottomSetBack
    
    ' Check if returned distance values are equal
    dBoundingHeight = dTopSetBack - dBottomSetBack
    If Abs(dBoundingHeight) < 0.001 Then
        GetBoundingStiffenerHeight = dTopSetBack
        Exit Function
    End If
    
    ' For Performance reasons
    ' Assume Edge Reinforcement orientation is "On Edge - Centered"
    ' And Ignore the the Stifened Plate thickness
    ' Return the smallest value
    If dTopSetBack < dBottomSetBack Then
        dBoundingHeight = dTopSetBack
    Else
        dBoundingHeight = dBottomSetBack
    End If
    
'$$$$zMsgBox "GetBoundingStiffenerHeight" & vbCrLf & _
        "oEdgeReinforcement.ReinforcementPosition :" & _
        oSDO_EdgeReinforcement.ReinforcementPosition & vbCrLf & _
        "dTopSetBack =" & Format(dTopSetBack, "0.####") & vbCrLf & _
        "dBottomSetBack =" & Format(dBottomSetBack, "0.####") & vbCrLf & _
        "GetBoundingStiffenerHeight =" & CStr(dBoundingHeight)
            
    ' Have the distance
    '   from the Edge Reinforcement Top to the Reinforced Plate Base/Offset
    '   from the Edge Reinforcement Bottom to the Reinforced Plate Base/Offset
    ' Determine which distance should be used as the Bounding Height
    ' Use the Bounded Profile's Mounting Face Port
    '   retreive the Normal Vector, this is pointing away from the Profile
    ' Use the Edge Reinforcement Top Face Port
    '   retreive the Normal Vector, this is pointing away from the Profile
    
    Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
    Set oSDO_ProfilePart.object = oBoundedProfile
    Set oPort = oSDO_ProfilePart.MountingFacePort
    oTopologyLocate.GetProjectedPointOnModelBody oPort.Geometry, oEndPosition, _
                                                 oPoint_MountingFace, _
                                                 oNormal_MountingFace
    Set oPort = Nothing
    Set oSDO_ProfilePart = Nothing
    
    Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
    Set oSDO_ProfilePart.object = oBoundingStiffener
    Set oPort = oSDO_ProfilePart.SubPort(JXSEC_TOP)
    oTopologyLocate.GetProjectedPointOnModelBody oPort.Geometry, oEndPosition, _
                                                 oPoint_TopFace, _
                                                 oNormal_TopFace
                                                 
    ' If the two normals are in opposite directions (Dot Product < 0.0)
    ' Then the Bounded Profile is Height is in same direction as ER Top
    ' return the distance from the Edge Reinforcement Top to the Reinforced Plate
    ' else the distance from the Edge Reinforcement Bottom to the Reinforced Plate
    dDot = oNormal_MountingFace.Dot(oNormal_TopFace)
    If dDot < 0# Then
        dBoundingHeight = dTopSetBack
    Else
        dBoundingHeight = dBottomSetBack
    End If
    
'$$$$zMsgBox "GetBoundingStiffenerHeight" & vbCrLf & _
        "oEdgeReinforcement.ReinforcementPosition :" & _
        oSDO_EdgeReinforcement.ReinforcementPosition & vbCrLf & _
        "dTopSetBack =" & Format(dTopSetBack, "0.####") & vbCrLf & _
        "dBottomSetBack =" & Format(dBottomSetBack, "0.####") & vbCrLf & _
        "MoundingFaceNormal:" & Format(oNormal_MountingFace.x, "0.####") & " , " & _
        Format(oNormal_MountingFace.y, "0.####") & " , " & _
        Format(oNormal_MountingFace.z, "0.####") & " , " & vbCrLf & _
        "oNormal_TopFace:" & Format(oNormal_TopFace.x, "0.####") & " , " & _
        Format(oNormal_TopFace.y, "0.####") & " , " & _
        Format(oNormal_TopFace.z, "0.####") & " , " & vbCrLf & _
        "dDot:" & Format(dDot, "0.####") & vbCrLf & _
        "GetBoundingStiffenerHeight =" & Str(dBoundingHeight)
    
    GetBoundingStiffenerHeight = dBoundingHeight
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetBoundingStiffenerHeight").Number
End Function

' ********************************************************************************
' Method: GetParentAnswer
'
' Description:
'   Gets the Question from from the Smart Occurenence Parent
'
' Inputs:
'
' Outputs:
' ********************************************************************************
Public Sub GetParentAnswer(oSmartObject As Object, sQuestion As String, _
                           sAnswer As String)
 On Error GoTo ErrorHandler
     
    Dim vAnswer As Variant
    
    Dim oParameterLogic As IJDParameterLogic
    Dim oMemberDescription As IJDMemberDescription
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartOccurrence As DEFINITIONHELPERSINTFLib.IJSmartOccurrence
    
    Dim oParentSmartClass As IJSmartClass
    Dim oParentSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    On Error GoTo ErrorHandler
    sAnswer = ""
    
    If TypeOf oSmartObject Is IJDMemberDescription Then
        Set oMemberDescription = oSmartObject
        Set oSmartOccurrence = oMemberDescription.CAO
        Set oSmartItem = oSmartOccurrence.ItemObject
        
    ElseIf TypeOf oSmartObject Is IJDParameterLogic Then
        Set oParameterLogic = oSmartObject
        Set oSmartItem = oParameterLogic.SmartItem
        Set oSmartOccurrence = oParameterLogic.SmartOccurrence
    
    Else
        Exit Sub
    End If
    
    Set oParentSmartClass = oSmartItem.Parent
    Set oParentSymbolDefinition = oParentSmartClass.SelectionRuleDef

    Set oCommonHelper = New DefinitionHlprs.CommonHelper
    vAnswer = oCommonHelper.GetAnswer(oSmartOccurrence, oParentSymbolDefinition, _
                                      sQuestion)
    sAnswer = vAnswer
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetParentAnswer").Number
End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub zvMsgBox(vText As Variant, _
                   Optional sDumArg1 As String, Optional sTitle As String)
On Error Resume Next
Dim sText As String
    sText = vText
    zMsgBox sText, sDumArg1, sTitle

End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub zMsgBox(sText As String, _
                   Optional sDumArg1 As String, Optional sTitle As String)
On Error Resume Next

Dim iFileNumber
Dim sFileName As String

Exit Sub
'$$$Debug $$$    MsgBox sText, , sTitle

    iFileNumber = FreeFile
    sFileName = "C:\Temp\TraceFile.txt"
    Open sFileName For Append Shared As #iFileNumber
    
    If Len(Trim(sTitle)) > 0 Then
        Write #iFileNumber, sTitle
    End If
    
    Write #iFileNumber, sText
    Close #iFileNumber
End Sub

' ********************************************************************************
' Method:
'   AreLandingCurvesIntersecting
' Description:
'   Checks the landing curves of input profile parts are intersecting or not
' ********************************************************************************
Public Function AreLandingCurvesIntersecting(oProfilePart1 As Object, oProfilePart2 As Object) As Boolean
On Error GoTo ErrorHandler

    'Initialize the boolean for the majority scenario
    AreLandingCurvesIntersecting = True
    
    'The profiles can be stiffeners, ERs, and Beams
    
    'If the profile parts are stiffenening different plates,
    'we do not need to check landing curve interesection geometrically

    If Not StiffeningSamePlate(oProfilePart1, oProfilePart2) Then
        Dim oLandingCrv1 As IJWireBody
        Dim oLandingCrv2 As IJWireBody
        
        Set oLandingCrv1 = GetProfilePartLandingCurve(oProfilePart1)
        Set oLandingCrv2 = GetProfilePartLandingCurve(oProfilePart2)
        
        If (Not oLandingCrv1 Is Nothing) And (Not oLandingCrv2 Is Nothing) Then
            'Check to see if the intersection between these landing curves exists or not.
            Dim oIntersectUtil As IIntersect
            Set oIntersectUtil = New Intersect
            
            Dim oCommonBody As IUnknown
            
            On Error Resume Next 'If the landing curves are not intersecting, we expect error.
            oIntersectUtil.GetCommonGeometry oLandingCrv1, oLandingCrv2, oCommonBody, 0
            If Err.Number <> 0 Then Err.Clear
            
            AreLandingCurvesIntersecting = (Not oCommonBody Is Nothing)
            
            Set oIntersectUtil = Nothing
            Set oCommonBody = Nothing
        End If
        
        'Clean up
        Set oLandingCrv1 = Nothing
        Set oLandingCrv2 = Nothing
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "AreLandingCurvesIntersecting").Number
End Function

' ********************************************************************************
' Method:
'   LandingCurve
' Description:
'   Gets the landingcurve for the given Stiffener/ER/Beam
' ********************************************************************************
Private Function GetProfilePartLandingCurve(oProfilePart As Object) As IJWireBody
On Error GoTo ErrorHandler
    
    Dim oStructDetailHelper As StructDetailHelper
    Dim oTopoLocate As IJTopologyLocate
    Dim oLandingCrv As IJWireBody
    
    Set oStructDetailHelper = New StructDetailHelper
    Set oTopoLocate = New TopologyLocate
    
    'Based on whether the part is derived from system or not,
    'we shall get the landing curve differently for performance
    
    Dim oIJStructGraph As IJStructGraph
    Set oIJStructGraph = oProfilePart
    
    If Not oIJStructGraph Is Nothing Then
        Dim oParentSystem As IJSystem
        oStructDetailHelper.IsPartDerivedFromSystem oIJStructGraph, oParentSystem
        
        'Derived from system?
        If Not oParentSystem Is Nothing Then
            Set oLandingCrv = oTopoLocate.GetProfileParentWireBody(oProfilePart)
        Else
            'Not derived from system
            Dim oPartSupport As GSCADSDPartSupport.IJPartSupport
            Dim oProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport
            
            Set oProfilePartSupport = New GSCADSDPartSupport.ProfilePartSupport
            Set oPartSupport = oProfilePartSupport
            Set oPartSupport.Part = oProfilePart
            
            ' Get the curve (default direction is SideUnspecified, which gives
            ' a curve through the load point)
            Dim oThicknessDir As IJDVector
            Dim bThicknessCentered As Boolean
            
            oProfilePartSupport.GetProfilePartLandingCurve oLandingCrv, _
                                                           oThicknessDir, _
                                                           bThicknessCentered
                        
            Set oProfilePartSupport = Nothing
            Set oPartSupport = Nothing
            Set oThicknessDir = Nothing
        End If
        
        Set oParentSystem = Nothing
    End If
    
    Set GetProfilePartLandingCurve = oLandingCrv
    
    Set oLandingCrv = Nothing
    Set oStructDetailHelper = Nothing
    Set oTopoLocate = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetProfilePartLandingCurve").Number
End Function

' ********************************************************************************
' Method:
'   StiffeningSamePlate
' Description:
'   Checks if the input profile parts stiffening the same plate.
' ********************************************************************************
Private Function StiffeningSamePlate(oProfilePart1 As Object, oProfilePart2 As Object) As Boolean
On Error GoTo ErrorHandler

    Dim bStiffeningSamePlate As Boolean
    bStiffeningSamePlate = False
    
    'If any of the input part is beam, we have the answer straight away
    If Not TypeOf oProfilePart1 Is IJBeam And _
       Not TypeOf oProfilePart2 Is IJBeam Then
        
        'None of the input part is beam.
        'We will get the stiffened plate, and compare.
        
        Dim oSDO_ProfilePart1 As New StructDetailObjects.ProfilePart
        Dim oSDO_ProfilePart2 As New StructDetailObjects.ProfilePart
        
        Set oSDO_ProfilePart1.object = oProfilePart1
        Set oSDO_ProfilePart2.object = oProfilePart2
        
        Dim oStiffenedPlate1 As Object
        Dim oStiffenedPlate2 As Object
        Dim bSystem As Boolean
        
        oSDO_ProfilePart1.GetStiffenedPlate oStiffenedPlate1, bSystem
        oSDO_ProfilePart2.GetStiffenedPlate oStiffenedPlate2, bSystem
        
        bStiffeningSamePlate = (oStiffenedPlate1 Is oStiffenedPlate2)
        
        Set oSDO_ProfilePart1 = Nothing
        Set oSDO_ProfilePart2 = Nothing
        Set oStiffenedPlate1 = Nothing
        Set oStiffenedPlate2 = Nothing
    End If
    
    StiffeningSamePlate = bStiffeningSamePlate
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "StiffeningSamePlate").Number
End Function

'
' This method estimates alligator web cut parameters
'
' Parameters are first estimated for profile on planar surface
' If the profile is on curved surface, parameters will be adjusted based on:
'    Angle between tangent at seam location and tangent at offset location
'    Whether cut is at start or end point of the profile
'
'    The slope angle is adjusted to ensure alligator cut will cut flange,
'    but not too much(too much means the flange is cut through)
'
'            oVDirAtOffset              oVDirAtOffset
'            ^                              ^
'            |                              |
'            |                              |
'            |                              |
'    *       |                     *        |
'      *     |                       *      |
'   * * * ------> oTangentAtCut   * * * <------ oTangentAtCut
'      * Start                       * End
'    *   Need adjust slope angle   *  Need adjust slope angle
'
'    If cut is at start point
'        If angle between oTangentAtCut and oVDirAtOffset is greater than 90 degrees,
'           Slope need to be adjusted based on the curvity of landing curve
'           LengthV need to be calculated based on adjusted slope
'
'        If angle between oTangentAtCut and oVDirAtOffset is less than 90 degrees,
'           Only LengthV need to be calculated based the curvity of landing curve
'
'    If cut is at end point
'        If angle between oTangentAtCut and oVDirAtOffset is less than 90 degrees,
'           Slope need to be adjusted based on the curvity of landing curve
'           LengthV need to be calculated based on adjusted slope

'        If angle between oTangentAtCut and oVDirAtOffset is greater than 90 degrees,
'           Only LengthV need to be calculated based the curvity of landing curve
'
' Note: All the results are estimated,you can always find some way to improve it!

Public Sub EstimateAlligatorWebCutParameters( _
               ByVal oWebCut As Object, _
               ByRef dLengthU As Double, _
               ByRef dLengthV As Double, _
               ByRef dSlope As Double, _
               Optional ByRef bIsLandingCurveLinear As Boolean = True)
On Error GoTo ErrorHandler

   Dim oWebCutWrapper As New StructDetailObjects.WebCut
    
   Set oWebCutWrapper.object = oWebCut
   
   Dim oProfilePart As StructDetailObjects.ProfilePart
   Dim dBoundedHeight As Double
   Dim dBoundingHeight As Double
   Dim dBoundedFT As Double
   Dim dBoundedWL As Double
   Dim oBoundedSystem As Object
   
   Set oProfilePart = New StructDetailObjects.ProfilePart
   Set oProfilePart.object = oWebCutWrapper.Bounded
   
   dBoundedHeight = Round(oProfilePart.Height, 3)
   dBoundedFT = Round(oProfilePart.flangeThickness, 3)
   dBoundedWL = Round(oProfilePart.WebLength, 3)
   Set oBoundedSystem = oProfilePart.ParentSystem
   
   Set oProfilePart.object = oWebCutWrapper.Bounding
   dBoundingHeight = Round(oProfilePart.Height, 3)
   
   Set oProfilePart = Nothing
   
   Dim dAngle As Double
   Dim dAdjustedSlope As Double
   Dim dSlopeChanged As Double
   Dim dTemp As Double
   
   ' Default alligator cut angle to 30 degrees
   dAngle = (3.1415926 * 30) / 180
   
   ' If the height difference is too much,adjust it to avoid failure of cut
   dTemp = Atn(Sqr(dBoundingHeight * dBoundingHeight / dBoundedFT / dBoundedFT - 1))
   If dTemp < dAngle Then
       dAngle = dTemp * 0.9
   End If
   
   ' Default cut slope near flange to 0.01 degrees
   dSlope = (3.1415916 * 0.01) / 180
      
   Dim oWebCutLocation As IJDPosition
   
   Set oWebCutLocation = oWebCutWrapper.BoundedLocation
   Set oWebCutWrapper = Nothing
   
   Dim oWireBodyUtil As New SGOWireBodyUtilities
   Dim oStartPoint As IJDPosition
   Dim oEndPoint As IJDPosition
   Dim oStartDir As IJDVector
   Dim oEndDir As IJDVector
   
   oWireBodyUtil.GetEndPointsAndTangents oBoundedSystem, oStartPoint, oStartDir, oEndPoint, oEndDir
   
   Dim oProjectedPoint As IJDPosition
   Dim oTangentAtCut As IJDVector
   
   oWireBodyUtil.GetClosestPointOnWire oBoundedSystem, oWebCutLocation, oProjectedPoint, oTangentAtCut
   
   Dim dDistToStart As Double
   Dim dDistToEnd As Double
   
   dDistToStart = oStartPoint.DistPt(oProjectedPoint)
   dDistToEnd = oEndPoint.DistPt(oProjectedPoint)
   Set oProjectedPoint = Nothing
   
   Dim oInwardOffsetDir As IJDVector
   Dim oOutwardOffsetDir As IJDVector
   Dim bCutAtStart As Boolean
   
   bCutAtStart = True
   Set oInwardOffsetDir = oTangentAtCut.Clone
   Set oOutwardOffsetDir = oTangentAtCut.Clone
   
   If dDistToStart > dDistToEnd Then
       ' Profile direction should point from start to end
       ' If at end, reverse direction to get offset point on landing curve
       bCutAtStart = False
       oInwardOffsetDir.Set -oTangentAtCut.x, -oTangentAtCut.y, -oTangentAtCut.z
   Else
       oOutwardOffsetDir.Set -oTangentAtCut.x, -oTangentAtCut.y, -oTangentAtCut.z
   End If

   '
   ' First estimate LengthV based on profile on planar surface:
   '
   ' K CosA = FT2 + (L - KSinA)SinB
   ' WL2 -PH1 + K = V
   '
   ' KCosA = FT2 + LSinB - KSinASinB
   ' WL2 -PH1 + K = V
   '
   ' K(CosA + SinASinB) = FT2 + LSinB)
   ' K = V - (WL2 - PH1)
   '
   ' (V - (WL2 - PH1))(CosA + SinASinB) = FT2 + VSinB/CosA
   '
   ' V(CosA+ SinASinB) - (WL2 - PH1)(CosA + SinASinB) = FT2 + VSinB/CosA
   '
   ' V(CosA + SinASinB - SinB/CosA) = FT2+ (WL2 - PH1)(CosA + SinASinB)
   '
   ' V = (FT2 + (WL2 - PH1)(CosA + SinASinB))/(CosA + SinASinB - SinB/CosA)
   dLengthV = (dBoundedFT + (dBoundedWL - dBoundingHeight) * (Cos(dAngle) + Sin(dAngle) * Sin(dSlope))) / (Cos(dAngle) + Sin(dAngle) * Sin(dSlope) - Sin(dSlope) / Sin(dAngle))
   dLengthU = dLengthV / Tan(dAngle)
  
   ' Adjust slope value based on the curvity of landing curve
   dAdjustedSlope = dSlope

   Dim oProfileAttributes As IJProfileAttributes
   Dim oMatrixAtCut As IJDT4x4
   
   Set oProfileAttributes = New ProfileUtils
   oProfileAttributes.GetProfileOrientationMatrix oBoundedSystem, oWebCutLocation, oMatrixAtCut
   
   Dim oUDirAtCut As IJDVector
   Dim oVDirAtCut As IJDVector
   
   Set oUDirAtCut = New DVector
   oUDirAtCut.x = oMatrixAtCut.IndexValue(0)
   oUDirAtCut.y = oMatrixAtCut.IndexValue(1)
   oUDirAtCut.z = oMatrixAtCut.IndexValue(2)
   
   Set oVDirAtCut = New DVector
   oVDirAtCut.x = oMatrixAtCut.IndexValue(4)
   oVDirAtCut.y = oMatrixAtCut.IndexValue(5)
   oVDirAtCut.z = oMatrixAtCut.IndexValue(6)
      
   Dim oTangentTemp As IJDVector
   Dim oInwardOffsetPoint As IJDPosition
   Dim oOutwardOffsetPoint As IJDPosition
   
   Set oTangentTemp = oInwardOffsetDir.Clone()
   oTangentTemp.x = oTangentTemp.x * dLengthU
   oTangentTemp.y = oTangentTemp.y * dLengthU
   oTangentTemp.z = oTangentTemp.z * dLengthU
   Set oInwardOffsetPoint = oWebCutLocation.Offset(oTangentTemp)
   
   Dim oTangentAtInwardOffsetPoint As IJDVector
   Dim oTangentAtOutwardOffsetPoint As IJDVector
   Dim oInwardProjectedPoint As IJDPosition
   
   oWireBodyUtil.GetClosestPointOnWire oBoundedSystem, oInwardOffsetPoint, oInwardProjectedPoint, oTangentAtInwardOffsetPoint
   Set oTangentAtOutwardOffsetPoint = oTangentAtCut.Clone()
   
   Dim dDot As Double
   
   dDot = oVDirAtCut.Dot(oTangentAtInwardOffsetPoint)
   
   If Abs(dDot) > 0 Then
       bIsLandingCurveLinear = False
       
       ' Tangent at offset is different tangent at cut,need to adjust slope and LengthV
       Dim dDotToGetAngle As Double
       Dim oTangentAtInwardOffsetCrossU As IJDVector
       Dim oTangentAtOutwardOffsetCrossU As IJDVector
       
       Set oTangentAtOutwardOffsetCrossU = oTangentAtOutwardOffsetPoint.Cross(oUDirAtCut)
       Set oTangentAtInwardOffsetCrossU = oTangentAtInwardOffsetPoint.Cross(oUDirAtCut)
       dDotToGetAngle = oTangentAtOutwardOffsetCrossU.Dot(oTangentAtInwardOffsetCrossU)
       
       Set oTangentAtOutwardOffsetCrossU = Nothing
       Set oTangentAtInwardOffsetCrossU = Nothing
       
       If dDotToGetAngle < -1 Then
          dDotToGetAngle = -1
       ElseIf dDotToGetAngle > 1 Then
          dDotToGetAngle = 1
       End If
       
       If Abs(dDotToGetAngle) < 0.000001 Then
          dSlopeChanged = 3.14159265 / 2
       Else
          dSlopeChanged = Atn(Sqr(1 / (dDotToGetAngle * dDotToGetAngle) - 1))
       End If
   
       If dDotToGetAngle < 0 Then
          dSlopeChanged = 3.14159265 - dSlopeChanged
       End If
       
       dSlopeChanged = GetAngleBetweenVectors(oTangentAtOutwardOffsetPoint, _
                                              oTangentAtInwardOffsetPoint, _
                                              oUDirAtCut)
       dAdjustedSlope = dSlopeChanged / 2
       
       ' Check angle between tangents at inward offset point and outward offset point
       ' to decide what need to be adjusted -- LengthV or LengthV and Slope
       Dim bAdjustVOnly As Boolean
       Dim dDotOfTangentAtOutOffsetAndVDirAtInOffset As Double
       
       Dim oMatrixAtInOffset As IJDT4x4
       Dim oVDirAtInOffset As New DVector
       
       oProfileAttributes.GetProfileOrientationMatrix oBoundedSystem, oInwardProjectedPoint, oMatrixAtInOffset
       oVDirAtInOffset.x = oMatrixAtInOffset.IndexValue(4)
       oVDirAtInOffset.y = oMatrixAtInOffset.IndexValue(5)
       oVDirAtInOffset.z = oMatrixAtInOffset.IndexValue(6)
       
       bAdjustVOnly = False
       dDotOfTangentAtOutOffsetAndVDirAtInOffset = oTangentAtOutwardOffsetPoint.Dot(oVDirAtInOffset)

       If bCutAtStart = True Then
           If dDotOfTangentAtOutOffsetAndVDirAtInOffset > 0 Then
               bAdjustVOnly = False
           Else
               bAdjustVOnly = True
           End If
       Else
           If dDotOfTangentAtOutOffsetAndVDirAtInOffset > 0 Then
               bAdjustVOnly = True
           Else
               bAdjustVOnly = False
           End If
       End If
       
'       MsgBox "bCutAtStart = " & bCutAtStart & vbNewLine & _
'              " Start: ( " & oStartPoint.x & " , " & oStartPoint.y & " , " & oStartPoint.z & " ) " & vbNewLine & _
'              " End:   ( " & oEndPoint.x & " , " & oEndPoint.y & " , " & oEndPoint.z & " ) " & vbNewLine & _
'              " TangentAtCut: ( " & oTangentAtCut.x & " , " & oTangentAtCut.y & " , " & oTangentAtCut.z & " )" & vbNewLine & _
'              " TangAtInOffset: ( " & oTangentAtInwardOffsetPoint.x & " , " & oTangentAtInwardOffsetPoint.y & " , " & oTangentAtInwardOffsetPoint.z & " )" & vbNewLine & _
'              " VDirAtInOffset: ( " & oVDirAtInOffset.x & " , " & oVDirAtInOffset.y & " , " & oVDirAtInOffset.z & " )" & vbNewLine & _
'              " dDotOfTangentAtOutOffsetAndVDirAtInOffset = " & dDotOfTangentAtOutOffsetAndVDirAtInOffset & vbNewLine & _
'              " bAdjustVOnly = " & bAdjustVOnly & vbNewLine & _
'              " dLengthV = " & dLengthV & vbNewLine & _
'              " dLengthU = " & dLengthU & vbNewLine & _
'              " dBoundedHeight = " & (dBoundedFT + dBoundedWL) & vbNewLine & _
'              " dBoundingHeight = " & dBoundingHeight & vbNewLine & _
'              " dDeltaHeight = " & (dBoundedFT + dBoundedWL - dBoundingHeight) & vbNewLine & _
'              " dAdjustedSlope = " & (dAdjustedSlope * 180) / 3.14159265
              
       If bAdjustVOnly = True Then
           '  Only adjust LengthV
           '           *
           '        *
           '  *
           '  *        *
           '  *     *
           '  *
           
           ' KCosA = FT2 + (L-KSinA)2Sin(B/2)
           ' K + WL2 - PH1 = V/CosB + (U - VTanB)Sin(B/2)
           '
           ' KCosA = FT2 + (V/SinA-KSinA)2Sin(B/2)
           ' K + WL2 - PH1 = V/CosB + (V/TanA - VTanB)Sin(B/2)
           '
           ' KCosA = FT2 + 2VSin(B/2)/SinA - 2KSinASin(B/2)
           ' K + WL2 - PH1 = V / CosB + VSin(B / 2) / TanA - VTanBSin(B / 2)
           '
           ' K(CosA + 2SinASin(B/2)) = FT2 + 2VSin(B/2)/SinA
           ' K = V[1/CosB + Sin(B/2)/TanA - TanBSin(B/2)] - (WL2 - PH1)
           '
           ' (CosA + 2SinASin(B/2))V[1/CosB + Sin(B/2)/TanA - TanBSin(B/2)] - [CosA+2SinASin(B/2)](WL2 - PH1) = FT2 + 2VSin(B/2)/SinA
           '
           ' V{[CosA + 2SinASin(B/2)][1/CosB + Sin(B/2)/TanA - TanBSin(B/2)] - 2Sin(B2)/SinA)} =
           ' FT2 + [CosA+2SinASin(B/2)](WL2 - PH1)
           '
           ' V = {FT2 + [CosA+2SinASin(B/2)](WL2-PH1)}/
           '        {[CosA + 2SinASin(B/2)][1/CosB + Sin(B/2)/TanA - TanBSin(B/2)]-2Sin(B2)/SinA}
           
           ' V = {FT2 + (WL2 - PH1)[CosA + 2SinASin(B/2)]}/
           '       {[1/CosB + (1/TanA-TanB)Sin(B/2)][CosA + 2SinASin(B/2)] - 2Sin(B2)SinA}
           '
           dLengthV = (dBoundedFT + (dBoundedWL - dBoundingHeight) * (Cos(dAngle) + 2 * Sin(dAngle) * Sin(dAdjustedSlope / 2))) / _
                      ((1 / Cos(dAdjustedSlope) + (1 / Tan(dAngle) - Tan(dAdjustedSlope)) * Sin(dAdjustedSlope / 2)) * (Cos(dAngle) + 2 * Sin(dAngle) * Sin(dAdjustedSlope / 2)) - 2 * Sin(dAdjustedSlope / 2) / Sin(dAngle))
           dLengthU = dLengthV / Tan(dAngle)
       Else
           ' Adjust slope and dLength
           
           '  *
           '  *     *
           '  *        *
           '  *
           '        *
           '           *
           ' KCos(A - Slope) = FT2
           ' WL2 -PH1 + K = V - USin(Slope)
           ' WL2 - PH1 + K = V -Sin(Slope)V/TanA
           '
           ' WL2 -PH1 + FT2 / Cos(A - Slope) = V(1 - Sin(Slope) / TanA)
           '
           ' V = (WL2 - PH1 + FT2 / Cos(A - Slope)) / (1 - Sin(Slope) / TanA)
           
           dSlope = dSlope + dAdjustedSlope
           dLengthV = (dBoundedWL - dBoundingHeight + dBoundedFT / Cos(dAngle - dSlope)) / (1 - Sin(dSlope) / Tan(dAngle))
           dLengthU = dLengthV / Tan(dAngle)
        End If
   End If
   
   Set oWebCutLocation = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "EstimateAlligatorWebCutParameters").Number
End Sub
Public Function GetAngleBetweenVectors(oVector1 As IJDVector, _
                                       oVector2 As IJDVector, _
                                       oNormal As IJDVector) As Double
   GetAngleBetweenVectors = 0
   
   If oVector1 Is Nothing Or oVector2 Is Nothing Or oNormal Is Nothing Then
      Exit Function
   End If
   
   Dim oCrossOfVector1AndNormal As IJDVector
   Dim oCrossOfVector2AndNormal As IJDVector
   Dim dDotToGetAngle As Double
   
   Set oCrossOfVector1AndNormal = oVector1.Cross(oNormal)
   Set oCrossOfVector2AndNormal = oVector2.Cross(oNormal)

   dDotToGetAngle = oCrossOfVector1AndNormal.Dot(oCrossOfVector2AndNormal)
   Set oCrossOfVector1AndNormal = Nothing
   Set oCrossOfVector2AndNormal = Nothing

   If dDotToGetAngle < -1 Then
      dDotToGetAngle = -1
   ElseIf dDotToGetAngle > 1 Then
      dDotToGetAngle = 1
   End If
   
   Dim dAngle As Double

   If Abs(dDotToGetAngle) < 0.000001 Then
      dAngle = 3.14159265 / 2
   Else
      dAngle = Atn(Sqr(1 / (dDotToGetAngle * dDotToGetAngle) - 1))
   End If
   
   If dDotToGetAngle < 0 Then
      dAngle = 3.14159265 - dAngle
   End If
   
   GetAngleBetweenVectors = dAngle
   
End Function



'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'   Given an MemberDescription
'   Checks that the SmartItem is a WebCut or a FlangeCut
'   Checks that the Bounded Part is a Profile Part
'   Checks if the Bounding Part is a Plate part
'   Checks if the Bounding Port is a Lateral Edge
'
Public Function IsProfileBoundedByPlateEdge(ByVal pMemberDescription As IJDMemberDescription) As Boolean
On Error GoTo ErrorHandler
                                    
    Dim lCtxId As Long
    Dim sError As String
    
    Dim oObject As Object
    Dim oBoundedPart As Object
    Dim oBoundedPort As Object
    Dim oBoundingPart As Object
    Dim oBoundingPort As Object
    
    Dim oStructPort As IJStructPort
    Dim oStructFeature As IJStructFeature
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    
    IsProfileBoundedByPlateEdge = False
    
    ' Verify given SmartItem is a Web Cut or Flange Cut Feature
    Set oObject = pMemberDescription.CAO
    If TypeOf oObject Is IJStructFeature Then
        Set oStructFeature = oObject
        If oStructFeature.get_StructFeatureType = SF_WebCut Then
            sError = "Retreiving Bounded/Bounding objects from WebCut"
            Set oSDO_WebCut = New StructDetailObjects.WebCut
            Set oSDO_WebCut.object = oObject
            
            Set oBoundedPart = oSDO_WebCut.Bounded
            Set oBoundedPort = oSDO_WebCut.BoundedPort
            
            Set oBoundingPart = oSDO_WebCut.Bounding
            Set oBoundingPort = oSDO_WebCut.BoundingPort
            
        ElseIf oStructFeature.get_StructFeatureType = SF_FlangeCut Then
            sError = "Retreiving Bounded/Bounding objects from FlangeCut"
            Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
            Set oSDO_FlangeCut.object = oObject
            
            Set oBoundedPart = oSDO_FlangeCut.Bounded
            Set oBoundedPort = oSDO_FlangeCut.BoundedPort
            
            Set oBoundingPart = oSDO_FlangeCut.Bounding
            Set oBoundingPort = oSDO_FlangeCut.BoundingPort
        Else
            sError = "Given IJStructFeature is NOT valid EndCut SF_WebCut or SF_FlangeCut"
            Exit Function
        End If
    Else
        sError = "Given Obect is NOT IJStructFeature"
        Exit Function
    End If
    
    ' Verify Bounded Part is a Profile part
    If oBoundedPart Is Nothing Then
        sError = "EndCut BoundedPart is Nothing"
        Exit Function
    ElseIf Not TypeOf oBoundedPart Is IJProfilePart Then
        sError = "EndCut BoundedPart is NOT IJProfilePart"
        Exit Function
    End If
    
    ' Verify Bounding Part is a Plate part
    If oBoundingPart Is Nothing Then
        sError = "EndCut BoundingPart is Nothing"
        Exit Function
    ElseIf Not TypeOf oBoundingPart Is IJPlate Then
        sError = "EndCut BoundingPart is NOT IJPlate"
        Exit Function
    End If
    
    ' Verify Bounding Port is a Lateral Face (edge) Port
    If oBoundingPort Is Nothing Then
        sError = "EndCut BoundingPort is Nothing"
        Exit Function
    ElseIf Not TypeOf oBoundingPort Is IJStructPort Then
        sError = "EndCut BoundingPort is NOT IJStructPort"
        Exit Function
    Else
        sError = "Retrieving EndCut BoundingPort Context Id"
        Set oStructPort = oBoundingPort
        lCtxId = oStructPort.ContextID
        If lCtxId And CTX_LATERAL Then
            IsProfileBoundedByPlateEdge = True
        End If

    End If
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IsProfileBoundedByPlateEdge", sError).Number
End Function

Public Function IsBoundingStiffPenetratingSupport( _
                         ByRef oEndCut As DEFINITIONHELPERSINTFLib.IJSmartOccurrence) As Boolean
    
    IsBoundingStiffPenetratingSupport = False
    
    ' ----------------------
    ' Get the end cut object
    ' ----------------------
    Dim oSDO_WC As StructDetailObjects.WebCut
    Set oSDO_WC = New StructDetailObjects.WebCut
    Set oSDO_WC.object = oEndCut
    
    ' -----------------------
    ' Get the stiffened plate
    ' -----------------------
    ' If not a stiffener, there is no stiffened plate for the bounding object to penetrate
    ' Return "False"
    Dim oStiffener As IJStiffener
    If TypeOf oSDO_WC.Bounded Is IJStiffener Then
        Set oStiffener = oSDO_WC.Bounded
    Else
        Exit Function
    End If
    
    Dim oStiffenedPlate As IJPlate
    Set oStiffenedPlate = oStiffener.PlateSystem
    
    ' -------------------------------
    ' If bounding object is a profile
    ' -------------------------------
    ' ToDo: What about members?
    If TypeOf oSDO_WC.Bounding Is IJProfile Then
        ' -------------------
        ' Get the root system
        ' -------------------
        Dim oBoundingSystem As IJConnectable
        Dim oStructDetailHelper As IJStructDetailHelper
        Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
        
        If TypeOf oSDO_WC.Bounding Is IJStructGraph Then
            oStructDetailHelper.IsPartDerivedFromSystem oSDO_WC.Bounding, oBoundingSystem, True
        End If
                
        ' ---------------------------------------------------
        ' If bounding object is connected to stiffened object
        ' ---------------------------------------------------
        Dim oConnections As IJElements
        Dim isConnected As Boolean
        
        oBoundingSystem.isConnectedTo oStiffenedPlate, isConnected, oConnections
        
        If isConnected Then
            ' --------------------------------------------
            ' If the stiffened plate is not a boundary and
            ' If they meet at a single point
            ' --> set result to true
            ' ----------------------------------------
            ' First thought was to use the ConnectionBehavior, but the
            ' root connection behavior is "Standard" while the detailed connection
            ' behavior is "Penetration".  We can't rely on detailing between these
            ' two other parts
            Dim oBoundaries As Collection
            Dim oBoundary As Object
            Dim oProfileAttr As IJProfileAttributes
            Set oProfileAttr = New ProfileUtils
            
            Set oBoundaries = oProfileAttr.GetProfileSystemBoundary(oBoundingSystem)
            
            For Each oBoundary In oBoundaries
                If oBoundary Is oStiffenedPlate Then
                    Exit Function
                End If
            Next oBoundary
            
            On Error Resume Next
            Dim oIntPoint As IJDPosition
            Set oIntPoint = oProfileAttr.GetProfileIntersectionPoint(oBoundingSystem, oStiffenedPlate)
            
            If Not oIntPoint Is Nothing Then
                IsBoundingStiffPenetratingSupport = True
            End If
        End If
    End If
    
    Exit Function
End Function

Public Function IsLapped(ByRef oBoundedPort As IJPort, Optional ByRef dLapDist As Double) As Boolean

    IsLapped = False
    
    ' ------------------------------------
    ' Get the bounded object
    ' ------------------------------------
    Dim oBoundedConn As Object
    Set oBoundedConn = oBoundedPort.Connectable
    
    ' -----------------------------
    ' Get the root geometry
    ' -----------------------------
    Dim oParentSystem As Object
    Dim oStructDetailHelper As IJStructDetailHelper
    Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
    
    If TypeOf oBoundedConn Is IJStructGraph Then
        oStructDetailHelper.IsPartDerivedFromSystem oBoundedConn, oParentSystem, True
    End If
        
    If oParentSystem Is Nothing Then
        oParentSystem = oBoundedConn
    End If
    
    Dim oRootLC As Object
    Dim oProfileAttr As IJProfileAttributes
    Set oProfileAttr = New ProfileUtils
    oProfileAttr.GetRootLandingCurveFromProfile oParentSystem, oRootLC
    
    ' ---------------------------------
    ' Get the operation that created it
    ' ---------------------------------
    Dim oProfileGraphicRepr As IJStructGraphicRepresentation
    Set oProfileGraphicRepr = oRootLC
    
    Dim oLCActiveEntity As Object
    Dim oAEHelper As StructGenericUtilities.StructBOHelper
    Set oAEHelper = New StructBOHelper
    
    oAEHelper.GetCreateActiveEntity oProfileGraphicRepr, oLCActiveEntity

    ' ------------------------------------------
    ' If by elements, it is a tripping stiffener
    ' ------------------------------------------
    If TypeOf oLCActiveEntity Is IJLandCrvBetweenElements_AE Then
        ' ----------------------------------------
        ' Determine which end we are on
        ' ----------------------------------------
        Dim oBoundedStructPort As IJStructPort
        Set oBoundedStructPort = oBoundedPort
        
        Dim eCtx As eUSER_CTX_FLAGS
        eCtx = oBoundedStructPort.ContextID
        
        Dim bIsStart As Boolean
        bIsStart = False
        If eCtx And CTX_BASE Then
            bIsStart = True
        End If
        
        ' -------------------------------
        ' Determine if lapped at this end
        ' -------------------------------
        Dim oLCBE As IJLandCrvBetweenElements_AE
        Set oLCBE = oLCActiveEntity
        
        Dim endCond As LandingCrvAttachmentMethod
        endCond = oLCBE.AttachmentType(bIsStart)
    
        If endCond = LCA_LAPPED Then
            IsLapped = True
            
            Dim uOff As Double
            Dim vOff As Double
            uOff = Abs(oLCBE.AttachmentUOffset(bIsStart))
            vOff = Abs(oLCBE.AttachmentVOffset(bIsStart))
            
            ' Expect that one of these is zero and one is greater than zero
            If uOff > vOff Then
                dLapDist = uOff
            Else
                dLapDist = vOff
            End If
        End If
    End If
End Function

Public Sub SelectAlongAxisEndToEndWebCut( _
                        ByVal sEndCutType As String, _
                        ByVal sWeldPartNumber As String, _
                        ByVal oSLH As IJDSelectorLogic)
   On Error GoTo ErrorHandler
    
   Dim sError As String
    
   Select Case sEndCutType
      Case "Welded"
         If sWeldPartNumber = "First" Then
            oSLH.Add "WebCut_AlongAxisPCEnd"
         Else
            oSLH.Add "WebCut_AlongAxisEnd"
         End If
           
      Case Else
   End Select
    
   Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectAlongAxisEndToEndWebCut", sError).Number
End Sub

Public Sub SelectAlongAxisEndToEndFlangeCut( _
                        ByVal sEndCutType As String, _
                        ByVal sWeldPartNumber As String, _
                        ByVal oSLH As IJDSelectorLogic)
   On Error GoTo ErrorHandler
    
   Dim sError As String
   Dim oFlangeCutWrapper As New StructDetailObjects.FlangeCut
   Dim oBoundedPart As Object
   
   Set oFlangeCutWrapper.object = oSLH.SmartOccurrence
   Set oBoundedPart = oFlangeCutWrapper.Bounded
   
   Dim sSectionType As String
   
   If TypeOf oBoundedPart Is IJStiffenerPart Then
      Dim oSDProfilePartWrapper As New StructDetailObjects.ProfilePart
      
      Set oSDProfilePartWrapper.object = oBoundedPart
      sSectionType = oSDProfilePartWrapper.sectionType
      
   ElseIf TypeOf oBoundedPart Is IJBeamPart Then
      Dim oSDBeamPartWrapper As New StructDetailObjects.BeamPart
      
      Set oSDBeamPartWrapper.object = oBoundedPart
      sSectionType = oSDBeamPartWrapper.sectionType
   Else
      sError = "Unknown part type"
      GoTo ErrorHandler
   End If
   
   Select Case sEndCutType
      Case "Welded"
         If sWeldPartNumber = "First" Then
            If sSectionType = "EA" Or _
               sSectionType = "UA" Or _
               sSectionType = gsB Then
               oSLH.Add "FlangeCut_AlongAxisAPCEnd"
            Else
               oSLH.Add "FlangeCut_AlongAxisTPCEnd"
            End If
         Else
            If sSectionType = "EA" Or _
               sSectionType = "UA" Or _
               sSectionType = gsB Then
               oSLH.Add "FlangeCut_AlongAxisAEnd"
            Else
               oSLH.Add "FlangeCut_AlongAxisTEnd"
            End If
         End If
           
      Case Else
         sError = "Unhandled end cut type: " & sEndCutType
         GoTo ErrorHandler
         
   End Select
    
   Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectAlongAxisEndToEndFlangeCut", sError).Number
End Sub

Public Function GetEndCutBoundingObjectFromAC( _
                ByVal oEndCut As Object, _
                Optional ByVal sWeldPart As String = "") As Object
   Set GetEndCutBoundingObjectFromAC = Nothing
        
   Dim oDesignChild As IJDesignChild
   Dim oEndCutParent As Object

   Set oDesignChild = oEndCut
   Set oEndCutParent = oDesignChild.GetParent
   Set oDesignChild = Nothing
   
   If TypeOf oEndCutParent Is IJAssemblyConnection Then
      If sWeldPart = "" Then
         Dim oCommonHelper As DefinitionHlprs.CommonHelper
         Dim vAnswer As Variant
         
         Dim oSO As IJSmartOccurrence
         Dim oSI As IJSmartItem
         Dim oSC As IJSmartClass
         
         Set oSO = oEndCut
         Set oSI = oSO.ItemObject
         If Not oSI Is Nothing Then
            Set oSC = oSI.Parent
           
            Set oCommonHelper = New DefinitionHlprs.CommonHelper
            vAnswer = oCommonHelper.GetAnswer(oSO, _
                                              oSC.SelectionRuleDef, _
                                              gsWeldPart)
            sWeldPart = vAnswer
         End If
      End If

      Dim oACWrapper As New StructDetailObjects.AssemblyConn
      Dim oBoundingObject As Object
      
      Set oACWrapper.object = oEndCutParent
      If sWeldPart = "First" Then
         Set oBoundingObject = oACWrapper.Port2 ' This part has PC connection
      Else
         Set oBoundingObject = oACWrapper.Port1 '
      End If
      
      Set oACWrapper = Nothing
   End If
    
   Set GetEndCutBoundingObjectFromAC = oBoundingObject
   
   Exit Function
   
End Function

Public Function CreatePCForAlongAxisEndCut( _
            ByVal oMD As IJDMemberDescription, _
            ByVal oResourceManager As IUnknown, _
            ByVal sStartClass As String, _
            ByVal eBoundingSubPort As JXSEC_CODE, _
            ByVal eBoundedSubPort As JXSEC_CODE) As Object
   On Error GoTo ErrorHandler
    
   Dim oEndCut As Object
   Dim oStructFeature As IJStructFeature
   Dim eStructFeature As StructFeatureTypes
   Dim sError As String
   
   Set oEndCut = oMD.CAO
   Set oStructFeature = oEndCut
   eStructFeature = oStructFeature.get_StructFeatureType

   Dim oBoundingPort As IJPort
   Dim oBoundedPart As Object
   
   ' Get bounding pPart's Port
   
   sError = "Getting bounding port"
   If eStructFeature = SF_WebCut Then
      Dim oWebCutWrapper As New StructDetailObjects.WebCut
      
      Set oWebCutWrapper.object = oEndCut
      If TypeOf oWebCutWrapper.BoundingPort Is IJPort Then
         Set oBoundingPort = oWebCutWrapper.BoundingPort
      Else
         ' Following call is particular to end to end along axis end cut
         Set oBoundingPort = GetEndCutBoundingObjectFromAC(oEndCut)
      End If
      
      Set oBoundedPart = oWebCutWrapper.Bounded
      
      Set oWebCutWrapper = Nothing
      
   Else
      Dim oFlangeCutWrapper As New StructDetailObjects.FlangeCut
      
      Set oFlangeCutWrapper.object = oEndCut
      If TypeOf oFlangeCutWrapper.BoundingPort Is IJPort Then
         Set oBoundingPort = oFlangeCutWrapper.BoundingPort
      Else
         ' Following call is particular to end to end along axis end cut
         Set oBoundingPort = GetEndCutBoundingObjectFromAC(oEndCut)
      End If
      
      Set oBoundedPart = oFlangeCutWrapper.Bounded

      Set oFlangeCutWrapper = Nothing
   End If

   Dim oSDOHelper As New StructDetailObjects.Helper
   
   Set oBoundingPort = oSDOHelper.GetEquivalentLastPort(oBoundingPort)

   ' Get bounded Part's Port
   Dim oBoundedPort As IJPort
   
   sError = "Get bounded port"
  
   If oSDOHelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
      Dim oProfilePartWrapper As New StructDetailObjects.ProfilePart
      
      Set oProfilePartWrapper.object = oBoundedPart
      Set oBoundedPort = oProfilePartWrapper.CutoutSubPort( _
                                  oEndCut, _
                                  eBoundedSubPort)
      Set oProfilePartWrapper = Nothing
      
   ElseIf oSDOHelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
      Dim oBeamPartWrapper As New StructDetailObjects.BeamPart
      
      Set oBeamPartWrapper.object = oBoundedPart
      Set oBoundedPort = oBeamPartWrapper.CutoutSubPort( _
                                  oEndCut, _
                                  eBoundedSubPort)
      Set oBeamPartWrapper = Nothing
   End If
   Set oBoundedPart = Nothing
   Set oSDOHelper = Nothing
   
   ' Create physical connection
   Dim oPCWrapper As New StructDetailObjects.PhysicalConn
  
   sError = "Creating Physical Connection"
   oPCWrapper.Create _
                   oResourceManager, _
                   oBoundedPort, _
                   oBoundingPort, _
                   sStartClass, _
                   oEndCut, _
                   ConnectionStandard
                   
   Set CreatePCForAlongAxisEndCut = oPCWrapper.object
  
   Set oBoundingPort = Nothing
   Set oBoundedPort = Nothing
   Set oPCWrapper = Nothing
   Set oEndCut = Nothing
  
   Exit Function
    
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "CreatePCForAlongAxisEndCut", sError).Number
End Function

Public Sub GetHeightAndFlangeThickness( _
               ByVal oObject As Object, _
               ByRef dHeight As Double, _
               ByRef dFlangeThickness As Double)
   On Error GoTo ErrorHandler
   Dim sMsg As String
   
   If oObject Is Nothing Then
      sMsg = "Invalid input"
      GoTo ErrorHandler
   End If
   
   Dim oSDOHelper As New StructDetailObjects.Helper
   Dim eObjType As sdwObjectType
   
   dHeight = 0
   dFlangeThickness = 0
   eObjType = oSDOHelper.ObjectType(oObject)
   Set oSDOHelper = Nothing
   
   If eObjType = SDOBJECT_STIFFENER Then
      sMsg = "Profile part"
      Dim oProfilePartWrapper As New StructDetailObjects.ProfilePart
      
      Set oProfilePartWrapper.object = oObject
      dHeight = oProfilePartWrapper.Height
      dFlangeThickness = oProfilePartWrapper.flangeThickness
      Set oProfilePartWrapper = Nothing
      
   ElseIf eObjType = SDOBJECT_BEAM Then
      sMsg = "Beam part"
      Dim oBeamPartWrapper As New StructDetailObjects.BeamPart
      
      Set oBeamPartWrapper.object = oObject
      dHeight = oBeamPartWrapper.Height
      dFlangeThickness = oBeamPartWrapper.flangeThickness
      Set oBeamPartWrapper = Nothing
      
   ElseIf eObjType = SDOBJECT_MEMBER Then
      sMsg = "Member part"
      Dim oMemberPartWrapper As New StructDetailObjects.MemberPart
      
      Set oMemberPartWrapper.object = oObject
      dHeight = oMemberPartWrapper.Height
      dFlangeThickness = oMemberPartWrapper.flangeThickness
      Set oMemberPartWrapper = Nothing
      
   End If
   
   Exit Sub
    
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "GetHeightAndFlangeThickness", sMsg).Number
                      
End Sub
'
' This method check if bounded beam V direction and bounding landing curve direction are the same
'
Public Function IsBoundaryLandindCurveAlongBeamVDirection( _
        ByVal oBoundedPart As Object, _
        ByVal oBoundingPart As Object, _
        ByVal oWebCut As Object) As Boolean
   On Error GoTo ErrorHandler

   IsBoundaryLandindCurveAlongBeamVDirection = False
   If oBoundedPart Is Nothing Or _
      oBoundingPart Is Nothing Or _
      oWebCut Is Nothing Then
      Exit Function
   End If
   
   If Not TypeOf oBoundingPart Is IJProfile Or _
      Not TypeOf oBoundingPart Is IJProfileER Then
      Exit Function
   End If
   
   If Not TypeOf oBoundedPart Is IJBeamPart Then
      Exit Function
   End If
   
   '    v along W            v perpendicular to W
   '
   '  Bounded  Bounding     Bounded  Bounding Along
   '                                    V
   '              ||                    ^
   '  ----------- ||        ----------- ||
   '              ||                    ||
   '  ----------- ||        ----------- ||
   '              ||
   '
   '  v           W                     V
   '  ^           ^                     ^
   '  |           |                     |
   '  x--> w      x --->U               .-->U
   '  u           V                     W
   '
   '
   ' Get web cut location
   Dim oWebCutWrapper As StructDetailObjects.WebCut
   Dim oWebCutLocation As IJDPosition
  
   Set oWebCutWrapper = New StructDetailObjects.WebCut
   Set oWebCutWrapper.object = oWebCut
   Set oWebCutLocation = oWebCutWrapper.BoundedLocation
   
   ' Get beam matrix at web cut
   Dim oBeamPartWrapper As StructDetailObjects.BeamPart
   Dim oSystem As Object
   Dim oBoundedMatrixAtCut As IJDT4x4
   
   Set oBeamPartWrapper = New StructDetailObjects.BeamPart
   Set oBeamPartWrapper.object = oBoundedPart
   Set oSystem = oBeamPartWrapper.ParentSystem
   Set oBeamPartWrapper = Nothing
   
   Dim oProfileAttributes As IJProfileAttributes
   
   Set oProfileAttributes = New ProfileUtils
   oProfileAttributes.GetProfileOrientationMatrix oSystem, oWebCutLocation, oBoundedMatrixAtCut
   Set oSystem = Nothing
   
   ' Get bounding matrix at web cut
   Dim oBoundingMatrixAtCut As IJDT4x4
   Dim oProfilePartWrapper As StructDetailObjects.ProfilePart
      
   Set oProfilePartWrapper = New StructDetailObjects.ProfilePart
   Set oProfilePartWrapper.object = oWebCutWrapper.Bounding
   Set oSystem = oProfilePartWrapper.ParentSystem
   Set oProfilePartWrapper = Nothing
   oProfileAttributes.GetProfileOrientationMatrix oSystem, oWebCutLocation, oBoundingMatrixAtCut
   Set oSystem = Nothing
   Set oWebCutLocation = Nothing
   Set oProfileAttributes = Nothing

   If Not oBoundedMatrixAtCut Is Nothing And _
      Not oBoundingMatrixAtCut Is Nothing Then
      Dim oBoundedVAtCut As IJDVector
      Dim oBoundingWAtCut As IJDVector
   
      Set oBoundedVAtCut = New DVector
      Set oBoundingWAtCut = New DVector

      oBoundedVAtCut.x = oBoundedMatrixAtCut.IndexValue(4)
      oBoundedVAtCut.y = oBoundedMatrixAtCut.IndexValue(5)
      oBoundedVAtCut.z = oBoundedMatrixAtCut.IndexValue(6)

      oBoundingWAtCut.x = oBoundingMatrixAtCut.IndexValue(8)
      oBoundingWAtCut.y = oBoundingMatrixAtCut.IndexValue(9)
      oBoundingWAtCut.z = oBoundingMatrixAtCut.IndexValue(10)

      Set oBoundedMatrixAtCut = Nothing
      Set oBoundingMatrixAtCut = Nothing

      Dim dDot As Double

      dDot = oBoundedVAtCut.Dot(oBoundingWAtCut)
      Set oBoundedVAtCut = Nothing
      Set oBoundingWAtCut = Nothing
      If (Abs(dDot) > 0.9) Then
         IsBoundaryLandindCurveAlongBeamVDirection = True
      Else
         IsBoundaryLandindCurveAlongBeamVDirection = False
      End If
   End If
   
   Exit Function
   
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "IsBoundaryLandindCurveAlongBeamVDirection", "Error").Number
End Function
 Public Sub GetBaseOrOffsetPortsAtAC(oAssembConn As StructDetailObjects.AssemblyConn, ByRef pBoundingPort As IJPort, ByRef pBoundedPort As IJPort)
    On Error GoTo ErrorHandler
    
    Dim phelper As New StructDetailObjects.Helper
    Dim oEndCutPosition As IJDPosition
    Set oEndCutPosition = New DPosition
    
    Dim oBasePort1 As IJPort, oBasePort2 As IJPort, oOffsetPort1 As IJPort, oOffsetPort2 As IJPort
    
    Set oEndCutPosition = oAssembConn.BoundGlobalShipLocation
    
    Dim oSDO_ProfilePart1 As New StructDetailObjects.ProfilePart
    Dim oSDO_ProfilePart2 As New StructDetailObjects.ProfilePart
    
    Dim oSDO_MemberPart1 As New StructDetailObjects.MemberPart
    Dim oSDO_MemberPart2 As New StructDetailObjects.MemberPart
    
    If phelper.ObjectType(oAssembConn.ConnectedObject1) = SDOBJECT_STIFFENER Then
        Set oSDO_ProfilePart1.object = oAssembConn.ConnectedObject1
        Set oBasePort1 = oSDO_ProfilePart1.BasePort(BPT_Base)
        Set oOffsetPort1 = oSDO_ProfilePart1.BasePort(BPT_Offset)
    ElseIf phelper.ObjectType(oAssembConn.ConnectedObject1) = SDOBJECT_BEAM Then
        Set oSDO_MemberPart1.object = oAssembConn.ConnectedObject1
        Set oBasePort1 = oSDO_MemberPart1.BasePort(BPT_Base)
        Set oOffsetPort1 = oSDO_MemberPart1.BasePort(BPT_Offset)
    End If
    
    Dim oPoint As IJDPosition
    Set oPoint = New DPosition
    
    Dim oVector As IJDVector
    Set oVector = New DVector
    
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
    oTopologyLocate.GetProjectedPointOnModelBody oBasePort1.Geometry, _
                                                 oEndCutPosition, _
                                                 oPoint, oVector
    Dim dDistance1 As Double
    Dim dDistance2 As Double
    
    dDistance1 = oEndCutPosition.DistPt(oPoint)
    Set oPoint = Nothing
    
    oTopologyLocate.GetProjectedPointOnModelBody oOffsetPort1.Geometry, _
                                                 oEndCutPosition, _
                                                 oPoint, oVector
    dDistance2 = oEndCutPosition.DistPt(oPoint)
    Set oPoint = Nothing
    
    If dDistance1 < dDistance2 Then
        Set pBoundedPort = oSDO_ProfilePart1.BasePortBeforeTrim(BPT_Base)
    Else
        Set pBoundedPort = oSDO_ProfilePart1.BasePortBeforeTrim(BPT_Offset)
    End If

    If phelper.ObjectType(oAssembConn.ConnectedObject2) = SDOBJECT_STIFFENER Then
        Set oSDO_ProfilePart2.object = oAssembConn.ConnectedObject2
        Set oBasePort2 = oSDO_ProfilePart2.BasePort(BPT_Base)
        Set oOffsetPort2 = oSDO_ProfilePart2.BasePort(BPT_Offset)
    ElseIf phelper.ObjectType(oAssembConn.ConnectedObject2) = SDOBJECT_BEAM Then
        Set oSDO_MemberPart2.object = oAssembConn.ConnectedObject2
        Set oBasePort2 = oSDO_MemberPart2.BasePort(BPT_Base)
        Set oOffsetPort2 = oSDO_MemberPart2.BasePort(BPT_Offset)
    End If
    oTopologyLocate.GetProjectedPointOnModelBody oBasePort2.Geometry, _
                                                 oEndCutPosition, _
                                                 oPoint, oVector
    dDistance1 = oEndCutPosition.DistPt(oPoint)
    Set oPoint = Nothing
    
    oTopologyLocate.GetProjectedPointOnModelBody oOffsetPort2.Geometry, _
                                                 oEndCutPosition, _
                                                 oPoint, oVector
    dDistance2 = oEndCutPosition.DistPt(oPoint)
    Set oPoint = Nothing
    
    If dDistance1 < dDistance2 Then
        Set pBoundingPort = oSDO_ProfilePart2.BasePortBeforeTrim(BPT_Base)
    Else
        Set pBoundingPort = oSDO_ProfilePart2.BasePortBeforeTrim(BPT_Offset)
    End If
    
    Exit Sub
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "IsBoundaryLandindCurveAlongBeamVDirection", "Error").Number

 End Sub
