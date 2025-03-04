Attribute VB_Name = "Common"
Option Explicit

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\EndCutRules\Common.bas"

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
    Dim oFlangeCut As Structdetailobjects.FlangeCut
    Set oFlangeCut = New Structdetailobjects.FlangeCut
    Set oFlangeCut.object = pSLH.SmartOccurrence
    bFreeEndCut = oFlangeCut.IsFreeEndCut
    
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
    
    Dim phelper As New Structdetailobjects.Helper
    
    ' Determine the Height of the BoundingObject
    Dim bWebCut_C2Spline As Boolean
    Dim dBoundingHeight As Double
    
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
        Dim oSDO_WebCut As Structdetailobjects.WebCut
        Set oSDO_WebCut = New Structdetailobjects.WebCut
        Set oSDO_WebCut.object = oFlangeCut.WebCut
        If LCase(Trim(oSDO_WebCut.ItemName)) = LCase("WebCut_C2Spline") Then
            bWebCut_C2Spline = True
        End If
        
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
        Dim oBoundingProfile As New Structdetailobjects.ProfilePart
        Set oBoundingProfile.object = oBoundingPart
        dBoundingHeight = oBoundingProfile.Height
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_BEAM Then
        Dim oBoundingBeam As New Structdetailobjects.BeamPart
        Set oBoundingBeam.object = oBoundingPart
        dBoundingHeight = oBoundingBeam.Height
    Else
        ' Bounding Part type is unknown, use Default value
        dBoundingHeight = 10    'meters
    End If
    
    'get bounded SectionType and Height
    Dim sSectionType As String
    Dim dBoundedHeight As Double
    If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
        Dim oBoundedProfile As New Structdetailobjects.ProfilePart
        Set oBoundedProfile.object = oBoundedPart
        sSectionType = oBoundedProfile.SectionType
        dBoundedHeight = oBoundedProfile.Height
        
    ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        Dim oBoundedBeam As New Structdetailobjects.BeamPart
        Set oBoundedBeam.object = oBoundedPart
        sSectionType = oBoundedBeam.SectionType
        dBoundedHeight = oBoundedBeam.Height
        
    Else
        'Bounded Part must be a Stiffener(or Edge Reinforcment) or Beam
        Exit Sub
    End If
        
    ' Check if Bounded Object has only Right Flange
    Dim bRightFlangeOnly As Boolean
    If (sSectionType = "EA") Or _
       (sSectionType = "UA") Or _
       (sSectionType = "B") Then
       bRightFlangeOnly = True
    Else
       bRightFlangeOnly = False
    End If
        
    ' the following is valid for both
    ' SDOBJECT_STIFFENER and SDOBJECT_BEAM bounded Part types
    If bFreeEndCut Then
        'For a FreeEnd FlangeCut,
        Select Case strEndCutType
            Case gsR
            Case gsRV
                ' Report but do not raise error
                strError = "Not handled -- FreeEnd Stiffener/ Beam"
                LogError Err, MODULE, "FlangeCutNoFlange"
                
            Case gsF, gsFV, gsS, gsC
                If bRightFlangeOnly Then
                    pSLH.Add "FlangeCut_F1"
                Else
                    pSLH.Add "FlangeCut_F4"
                End If
                
            Case gsW
                If bRightFlangeOnly Then
                    pSLH.Add "FreeEndFlangeCut_W1"
                Else
                    pSLH.Add "FreeEndFlangeCut_W4"
                End If
        End Select
        
        Exit Sub
    End If
               
    
    If phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
        'Profile bounded by profile case.
        'If the landing curves are not intersecting, we would go with a simple flange cut.
        If Not AreLandingCurvesIntersecting(oBoundedPart, oBoundingPart) Then
            If bRightFlangeOnly Then
                pSLH.Add "FlangeCut_F1"
            Else
                pSLH.Add "FlangeCut_F4"
            End If
            
            Exit Sub
        End If
    End If

    ' the following is valid for both
    ' SDOBJECT_STIFFENER and SDOBJECT_BEAM bounded Part types
    Select Case strEndCutType
        Case gsR
        Case gsRV
            ' Report but do not raise error
            strError = "Not handled -- bounded Stiffener/ Beam"
            LogError Err, MODULE, "FlangeCutNoFlange"
            
        Case gsF, gsFV, gsS, gsC
            If bRightFlangeOnly Then
                    pSLH.Add "FlangeCut_F1"
                Else
                    pSLH.Add "FlangeCut_F4"
            End If
            
        Case gsW
            If (dBoundingHeight - 0.015) >= dBoundedHeight Then
                
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
                    If sSectionType = "B" Then
                        pSLH.Add "FlangeCut_BW1"
                    Else
                        pSLH.Add "FlangeCut_W1"
                    End If
                Else
                    pSLH.Add "FlangeCut_W4"
                End If
            
            Else
                If bRightFlangeOnly Then
                    pSLH.Add "FlangeCut_F1"
                Else
                    pSLH.Add "FlangeCut_F4"
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
    Dim phelper As New Structdetailobjects.Helper
    Dim Dist As Double, alpha As Double
        
    If phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
        Dim oBoundingProfile As New Structdetailobjects.ProfilePart
        Set oBoundingProfile.object = oBoundingPart
        dBoundingHeight = oBoundingProfile.Height
        dFlangeThickness = oBoundingProfile.FlangeThickness
        If oBoundingProfile.SectionType = "B" Then
            Dist = 0.01
        Else
            Dist = 0.035
        End If
    
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_BEAM Then
        Dim oBoundingBeam As New Structdetailobjects.BeamPart
        Set oBoundingBeam.object = oBoundingPart
        dBoundingHeight = oBoundingBeam.Height
        dFlangeThickness = oBoundingBeam.FlangeThickness
        If oBoundingBeam.SectionType = "B" Then
            Dist = 0.01
        Else
            Dist = 0.035
        End If
       
    Else
        Dist = 0.035
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
    Dim oBoundedProfile As Structdetailobjects.ProfilePart
    Dim oBoundedBeam As Structdetailobjects.BeamPart
    
    If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
        Set oBoundedProfile = New Structdetailobjects.ProfilePart
        Set oBoundedProfile.object = oBoundedPart
        strSectionType = oBoundedProfile.SectionType
        
    ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        Set oBoundedBeam = New Structdetailobjects.BeamPart
        Set oBoundedBeam.object = oBoundedPart
        strSectionType = oBoundedBeam.SectionType
    End If
    
    If phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
        
        'Profile bounded by profile case.
        'If the landing curves are not intersecting, we would go with a simple flange cut.
        If Not AreLandingCurvesIntersecting(oBoundedPart, oBoundingPart) Then
            If strSectionType = "EA" Or _
               strSectionType = "UA" Or _
               strSectionType = "B" Then
               
                pSLH.Add "FlangeCut_F1"
            Else
                pSLH.Add "FlangeCut_F4"
            End If
                
            Exit Sub
        End If
    End If
    
    Dim oBoundedPort As IJPort
    Set oBoundedPort = oBoundedObject
    
    'Get the name of the port, used to determine if part is connected
    'to the web or the flange
    Dim dPortValue As Long
    Dim oPartInfo As New Structdetailobjects.Helper
    dPortValue = oPartInfo.GetBoundingProfileFace(oBoundedPort, oBoundingPort)
    
    'Hardcode the angle to 90 deg
    alpha = 90
    
    'get bounded height
    Dim dBoundedHeight As Double
    
    If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
        dBoundedHeight = oBoundedProfile.Height
        
            Select Case strEndCutType
                Case gsR, gsRV
                    strError = (" R, RV are not allowed if there is a flange")
                    LogError Err, MODULE, "FlangeCutFlanged", strError   ' Report but do not raise error
                
                Case gsF, gsFV, gsS, gsC
                    If strSectionType = "EA" Or _
                       strSectionType = "UA" Or _
                       strSectionType = "B" Then
                        pSLH.Add "FlangeCut_F1"
                    Else
                        pSLH.Add "FlangeCut_F4"
                    End If
                        
                Case gsW
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
                            pSLH.Add "FlangeCut_W1"
                        ElseIf strSectionType = "B" Then
                            pSLH.Add "FlangeCut_BW1"
                        Else
                            pSLH.Add "FlangeCut_W4"
                        End If
                    Else
                        If (dBoundingHeight - dFlangeThickness - Dist) >= dBoundedHeight Then
                            If strSectionType = "EA" Or _
                               strSectionType = "UA" Then
                                pSLH.Add "FlangeCut_W1"
                            ElseIf strSectionType = "B" Then
                                pSLH.Add "FlangeCut_BW1"
                            Else
                                pSLH.Add "FlangeCut_W4"
                            End If
                        Else
                            If (oBoundedProfile.SectionType = "EA") Or _
                                (oBoundedProfile.SectionType = "UA") Or _
                                (oBoundedProfile.SectionType = "B") Then
                                    pSLH.Add "FlangeCut_F1"
                            Else
                                    pSLH.Add "FlangeCut_F4"
                            End If
                        End If
                    End If
        
            End Select
    
    ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        dBoundedHeight = oBoundedBeam.Height

            Select Case strEndCutType
                Case gsR, gsRV
                    strError = " R, RV are not allowed if there is a flange"
                    LogError Err, MODULE, "FlangeCutFlanged", strError   ' Report but do not raise error
                
                Case gsF, gsFV, gsS, gsC
                    If strSectionType = "EA" Or _
                       strSectionType = "UA" Or _
                       strSectionType = "B" Then
                        pSLH.Add "FlangeCut_F1"
                    Else
                        pSLH.Add "FlangeCut_F4"
                    End If
                        
                Case gsW
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
                            pSLH.Add "FlangeCut_W1"
                        ElseIf strSectionType = "B" Then
                            pSLH.Add "FlangeCut_BW1"
                        Else
                            pSLH.Add "FlangeCut_W4"
                        End If
                    Else
                        If (dBoundingHeight - dFlangeThickness - Dist) >= dBoundedHeight Then
                            If strSectionType = "EA" Or _
                               strSectionType = "UA" Then
                                pSLH.Add "FlangeCut_W1"
                            ElseIf strSectionType = "B" Then
                                pSLH.Add "FlangeCut_BW1"
                            Else
                                pSLH.Add "FlangeCut_W4"
                            End If
                        Else
                            If strSectionType = "EA" Or _
                               strSectionType = "UA" Or _
                               strSectionType = "B" Then
                                    pSLH.Add "FlangeCut_F1"
                            Else
                                    pSLH.Add "FlangeCut_F4"
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
    Dim phelper As New Structdetailobjects.Helper
    
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

    'get bounded height
    Dim dBoundedHeight As Double
    If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
        Dim oBoundedProfile As New Structdetailobjects.ProfilePart
        Set oBoundedProfile.object = oBoundedPart
        dBoundedHeight = oBoundedProfile.Height
        Select Case strEndCutType
            Case gsW
                If strWeldPartNumber = "First" Then
                    If (oBoundedProfile.SectionType = "EA") Or _
                        (oBoundedProfile.SectionType = "UA") Or _
                        (oBoundedProfile.SectionType = "B") Then
                       pSLH.Add "FlangeCut_W1PCEnd"
zMsgBox "::FlangeCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add FlangeCut_W1PCEnd"
        
                    Else
                        pSLH.Add "FlangeCut_W4PCEnd"
zMsgBox "::FlangeCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add FlangeCut_W4PCEnd"
                    End If
                Else
                    If (oBoundedProfile.SectionType = "EA") Or _
                        (oBoundedProfile.SectionType = "UA") Or _
                        (oBoundedProfile.SectionType = "B") Then
                        pSLH.Add "FlangeCut_W1End"
zMsgBox "::FlangeCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add FlangeCut_W1End"
                    Else
                        pSLH.Add "FlangeCut_W4End"
zMsgBox "::FlangeCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add FlangeCut_W4End"
                    End If
                End If
            
            Case Else
zMsgBox "::FlangeCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add (Nothing)"
        End Select
        
    ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        Dim oBoundedBeam As New Structdetailobjects.BeamPart
        Set oBoundedBeam.object = oBoundedPart
        dBoundedHeight = oBoundedBeam.Height
        Select Case strEndCutType
            Case gsW
                If strWeldPartNumber = "First" Then
                    If (oBoundedBeam.SectionType = "EA") Or _
                        (oBoundedBeam.SectionType = "UA") Or _
                        (oBoundedBeam.SectionType = "B") Then
                        pSLH.Add "FlangeCut_W1PCEnd"
                    Else
                        pSLH.Add "FlangeCut_W4PCEnd"
                    End If
                Else
                    If (oBoundedBeam.SectionType = "EA") Or _
                        (oBoundedBeam.SectionType = "UA") Or _
                        (oBoundedBeam.SectionType = "B") Then
                        pSLH.Add "FlangeCut_W1End"
                    Else
                        pSLH.Add "FlangeCut_W4End"
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
    Dim oSDO_Helper As Structdetailobjects.Helper
    Dim oEdgeReinforcement As Structdetailobjects.EdgeReinforcement
    
    
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
    Dim bWebCut_C2Spline As Boolean
    
    dBoundingHeight = 0#
    bWebCut_C2Spline = False

    Set oSDO_Helper = New Structdetailobjects.Helper
    If (bIsPlate) Then
        Dim oWebCut As New Structdetailobjects.WebCut
        Set oWebCut.object = pSLH.SmartOccurrence
                
        Dim dFlangeClearance As Double
        dFlangeClearance = oWebCut.BoundingPlateFlangeClearance
        
        Set oWebCut = Nothing
        
        If dFlangeClearance > 0.025 Then
            dBoundingHeight = 10  'meters
        Else
            ' Special case for EndCutType: W, C
            ' Bounding object is a IJPlate
            ' AND it is shorter then the Bounded Profile
            bWebCut_C2Spline = True
        End If
    Else
        If oSDO_Helper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
            'Profile bounded by profile case.
            'If the landing curves are not intersecting, we would go with a simple web cut.
        
            If Not AreLandingCurvesIntersecting(oBoundedPart, oBoundingPart) Then
                pSLH.Add "WebCut_C1Spline"
                Exit Sub
            End If
        End If
        
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
        Dim oBoundedProfile As New Structdetailobjects.ProfilePart
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
            Set oEdgeReinforcement = New Structdetailobjects.EdgeReinforcement
            Set oEdgeReinforcement.object = oBoundedPart
            
            'check if Edge Reinforcement is on edge
            If Not oEdgeReinforcement.ReinforcementPosition = "OnFace" Then
                'if the edge reinforcement a Flat bar.
                'If not a flat bar simply cut neat to bounding object regardless of end cut type.
                If oBoundedProfile.SectionType = "FB" Then
                    Select Case strEndCutType
                        Case gsF, gsFV, gsS
                            'Edge Reinforcement
                            'end cut type F, FV, or S.
                            'Select web cut that looks like a sniped Tee flange cut
                            pSLH.Add "FreeEndWebCut_F1_ER"
                        
                        Case Else
                            pSLH.Add "WebCut_C1Spline"
                        End Select
                Else
                    pSLH.Add "WebCut_C1Spline"
                End If
            End If
            
            ' if the edge reiforcement is on face, treat it like a stiffener
        End If
       
        Select Case strEndCutType
            Case gsR
                If dBoundingHeight < dBoundedHeight * 2 Then
                    strError = "bounding profile too short for R endcuttype"
                    LogError Err, MODULE, "WebCutNoFlange", strError   ' Report but do not raise error
                Else
                    pSLH.Add "WebCut_R1"
                End If
               
            Case gsRV
                If dBoundingHeight < dBoundedHeight * 2 Then
                    strError = "bounding profile too short for RV endcuttype"
                    LogError Err, MODULE, "WebCutNoFlange", strError   ' Report but do not raise error
                Else
                    pSLH.Add "WebCut_RV1"
                End If
               
            Case gsF
                If ((oBoundedProfile.SectionType = "B") Or _
                   (oBoundedProfile.SectionType = "FB")) Then
                    pSLH.Add "WebCut_F1B"
                Else
                    pSLH.Add "WebCut_F1"
                End If
                
            Case gsFV
                pSLH.Add "WebCut_FV1"
                
            Case gsS
                pSLH.Add "WebCut_S1"
               
            Case gsW, gsC
                If bWebCut_C2Spline Then
                    ' Special case for EndCutType: W, C
                    ' Bounding object is a IJPlate
                    ' AND it is shorter then the Bounded Profile
                    pSLH.Add "WebCut_C2Spline"
                   
                
                ElseIf (dBoundingHeight - 0.015) >= dBoundedHeight Then
                    If (bIsPlate) Then
                        'This end cut symbol uses a spline for the bounding curve.
                        'To minimize the impact on existing catalogs, the old WebCut_C1
                        'symbol is still used for cases where the bounding object is not
                        'a plate.
                        pSLH.Add "WebCut_C1Spline"
                        
                    Else
                        pSLH.Add "WebCut_C1Spline"
                    End If
                    
                Else
                    'Get the name of the port, used to determine side for the symbol
                    dPortValue = oSDO_Helper.GetBoundingProfileFace(oBoundedPort, oBoundingPort)
                    oSDO_Helper.GetObjectTypeData oBoundingPart, sTypeObject, sObjectType
                    If dPortValue = JXSEC_WEB_LEFT Or _
                        dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If sTypeObject = "EdgeReinforcement" Then
                                pSLH.Add "WebCut_C2Left_ER"
                            Else
                                pSLH.Add "WebCut_C2Left"
                            End If
                    ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If sTypeObject = "EdgeReinforcement" Then
                                pSLH.Add "WebCut_C2Right_ER"
                            Else
                                pSLH.Add "WebCut_C2Right"
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

                        pSLH.Add "WebCut_C1Spline"
                    End If
                End If
            
        End Select
        
    ElseIf oSDO_Helper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        Dim oBoundedBeam As New Structdetailobjects.BeamPart
        Set oBoundedBeam.object = oBoundedPart
        dBoundedHeight = oBoundedBeam.Height
        
        Select Case strEndCutType
            Case gsR
                If dBoundingHeight < dBoundedHeight * 2 Then
                    strError = "bounding profile too short for R endcuttype"
                    LogError Err, MODULE, "WebCutNoFlange", strError   ' Report but do not raise error
                Else
                    pSLH.Add "WebCut_R1"
                End If
            
            Case gsRV
                If dBoundingHeight < dBoundedHeight * 2 Then
                    strError = "bounding profile too short for RV endcuttype"
                    LogError Err, MODULE, "WebCutNoFlange", strError   ' Report but do not raise error
                Else
                    pSLH.Add "WebCut_RV1"
                End If
            
            Case gsF
                If ((oBoundedBeam.SectionType = "B") Or _
                   (oBoundedBeam.SectionType = "FB")) Then
                    pSLH.Add "WebCut_F1B"
                Else
                    pSLH.Add "WebCut_F1"
                End If
                
            Case gsFV
                pSLH.Add "WebCut_FV1"
                
            Case gsS
                pSLH.Add "WebCut_S1"
            
            Case gsW, gsC
                If (dBoundingHeight - 0.015) >= dBoundedHeight Then
                    If (bIsPlate) Then
                        'This end cut symbol uses a spline for the bounding curve.
                        'To minimize the impact on existing catalogs, the old WebCut_C1
                        'symbol is still used for cases where the bounding object is not
                        'a plate.
                        pSLH.Add "WebCut_C1Spline"
                    Else
                        pSLH.Add "WebCut_C1Spline"
                    End If
                    
                Else
                    'Get the name of the port, used to determine side for the symbol
                    dPortValue = oSDO_Helper.GetBoundingProfileFace(oBoundedPort, oBoundingPort)
                                    
                    If dPortValue = JXSEC_WEB_LEFT Or _
                        dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                        
                        pSLH.Add "WebCut_C2Left"
                    ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                        
                        pSLH.Add "WebCut_C2Right"
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
                        
                        pSLH.Add "WebCut_C1Spline"
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
    Dim Dist As Double
    Dim phelper As New Structdetailobjects.Helper
    
    If phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
        Dim oBoundingProfile As New Structdetailobjects.ProfilePart
        Set oBoundingProfile.object = oBoundingPart
        dBoundingHeight = oBoundingProfile.Height
        dFlangeThickness = oBoundingProfile.FlangeThickness
        If oBoundingProfile.SectionType = "B" Then
            Dist = 0.01
        Else
            Dist = 0.035
        End If
        Set oSDO_Bounding = oBoundingProfile
        
     ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_BEAM Then
        Dim oBoundingBeam As New Structdetailobjects.BeamPart
        Set oBoundingBeam.object = oBoundingPart
        dBoundingHeight = oBoundingBeam.Height
        dFlangeThickness = oBoundingBeam.FlangeThickness
        If oBoundingBeam.SectionType = "B" Then
            Dist = 0.01
        Else
            Dist = 0.035
        End If
        Set oSDO_Bounding = oBoundingBeam
    
    Else
            Dist = 0.035
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
    
    If phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
        'Profile bounded by profile case.
        'If the landing curves are not intersecting, we would go with a simple web cut.
        If Not AreLandingCurvesIntersecting(oBoundedPart, oBoundingPart) Then
            pSLH.Add "WebCut_C1Spline"
            Exit Sub
        End If
    End If
        
    Dim oBoundedPort As IJPort
    Set oBoundedPort = oBoundedObject
    
    'Get the name of the port, used to determine side for the symbol
    Dim dPortValue As Long
    dPortValue = phelper.GetBoundingProfileFace(oBoundedPort, oBoundingPort)
   
    'check to see if the bounded part is a beam or a stiffener
    Dim dBoundedHeight As Double
    If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
        Dim oBoundedProfile As New Structdetailobjects.ProfilePart
        Set oBoundedProfile.object = oBoundedPart
        dBoundedHeight = oBoundedProfile.Height
    
        Call GetMeasurementSymbolData(pSLH, oBoundedProfile, oSDO_Bounding, dPortValue, alpha)
        
        Select Case strEndCutType
            Case gsR, gsRV
                strError = " R, RV are not allowed if there is a flange"
                LogError Err, MODULE, "WebCutFlanged"   ' Report but do not raise error
            
            Case gsF, gsFV, gsS
                If ((oBoundedProfile.SectionType = "B") Or _
                   (oBoundedProfile.SectionType = "FB")) Then
                    pSLH.Add "WebCut_F1B"
                Else
                    pSLH.Add "WebCut_F1"
                End If
            
            Case gsW
                If ((oBoundedProfile.SectionType = "BUT") Or _
                    (oBoundedProfile.SectionType = "BUTL3") Or _
                    (oBoundedProfile.SectionType = "T_XType") Or _
                    (oBoundedProfile.SectionType = "TSType") Or _
                    (oBoundedProfile.SectionType = "BUTL2")) And _
                    (dBoundingHeight = dBoundedHeight) Then
                    If 1.588249619 >= alpha And alpha >= 1.553343034 Then
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
                                pSLH.Add "WebCutI_C3Left"
                            ElseIf (oSDO_Bounding.SectionType = "T_XType") Or _
                                    (oSDO_Bounding.SectionType = "TSType") Then
                                pSLH.Add "WebCutT_W1Left"
                            Else
                                pSLH.Add "WebCut_W1Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.SectionType = "B") Then
                                pSLH.Add "WebCut_C3RBulb"
                            ElseIf (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
                                pSLH.Add "WebCutI_C3Right"
                            ElseIf (oSDO_Bounding.SectionType = "T_XType") Or _
                                   (oSDO_Bounding.SectionType = "TSType") Then
                                pSLH.Add "WebCutT_W1Right"
                            Else
                                pSLH.Add "WebCut_W1Right"
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
                            pSLH.Add "WebCut_C1Spline"
                        End If
                    ElseIf alpha < 1.553343034 Then
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
                                pSLH.Add "WebCutI_C3Left"
                            Else
                                pSLH.Add "WebCut_C3Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.SectionType = "B") Then
                                pSLH.Add "WebCut_C3RBulb"
                            ElseIf (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
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
                            pSLH.Add "WebCut_C1Spline"
                        End If
                    ElseIf alpha > 1.588249619 Then
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
                                pSLH.Add "WebCutI_C3Left"
                            Else
                                pSLH.Add "WebCut_C3Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
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
                            pSLH.Add "WebCut_C1Spline"
                        End If
                    End If
                Else
                    If (dBoundingHeight - dFlangeThickness - Dist) >= dBoundedHeight Then
                        pSLH.Add "WebCut_C1Spline"
                    Else
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
                                pSLH.Add "WebCutI_C3Left"
                            Else
                                pSLH.Add "WebCut_C3Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.SectionType = "B") Then
                                pSLH.Add "WebCut_C3RBulb"
                            Else
                                If (oSDO_Bounding.SectionType = "I") Or _
                                    (oSDO_Bounding.SectionType = "ISType") Then
                                    pSLH.Add "WebCutI_C3Right"
                                Else
                                    pSLH.Add "WebCut_C3Right"
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
                            pSLH.Add "WebCut_C1Spline"
                        End If
                    End If
                End If
            
            Case gsC
                If (dBoundingHeight - dFlangeThickness - Dist) >= dBoundedHeight Then
                    pSLH.Add "WebCut_C1Spline"
                Else
                    If dPortValue = JXSEC_WEB_LEFT Or _
                        dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                        If (oSDO_Bounding.SectionType = "I") Or _
                            (oSDO_Bounding.SectionType = "ISType") Then
                                pSLH.Add "WebCutI_C3Left"
                        Else
                            pSLH.Add "WebCut_C3Left"
                        End If
                    ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                        If (oSDO_Bounding.SectionType = "B") Then
                                pSLH.Add "WebCut_C3RBulb"
                        ElseIf (oSDO_Bounding.SectionType = "I") Or _
                              (oSDO_Bounding.SectionType = "ISType") Then
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
                            pSLH.Add "WebCut_C1Spline"
                    End If
                End If
                
            Set phelper = Nothing
            
        End Select
        
    ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
        Dim oBoundedBeam As New Structdetailobjects.BeamPart
        Set oBoundedBeam.object = oBoundedPart
        dBoundedHeight = oBoundedBeam.Height
    
        Call GetMeasurementSymbolData(pSLH, oBoundedBeam, oSDO_Bounding, dPortValue, alpha)
        
        Select Case strEndCutType
            Case gsR, gsRV
                strError = " R, RV are not allowed if there is a flange"
                LogError Err, MODULE, "WebCutFlanged"   ' Report but do not raise error
            
            Case gsF, gsFV, gsS
                If ((oBoundedBeam.SectionType = "B") Or _
                   (oBoundedBeam.SectionType = "FB")) Then
                    pSLH.Add "WebCut_F1B"
                Else
                    pSLH.Add "WebCut_F1"
                End If
            
            Case gsW
                If ((oBoundedBeam.SectionType = "BUT") Or _
                    (oBoundedBeam.SectionType = "BUTL3") Or _
                    (oBoundedBeam.SectionType = "T_XType") Or _
                    (oBoundedBeam.SectionType = "TSType") Or _
                    (oBoundedBeam.SectionType = "BUTL2")) And _
                    (dBoundingHeight = dBoundedHeight) Then
                    If 1.588249619 >= alpha And alpha >= 1.553343034 Then
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
                                pSLH.Add "WebCutI_C3Left"
                            Else
                                pSLH.Add "WebCut_W1Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.SectionType = "B") Then
                                pSLH.Add "WebCut_C3RBulb"
                            ElseIf (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
                                pSLH.Add "WebCutI_C3Right"
                            Else
                                pSLH.Add "WebCut_W1Right"
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
                            pSLH.Add "WebCut_C1Spline"
                        End If
                    ElseIf alpha < 1.553343034 Then
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
                                pSLH.Add "WebCutI_C3Left"
                            Else
                                pSLH.Add "WebCut_C3Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.SectionType = "B") Then
                                pSLH.Add "WebCut_C3RBulb"
                            ElseIf (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
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
                            pSLH.Add "WebCut_C1Spline"
                        End If
                    ElseIf alpha > 1.588249619 Then
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
                                pSLH.Add "WebCutI_C3Left"
                            Else
                                pSLH.Add "WebCut_C3Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.SectionType = "B") Then
                                pSLH.Add "WebCut_C3RBulb"
                            ElseIf (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
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
                            pSLH.Add "WebCut_C1Spline"
                        End If
                    End If
                Else
                    If (dBoundingHeight - dFlangeThickness - Dist) >= dBoundedHeight Then
                        pSLH.Add "WebCut_C1Spline"
                    Else
                        If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                            If (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
                                pSLH.Add "WebCutI_C3Left"
                            Else
                                pSLH.Add "WebCut_C3Left"
                            End If
                        ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                            dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                            If (oSDO_Bounding.SectionType = "B") Then
                                pSLH.Add "WebCut_C3RBulb"
                            ElseIf (oSDO_Bounding.SectionType = "I") Or _
                                (oSDO_Bounding.SectionType = "ISType") Then
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
                            pSLH.Add "WebCut_C1Spline"
                        End If
                    End If
                End If
            
            Case gsC
                If (dBoundingHeight - dFlangeThickness - Dist) >= dBoundedHeight Then
                        pSLH.Add "WebCut_C1Spline"
                Else
                    If dPortValue = JXSEC_WEB_LEFT Or _
                            dPortValue = JXSEC_TOP_FLANGE_LEFT Or _
                            dPortValue = JXSEC_BOTTOM_FLANGE_LEFT Then
                        If (oSDO_Bounding.SectionType = "I") Or _
                            (oSDO_Bounding.SectionType = "ISType") Then
                                pSLH.Add "WebCutI_C3Left"
                        Else
                            pSLH.Add "WebCut_C3Left"
                        End If
                    ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                        dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                        dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Then
                        If (oSDO_Bounding.SectionType = "B") Then
                                pSLH.Add "WebCut_C3RBulb"
                        ElseIf (oSDO_Bounding.SectionType = "I") Or _
                              (oSDO_Bounding.SectionType = "ISType") Then
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
                            pSLH.Add "WebCut_C1Spline"
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
        Select Case strEndCutType
            Case gsW, gsC
                If strWeldPartNumber = "First" Then
                    pSLH.Add "WebCut_C1PCEnd"
zMsgBox "EndToEndWebCutSel::WebCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add WebCut_C1PCEnd"
        
                Else
                    pSLH.Add "WebCut_C1End"
zMsgBox "EndToEndWebCutSel::WebCutNoFlangeEndToEnd" & vbCrLf & _
        "   pSL.Add WebCut_C1End"
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
        Case gsW, gsC
        
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
        If (oSDO_Bounding.SectionType = "B") Or _
            (oSDO_Bounding.SectionType = "UA") Or _
            (oSDO_Bounding.SectionType = "EA") Or _
            (oSDO_Bounding.SectionType = "BUT") Or _
            (oSDO_Bounding.SectionType = "BUTL2") Or _
            (oSDO_Bounding.SectionType = "TSType") Or _
            (oSDO_Bounding.SectionType = "T_XType") Or _
            (oSDO_Bounding.SectionType = "BUTL3") Then
                       
            '**********************************************************************************
            ' Measurement Symbol for getting angle of the bounding profile
            '**********************************************************************************
            'dim the base plate, which will be input to the symbol file
            Dim oBasePlate As IJPlate
            Dim bIsSystem As Boolean
            oSDO_Bounded.GetStiffenedPlate oBasePlate, bIsSystem
        
            'Get web cut object
            Dim oWebCut As Structdetailobjects.WebCut
            Set oWebCut = New WebCut
            Set oWebCut.object = pSLH.SmartOccurrence
        
            Dim MSSymbol As New Structdetailobjects.Measurement
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
            Select Case oSDO_Bounding.SectionType
                Case "B"
                    ssymbolname = "Measurement\CrossSectionB.sym"
                    'Set the input parameter for the symbol file
                    MSSymbol.AddInputParameter "WebOffset", 0.02
                Case "UA", "EA"
                    ssymbolname = "Measurement\CrossSectionUA.sym"
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
        Dim oSDO_Profile As Structdetailobjects.ProfilePart
        Set oSDO_Profile = New Structdetailobjects.ProfilePart
        Set oSDO_Profile.object = oProfilePartObject
        
        sSectionType = oSDO_Profile.SectionType
        
        Set oSDO_Profile = Nothing
        Set oSDO_Profile = Nothing
        
    ElseIf TypeOf oProfilePartObject Is IJBeam Then
        Dim oSDO_Beam As Structdetailobjects.BeamPart
        Set oSDO_Beam = New Structdetailobjects.BeamPart
        Set oSDO_Beam.object = oProfilePartObject
        
        sSectionType = oSDO_Beam.SectionType
        
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
    
    Dim oSDO_Helper As Structdetailobjects.Helper
    Set oSDO_Helper = New Structdetailobjects.Helper
     
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
        Dim oSDO_ProfilePart As Structdetailobjects.ProfilePart
        Set oSDO_ProfilePart = New Structdetailobjects.ProfilePart
        Set oSDO_ProfilePart.object = oBoundedConnectable
        If bUseFlangeData Then
            dCuttingDepth = oSDO_ProfilePart.Width
            If dCuttingDepth < 0.001 Then
                dCuttingDepth = oSDO_ProfilePart.WebThickness
            End If
        Else
            dCuttingDepth = oSDO_ProfilePart.WebThickness
        End If
        Set oSDO_ProfilePart = Nothing
    
    ElseIf sObjectType = SDOBJECT_BEAM Then
        Dim oSDO_BoundedBeam As Structdetailobjects.BeamPart
        Set oSDO_BoundedBeam = New Structdetailobjects.BeamPart
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
    
    Dim oSDO_Helper As Structdetailobjects.Helper
    Dim oSDO_BeamPart As Structdetailobjects.BeamPart
    Dim oSDO_ProfilePart As Structdetailobjects.ProfilePart
    Dim oSDO_EdgeReinforcement As Structdetailobjects.EdgeReinforcement
    
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
    Set oSDO_Helper = New Structdetailobjects.Helper
    oSDO_Helper.GetObjectTypeData oBoundingStiffener, sTypeObject, sObjectType
    If LCase(Trim(sTypeObject)) <> LCase("EdgeReinforcement") Then
    
        If InStr(LCase(Trim(sTypeObject)), LCase("Beam")) > 0 Then
            ' Retrieve the Bounding Beam's Height
            Set oSDO_BeamPart = New Structdetailobjects.BeamPart
            Set oSDO_BeamPart.object = oBoundingStiffener
            dBoundingHeight = oSDO_BeamPart.Height
        
        ElseIf InStr(LCase(Trim(sTypeObject)), LCase("Stiffener")) > 0 Then
            ' Retrieve the Bounding Stiffener's Height
            Set oSDO_ProfilePart = New Structdetailobjects.ProfilePart
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
    Set oSDO_EdgeReinforcement = New Structdetailobjects.EdgeReinforcement
    Set oSDO_EdgeReinforcement.object = oBoundingStiffener
    
    If LCase(Trim(oSDO_EdgeReinforcement.ReinforcementPosition)) = LCase("OnFace") Then
        ' Edge Reinforcement orientation is "OnFace"
        ' use the Edge Reinforcement's Height
        Set oSDO_ProfilePart = New Structdetailobjects.ProfilePart
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
    
    Set oSDO_ProfilePart = New Structdetailobjects.ProfilePart
    Set oSDO_ProfilePart.object = oBoundedProfile
    Set oPort = oSDO_ProfilePart.MountingFacePort
    oTopologyLocate.GetProjectedPointOnModelBody oPort.Geometry, oEndPosition, _
                                                 oPoint_MountingFace, _
                                                 oNormal_MountingFace
    Set oPort = Nothing
    Set oSDO_ProfilePart = Nothing
    
    Set oSDO_ProfilePart = New Structdetailobjects.ProfilePart
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
        
        Dim oSDO_ProfilePart1 As New Structdetailobjects.ProfilePart
        Dim oSDO_ProfilePart2 As New Structdetailobjects.ProfilePart
        
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
    
    Dim oSDO_WebCut As Structdetailobjects.WebCut
    Dim oSDO_FlangeCut As Structdetailobjects.FlangeCut
    
    IsProfileBoundedByPlateEdge = False
    
    ' Verify given SmartItem is a Web Cut or Flange Cut Feature
    Set oObject = pMemberDescription.CAO
    If TypeOf oObject Is IJStructFeature Then
        Set oStructFeature = oObject
        If oStructFeature.get_StructFeatureType = SF_WebCut Then
            sError = "Retreiving Bounded/Bounding objects from WebCut"
            Set oSDO_WebCut = New Structdetailobjects.WebCut
            Set oSDO_WebCut.object = oObject
            
            Set oBoundedPart = oSDO_WebCut.Bounded
            Set oBoundedPort = oSDO_WebCut.BoundedPort
            
            Set oBoundingPart = oSDO_WebCut.Bounding
            Set oBoundingPort = oSDO_WebCut.BoundingPort
            
        ElseIf oStructFeature.get_StructFeatureType = SF_FlangeCut Then
            sError = "Retreiving Bounded/Bounding objects from FlangeCut"
            Set oSDO_FlangeCut = New Structdetailobjects.FlangeCut
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
