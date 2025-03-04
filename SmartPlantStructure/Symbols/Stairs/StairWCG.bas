Attribute VB_Name = "StairWCG"
Option Explicit
'*******************************************************************
'  Copyright (C) 1998, Intergraph Corporation.  All rights reserved.
'
'  File:  GenericWCG.bas
'
'  Description: Weight and Center of Gravity calculator.  This uses the output objects
'       from the symbol to make the calculations.  This assumes that all out put objects
'       are projections, although that is not necessary.
'       The Paint Area is also calculated.  Paint area does not take into account the end of a projection. (capping surface)
'       The Weight, CG, and Paint Area are all slightly off due to the fact that the steps protrude through the side frame.  (vertical ladder)
'
'  History:
'           03/08/2000 - Jeff Machamer - Creation
'           03/21/2000 - Jeff Machamer - Added calculation of Paint Area (Surface Area)
'           01/30/2000 - Aniket Patil - Rewrite for weight and CG.
'           08/24/2005 - J Schwartz - Stair SideFrame_SPSSectionRefStandard attribute is limited to 30 characters since the attribute could be
'                                       used as the name of an Oracle view column name. Chopped one character out (last "a") --> SideFrame_SPSSectionRefStandrd
'           10/08/2009 - GG - DM#172947: Corrected many calculations for weight and CG
'
'******************************************************************
Private Const E_FAIL = -2147467259
Const PI As Double = 3.14159265358979

Public Sub CalcWCG(oStair As ISPSStair, ByVal PartInfoCol As IJDInfosCol, _
                    ByRef weight As Double, _
                    ByRef COGX As Double, _
                    ByRef COGY As Double, _
                    ByRef COGZ As Double)
Const METHOD = "CalcWCG"
On Error GoTo ErrHandler
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    
    Dim i As Integer
    Dim TempWeight As Double, Length As Double, Width As Double, Angle As Double
    Dim AccumWeight As Double, Height As Double, StepPitch As Double
    Dim HandrailPostPitch As Double, PostHt As Double
    Dim NumSteps As Long, NumMidRails As Long
    Dim strMaterial As String, strGrade As String, density As Double
    Dim iMaterial As IJDMaterial
    Dim SideFrameSection As String, SideFrameReferenceStandard As String
    Dim HandRailSection As String, HandRailReferenceStandard As String
    Dim StepSection As String, StepReferenceStandard As String
    Dim StepCSArea As Double, SideFrameCSArea As Double, HandRailCSArea As Double
    Dim WithTopLanding As Boolean
    Dim TopLandingLength As Double
    Dim PlatformThickness As Double
    Dim PlatformSideWidthZ As Double
    Dim SideWidthY As Double
    Dim SideWidthZ As Double
    Dim SideFrameSectionCP As Integer
    Dim Depth As Double
    Dim StepDepth As Double
    Dim NumPosts As Integer
    Dim HandRailSpacing As Double
   
    Dim OccAttrs As IJDAttributes
    Dim PartAttrs As IJDAttributes
    
    Set OccAttrs = oStair
    Angle = GetAttribute1(OccAttrs, "Angle", PartInfoCol)
    Height = GetAttribute1(OccAttrs, "Height", PartInfoCol)
    Width = GetAttribute1(OccAttrs, "Width", PartInfoCol)
    NumSteps = GetAttribute1(OccAttrs, "NumSteps", PartInfoCol)
    StepPitch = GetAttribute1(OccAttrs, "StepPitch", PartInfoCol)
    StepSection = GetAttribute1(OccAttrs, "Step_SPSSectionName", PartInfoCol)
    StepReferenceStandard = GetAttribute1(OccAttrs, "Step_SPSSectionRefStandard", PartInfoCol)
    SideFrameSection = GetAttribute1(OccAttrs, "SideFrm_SPSSectionName", PartInfoCol)
    SideFrameReferenceStandard = GetAttribute1(OccAttrs, "SideFrm_SPSSectionRefStandard", PartInfoCol)
    HandRailSection = GetAttribute1(OccAttrs, "HandRail_SPSSectionName", PartInfoCol)
    HandRailReferenceStandard = GetAttribute1(OccAttrs, "HandRail_SPSSectionRefStandard", PartInfoCol)
    Length = GetAttribute1(OccAttrs, "Length", PartInfoCol)
    NumMidRails = GetAttribute1(OccAttrs, "NumMidRails", PartInfoCol)
    HandrailPostPitch = GetAttribute1(OccAttrs, "HandrailPostPitch", PartInfoCol)
    PostHt = GetAttribute1(OccAttrs, "PostHeight", PartInfoCol)
    strMaterial = GetAttribute1(OccAttrs, "Primary_SPSMaterial", PartInfoCol)
    strGrade = GetAttribute1(OccAttrs, "Primary_SPSGrade", PartInfoCol)
    WithTopLanding = GetAttribute1(OccAttrs, "WithTopLanding", PartInfoCol)
    TopLandingLength = GetAttribute1(OccAttrs, "TopLandingLength", PartInfoCol)
    PlatformThickness = GetAttribute1(OccAttrs, "PlatformThickness", PartInfoCol)
    SideFrameSectionCP = GetAttribute1(OccAttrs, "SideFrameSectionCP", PartInfoCol)
    PlatformSideWidthZ = GetCSAttribData(SideFrameSection, SideFrameReferenceStandard, "ISTRUCTCrossSectionDimensions", "Depth")
    StepDepth = GetCSAttribData(StepSection, StepReferenceStandard, "ISTRUCTCrossSectionDimensions", "Width")
    'only support step section CP5 as default CP
    StepDepth = StepDepth / 2
    
    'Right now, only default CP (CP7) is supported for SideFrameSectionCP.
    
    Set iMaterial = GetMaterialObject(strMaterial, strGrade)
    
    If Not iMaterial Is Nothing Then
        density = iMaterial.density
    Else
        GoTo ErrHandler
    End If
     
    SideFrameCSArea = GetCSAttribData(SideFrameSection, SideFrameReferenceStandard, "IStructCrossSectionDimensions", "Area")
    StepCSArea = GetCSAttribData(StepSection, StepReferenceStandard, "IStructCrossSectionDimensions", "Area")
    HandRailCSArea = GetCSAttribData(HandRailSection, HandRailReferenceStandard, "IStructCrossSectionDimensions", "Area")
    
    NumPosts = Int(((Length - 0.6) / HandrailPostPitch) + 1)
    
    HandRailSpacing = PostHt / (NumMidRails + 1)
        
    'Calculation of CG
    Dim AccumVol As Double, ProjVol As Double, Pos As Double
    Dim AccumCogPos As New DPosition
    Dim CogPos As New DPosition
    Dim ULC As New DPosition   'Upper Left
    Dim URC As New DPosition   'Upper Right
    Dim LLC As New DPosition   'Lower Left
    Dim LRC As New DPosition   'Lower Left
    Dim Justify As Long
    Justify = GetAttribute1(OccAttrs, "Justification", PartInfoCol)
    Angle = (PI / 2 - Angle)
    
    SideWidthY = PlatformSideWidthZ / Cos(Angle)
    SideWidthZ = PlatformSideWidthZ / Sin(Angle)
    
    If Justify = 2 Then
        Pos = -(Width / 2)
    ElseIf Justify = 3 Then
        Pos = (Width / 2)
    End If
   
    ULC.x = -Width / 2 '- Pos
    ULC.y = 0#
    ULC.z = 0#
    
    URC.x = Width / 2 '- Pos
    URC.y = 0#
    URC.z = 0#
    
    LLC.x = -Width / 2 '- Pos
    LLC.y = -Tan(Angle) * Height
    LLC.z = -Height
    
    LRC.x = Width / 2 '- Pos
    LRC.y = -Tan(Angle) * Height
    LRC.z = -Height
    
    'end treatment (left + right)
    If NumMidRails >= 1 Then
        CogPos.x = 0
        CogPos.y = 0
        CogPos.z = ULC.z + (PostHt + HandRailSpacing) / 2
    
        ProjVol = (PostHt - HandRailSpacing) * HandRailCSArea * 2
        AccumVol = AccumVol + ProjVol

        If AccumVol > 0 Then
            AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
            AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
            AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
        End If
    End If
    
    If WithTopLanding Then
        'PlatForm
        CogPos.x = 0
        CogPos.y = -TopLandingLength / 2
        CogPos.z = ULC.z - PlatformThickness / 2
        
        ProjVol = TopLandingLength * Width * PlatformThickness
        AccumVol = AccumVol + ProjVol
        
        If AccumVol > 0 Then
            AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
            AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
            AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
        End If
        
        'Middle Top landing
        If TopLandingLength > 1.5 Then
            CogPos.x = 0
            CogPos.y = -TopLandingLength / 2
            CogPos.z = ULC.z - PlatformSideWidthZ / 2 'CP7
            
            ProjVol = Sqr(TopLandingLength * TopLandingLength + Width * Width) * SideFrameCSArea
            AccumVol = AccumVol + ProjVol
            
            If AccumVol > 0 Then
                AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
                AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
                AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
            End If
        End If
        
        'Left + right Side Top landing
        CogPos.x = 0
        CogPos.y = -TopLandingLength / 2
        CogPos.z = ULC.z - PlatformSideWidthZ / 2 'CP7
        
        ProjVol = TopLandingLength * SideFrameCSArea * 2
        AccumVol = AccumVol + ProjVol
        
        If AccumVol > 0 Then
            AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
            AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
            AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
        End If
                
        'Front + Back landing
        CogPos.x = 0
        CogPos.y = -TopLandingLength / 2
        CogPos.z = -PlatformSideWidthZ / 2  'CP7
        
        ProjVol = Width * SideFrameCSArea * 2
        AccumVol = AccumVol + ProjVol
        
        If AccumVol > 0 Then
            AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
            AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
            AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
        End If
        
'        '4 Posts
        CogPos.x = 0
        CogPos.y = -TopLandingLength / 2
        CogPos.z = PostHt / 2
'
        ProjVol = PostHt * HandRailCSArea * 4
        AccumVol = AccumVol + ProjVol

        If AccumVol > 0 Then
            AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
            AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
            AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
        End If
   
        'Top Handrails on Toplanding (Left + Right)
        CogPos.x = 0
        CogPos.y = -TopLandingLength / 2
        CogPos.z = ULC.z + PostHt
        
        ProjVol = TopLandingLength * HandRailCSArea * 2
        AccumVol = AccumVol + ProjVol
        
        If AccumVol > 0 Then
            AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
            AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
            AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
        End If
        
        'Bottom Handrail on Toplanding (left + right)
        If NumMidRails >= 1 Then
            CogPos.x = 0
            CogPos.y = -TopLandingLength / 2
            CogPos.z = ULC.z + PostHt - (HandRailSpacing * (NumMidRails))
        
            ProjVol = TopLandingLength * HandRailCSArea * 2
            AccumVol = AccumVol + ProjVol

            If AccumVol > 0 Then
                AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
                AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
                AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
            End If
        End If
    
        'MiddleRails on Toplanding (left + right)
        If NumMidRails >= 2 Then
            For i = 2 To NumMidRails
                CogPos.x = 0
                CogPos.y = -TopLandingLength / 2
                CogPos.z = ULC.z + PostHt - (HandRailSpacing * (i - 1))
                
                ProjVol = TopLandingLength * HandRailCSArea * 2
                AccumVol = AccumVol + ProjVol
            
                If AccumVol > 0 Then
                    AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
                    AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
                    AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
                End If
            Next i
        End If
               
        LLC.y = LLC.y - TopLandingLength
        LRC.y = LRC.y - TopLandingLength
        ULC.y = ULC.y - TopLandingLength
        URC.y = URC.y - TopLandingLength
    End If
 
'   CG of Left + Right stringer
    CogPos.x = 0
    CogPos.y = (LLC.y + ULC.y) / 2 + SideWidthY / 2 'CP7
    CogPos.z = (LLC.z + ULC.z) / 2 - SideWidthZ / 2 'CP7
           
    ProjVol = Length * SideFrameCSArea * 2
    AccumVol = AccumVol + ProjVol
    
    If AccumVol > 0 Then
        AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
        AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
        AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
    End If
     
'    CG of Steps
    Dim StPt As New DPosition
    Dim endpt As New DPosition
    Set StPt = ULC.Clone
    Set endpt = URC.Clone
    Dim TmpDepth As Double
    Dim CGx As Double, CGy As Double, CGz As Double
    'Dim xp As Double, yp As Double
    'Dim DepthCorrection As Double
    'TmpDepth = (Depth / 2) / Sin(Angle)
    'DepthCorrection = 0
    'Right now, only default CP (CP7) is supported. Comment out the following code
'    Select Case SideFrameSectionCP
'            Case 7, 8, 9, 14  'top edge
'                DepthCorrection = 0
'            Case 4, 5, 6  'along half depth
'                DepthCorrection = (Depth / 2) / Sin(Angle)
'            Case 1, 2, 3, 11 'bottom  edge
'                DepthCorrection = Depth / Sin(Angle)
'            Case 10, 12, 13  'along centroid in depth direction
'                CGy = GetCSAttribData(SideFrameSection, SideFrameReferenceStandard, "ISTRUCTCrossSectionDesignProperties", "CentroidY")
'                DepthCorrection = CGy / Sin(Angle)
'            Case 15  'shear center
'                yp = GetCSAttribData(SideFrameSection, SideFrameReferenceStandard, "IJUAL", "yp")
'                DepthCorrection = yp / Sin(Angle)
'        End Select
        
    'TmpDepth = TmpDepth - DepthCorrection
    
    For i = 1 To NumSteps
                
        CogPos.x = 0
        CogPos.y = URC.y - (Tan(Angle)) * (((StepPitch * i) - SideWidthY / 2 + StepDepth))
        CogPos.z = URC.z - (i * StepPitch + StepDepth)
         
        ProjVol = Width * StepCSArea
        AccumVol = AccumVol + ProjVol
        
        If AccumVol > 0 Then
            AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
            AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
            AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
        End If

    Next i
    
'   Calculate CG for posts
    Set StPt = ULC.Clone
     
    Dim NewSpacing As Double
    NewSpacing = Length / NumPosts
    
    StPt.z = StPt.z - Cos(Angle) * 0.3
    StPt.y = StPt.y - Sin(Angle) * 0.3
    
    For i = 1 To NumPosts
        If i <> 1 Then
            StPt.y = StPt.y - (Sin(Angle) * NewSpacing)
            StPt.z = StPt.z - (Cos(Angle) * NewSpacing)
        End If
        
        CogPos.x = 0
        CogPos.y = StPt.y
        CogPos.z = StPt.z + PostHt / 2
        
        ProjVol = PostHt * HandRailCSArea * 2
        AccumVol = AccumVol + ProjVol
        
        If AccumVol > 0 Then
            AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
            AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
            AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
        End If
        
    Next i
    
        
'   calculate for CG of Handrails
    Set StPt = ULC.Clone
    Set endpt = LLC.Clone
    
'   Top rail (left + right)
    CogPos.x = 0
    CogPos.y = (endpt.y + StPt.y) / 2
    CogPos.z = (endpt.z + StPt.z) / 2 + PostHt
    
    ProjVol = Length * HandRailCSArea * 2
    AccumVol = AccumVol + ProjVol

    If AccumVol > 0 Then
        AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
        AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
        AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
    End If

'   Bottom rails (left + right)
    If NumMidRails >= 1 Then
        CogPos.x = 0
        CogPos.y = (endpt.y + StPt.y) / 2
        CogPos.z = (endpt.z + StPt.z) / 2 + PostHt - HandRailSpacing * NumMidRails
        
        ProjVol = Length * HandRailCSArea * 2
        AccumVol = AccumVol + ProjVol

        If AccumVol > 0 Then
            AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
            AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
            AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
        End If
    End If
    
'   MiddleRails (left + right)
    If NumMidRails >= 2 Then
        For i = 2 To NumMidRails
            CogPos.x = 0
            CogPos.y = (endpt.y + StPt.y) / 2
            CogPos.z = (endpt.z + StPt.z) / 2 + PostHt - (HandRailSpacing * (i - 1))
            
            ProjVol = Length * HandRailCSArea * 2
            AccumVol = AccumVol + ProjVol
            
            If AccumVol > 0 Then
                AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
                AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
                AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
            End If
        Next i
    End If
                  
    'end treatment (left + right)
    If NumMidRails >= 1 Then
        CogPos.x = 0
        CogPos.y = LLC.y
        CogPos.z = LLC.z + (PostHt + HandRailSpacing) / 2
    
        ProjVol = (PostHt - HandRailSpacing) * HandRailCSArea * 2
        AccumVol = AccumVol + ProjVol

        If AccumVol > 0 Then
            AccumCogPos.x = AccumCogPos.x + (CogPos.x - AccumCogPos.x) * ProjVol / AccumVol
            AccumCogPos.y = AccumCogPos.y + (CogPos.y - AccumCogPos.y) * ProjVol / AccumVol
            AccumCogPos.z = AccumCogPos.z + (CogPos.z - AccumCogPos.z) * ProjVol / AccumVol
        End If
    End If
    COGX = AccumCogPos.x - Pos
    COGY = AccumCogPos.y
    COGZ = AccumCogPos.z
    weight = AccumVol * density
    
    Set OccAttrs = Nothing
Exit Sub
ErrHandler:
    Err.Raise E_FAIL
    oErrors.Add Err.Number, METHOD, Err.Description
End Sub
