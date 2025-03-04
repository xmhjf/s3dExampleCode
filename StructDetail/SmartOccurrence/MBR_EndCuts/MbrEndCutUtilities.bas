Attribute VB_Name = "MbrEndCutUtilities"
'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'File : MbrEndCutUtilities.bas
'
'Author : D.A. Trent
'
'Description :
'   Utilites for determining Type Of EndCuts to be Placed on Member to Member Assembly Connections
'
'History:
'*****************************************************************************
Option Explicit
Private Const MODULE = "StructDetail\Data\SmartOccurrence\Mbr_EndCuts\MbrEndCutUtilities"
'
Public Const C_Port_Base = "Base"
Public Const C_Port_Offset = "Offset"
Public Const C_Port_Lateral = "Lateral"

Public Const C_Port_Top = "Top"
Public Const C_Port_Bottom = "Bottom"
Public Const C_Port_WebLeft = "WebLeft"
Public Const C_Port_WebRight = "WebRight"
'

'
'*************************************************************************
'Function
'Parent_SmartItemName
'
'Abstract
'   Given the Smart Occurent Object
'    return the Owning Smart item Name
'
'input
'   oOccurrenceObject
'
'Return
'   sParentItemName
'   oParentObject
'
'Exceptions
'
'***************************************************************************
Public Sub Parent_SmartItemName(oOccurrenceObject As Object, _
                                sParentItemName As String, _
                                Optional oParentObject As Object = Nothing)
Const METHOD = "MbrEndCutUtilities::Parent_SmartItemName"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim nCount As Long
    Dim oColObject As Object
    Dim oRelationShip As IJDRelationship
    Dim oAssocRelation As IJDAssocRelation
    Dim oRelationShipCol As IJDRelationshipCol
    Dim oParent_SmartOccurrence As IJSmartOccurrence
    
    sParentItemName = ""
    If Not TypeOf oOccurrenceObject Is IJSmartOccurrence Then
        Exit Sub
    End If
                
    ' Need to check if given SmartOccurrence
    ' is a member of a CustomAssembly relationship
    If TypeOf oOccurrenceObject Is IJDAssocRelation Then
        Set oAssocRelation = oOccurrenceObject
        Set oColObject = oAssocRelation.CollectionRelations("IJFullObject", _
                                                            "toAssembly")
        If oColObject Is Nothing Then
        
        ElseIf TypeOf oColObject Is IJDRelationshipCol Then
            Set oRelationShipCol = oColObject
            nCount = oRelationShipCol.Count
                          
            If nCount > 0 Then
                Set oRelationShip = oRelationShipCol.Item(1)
                Set oParentObject = oRelationShip.Target
            End If
        
        End If
    End If
    
    If Not oParentObject Is Nothing Then
        If TypeOf oParentObject Is IJSmartOccurrence Then
            Set oParent_SmartOccurrence = oParentObject
            sParentItemName = oParent_SmartOccurrence.Item
        End If
    Else
        Set oParent_SmartOccurrence = oOccurrenceObject
        sParentItemName = oParent_SmartOccurrence.Item
    End If
    
    Exit Sub

    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'Parent_WebTypeCase
'
'Abstract
'   Given the Smart Occurent Object
'    return the WebCut Type from the Owning Smart Item Name
'
'input
'   oOccurrenceObject
'
'Return
'   sWebTypeCase
'
'Exceptions
'
'Notes:
' Expect the Owning Smart Item Name to be in the following format:
'   sId_sWebType_sFlangeType
' Returns the text between the first "_" and the second "_"
'
'***************************************************************************
Public Sub Parent_WebTypeCase(oEndCutObject As Object, _
                              sWebTypeCase As String)
Const METHOD = "MbrEndCutUtilities::Parent_WebTypeCase"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sTemp As String
    Dim sTrim As String
    Dim sEndCutItem As String
    
    Dim iLoc As Long
    
    Dim oParentObject As Object
    
    ' Retreive the Owning AssemblyConnection
    sWebTypeCase = ""
    Parent_SmartItemName oEndCutObject, sEndCutItem, oParentObject
    
    sTrim = Trim(sEndCutItem)
    If Len(sTrim) < 1 Then
        Exit Sub
    End If
    
    iLoc = InStr(sTrim, "_")
    If iLoc > 0 Then
        sTemp = Mid(sTrim, iLoc + 1)
        sTrim = Trim(sTemp)
        iLoc = InStr(sTrim, "_")
        If iLoc > 1 Then
            sTemp = Left(sTrim, iLoc - 1)
            sWebTypeCase = Trim(sTemp)
        Else
            sWebTypeCase = sTrim
        End If
    
    Else
        sWebTypeCase = sTrim
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'Parent_FlangeTypeCase
'
'Abstract
'   Given the Smart Occurent Object
'    return the FlangeCut Type from the Owning Smart Item Name
'
'input
'   oOccurrenceObject
'
'Return
'   sWebTypeCase
'
'Exceptions
'
'Notes:
' Expect the Owning Smart Item Name to be in the following format:
'   sId_sWebType_sFlangeType
' Returns the text after the second "_"
'
'***************************************************************************
Public Sub Parent_FlangeTypeCase(oEndCutObject As Object, _
                                 sFlangeTypeCase As String)
Const METHOD = "MbrEndCutUtilities::Parent_FlangeTypeCase"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sTemp As String
    Dim sTrim As String
    Dim sEndCutItem As String
    
    Dim iLoc As Long
    
    Dim oParentObject As Object
    
    ' Retreive the Owning AssemblyConnection
    sFlangeTypeCase = ""
    Parent_SmartItemName oEndCutObject, sEndCutItem, oParentObject
    
    sTrim = Trim(sEndCutItem)
    If Len(sTrim) < 1 Then
        Exit Sub
    End If
    
    iLoc = InStr(sTrim, "_")
    If iLoc > 0 Then
        sTemp = Mid(sTrim, iLoc + 1)
        sTrim = Trim(sTemp)
        iLoc = InStr(sTrim, "_")
        If iLoc > 0 Then
            sTemp = Mid(sTrim, iLoc + 1)
            sFlangeTypeCase = Trim(sTemp)
        Else
            sFlangeTypeCase = sTrim
        End If
    Else
        sFlangeTypeCase = sTrim
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'WebType_Select
'
'Abstract
'   Based on the Bounded and Bounding Data:
'       Cross Section Types, Idealized Boundary, Orientation
'   Determine the actual Symbol File used to create the Web Cut
'
'input
'   pSelectorLogic
'   oBoundedData
'   oBoundingData
'   sWebTypeCase
'
'Return
'
'Exceptions
'
'Notes:
'   Symbol Files for Bounding Tubluar Cross Section
'   Symbol Files for Bounded Tubluar Cross Section (End to End case)
'   Symbol Files for Bounding with Top/Bottom Left Flange sections
'   Symbol Files for Bounding with Top/Bottom Right Flange sections
'
'***************************************************************************
Public Sub WebType_Select(pSelectorLogic As IJDSelectorLogic, _
                          oBoundedData As MemberConnectionData, _
                          oBoundingData As MemberConnectionData, _
                          sWebTypeCase As String, _
                          bEndToEndCase As Boolean)
Const METHOD = "MbrEndCutUtilities::WebType_Select"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim lStatus As Long
    
    Dim sConfig As String
    Dim sIdealizedBounded As String
    Dim sIdealizedBoundary As String
    
    Dim bFlanges As Boolean
    Dim bTopFlange As Boolean
    Dim bBottomFlange As Boolean
    
    Dim oEndCutObject As Object
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    
    ' Retreive/Calculate/Determine the Idealized Boundary
    sMsg = "Unknown Error"
    CheckIdealizedBoundary oBoundedData, oBoundingData, sIdealizedBoundary
    CheckIdealizedBoundary oBoundingData, oBoundedData, sIdealizedBounded
    
    ' Determine Configuration between Bounded and Bounding Member
    EndCut_WebFlangeConfig oBoundedData, oBoundingData, sConfig
    
    ' Base on the Orientation between the Bounding and Bounded Members
    ' Determine if Bounding Member Top and Bottom Flanges are to be considered
    bFlanges = False
    bTopFlange = False
    bBottomFlange = False
    
    If Trim(LCase(sConfig)) = LCase("Top_Top") Or _
       Trim(LCase(sConfig)) = LCase("Bottom_Top") Then
        EndCut_IdealizedWebFlanges oBoundingData.MemberPart, _
                                   sIdealizedBoundary, _
                                   bTopFlange, bBottomFlange
        If Trim(LCase(sConfig)) = LCase("Bottom_Top") Then
            bFlanges = bTopFlange
            bTopFlange = bBottomFlange
            bBottomFlange = bFlanges
        End If
    
        ' Set flag indicating that both Top and Bottom Flange sections exist
        bFlanges = False
        If bTopFlange Then
            If bBottomFlange Then
                bFlanges = True
            End If
        End If
    End If
    
    ' If Bounding Cross Section is Tubular or
    ' if Bounded Cross Section is Tubular
    If Trim(LCase(sIdealizedBoundary)) = LCase(eIdealized_BoundingTube) Then
        ' Bounding Cross Section is Tubular
        If bEndToEndCase Then
            ' Bounding Cross Section is Tubular (use Idealized Boundary)
            If Trim(LCase(sIdealizedBounded)) = LCase(eIdealized_BoundingTube) Then
                pSelectorLogic.Add "M_Web_Weld_Tube00"
            Else
                pSelectorLogic.Add "M_Web_Weld_01"
            End If
        Else
            ' Bounding Cross Section is Tubular (single Saddle Type Cut)
            pSelectorLogic.Add "M_Web_Weld_Tube01"
            pSelectorLogic.Add "M_Web_Weld_Tube02"
            pSelectorLogic.Add "M_Web_Weld_Tube03"
        End If
        
        Exit Sub
    
    ElseIf Trim(LCase(sIdealizedBounded)) = LCase(eIdealized_BoundingTube) Then
        ' Bounded Cross Section is Tubular (use Idealized Boundary)
        pSelectorLogic.Add "M_Web_Weld_Tube00"
        Exit Sub
    End If
    
    ' For both Web Top/Bottom are Weld Type (Straight)
    If InStr(LCase(sWebTypeCase), LCase("W1W1")) > 0 Then
        If Trim(LCase(sIdealizedBoundary)) = LCase(eIdealized_Top) Then
            ' Bounding IdealizedBoundary is Top Flange
            pSelectorLogic.Add "M_Web_Weld_01"
        
        ElseIf Trim(LCase(sIdealizedBoundary)) = LCase(eIdealized_Bottom) Then
            ' Bounding IdealizedBoundary is Bottom Flange
            pSelectorLogic.Add "M_Web_Weld_01"
        
        ElseIf Trim(LCase(sIdealizedBoundary)) = LCase(eIdealized_WebLeft) Then
            ' Bounding IdealizedBoundary is Web Left
            ' check if Bounding Member has Left Top/Bottom Flange sections
            pSelectorLogic.Add "M_Web_Weld_01"
            
            If bFlanges Then
                pSelectorLogic.Add "M_Web_Weld_FL"
            ElseIf bTopFlange Then
                pSelectorLogic.Add "M_Web_Weld_FLT"
            ElseIf bBottomFlange Then
                pSelectorLogic.Add "M_Web_Weld_FLB"
            End If
        
        ElseIf Trim(LCase(sIdealizedBoundary)) = LCase(eIdealized_WebRight) Then
            ' Bounding IdealizedBoundary is Web Right
            ' check if Bounding Member has Right Top/Bottom Flange sections
            ' (use same for both Flange Right and Flange Left cases)
            pSelectorLogic.Add "M_Web_Weld_01"
            
            If bFlanges Then
'ToDo                pSelectorLogic.Add "M_Web_Weld_FR"
                pSelectorLogic.Add "M_Web_Weld_FL"
            ElseIf bTopFlange Then
'ToDo                pSelectorLogic.Add "M_Web_Weld_FRT"
                pSelectorLogic.Add "M_Web_Weld_FLT"
            ElseIf bBottomFlange Then
'ToDo                pSelectorLogic.Add "M_Web_Weld_FRB"
                pSelectorLogic.Add "M_Web_Weld_FLB"
            End If
        
        ElseIf Trim(LCase(sIdealizedBoundary)) = LCase(eIdealized_EndBaseFace) Then
            ' Bounding IdealizedBoundary is Base (use Idealized Boundary)
            pSelectorLogic.Add "M_Web_Weld_01"
        
        ElseIf Trim(LCase(sIdealizedBoundary)) = LCase(eIdealized_EndOffsetFace) Then
            ' Bounding IdealizedBoundary is Offset (use Idealized Boundary)
            pSelectorLogic.Add "M_Web_Weld_01"
        
        Else
            pSelectorLogic.Add "M_Web_Weld_01"
        End If
        
    ' For Web Top is Snipe Type (Angled), Web Bottom is Weld Type (Straight)
    ElseIf InStr(LCase(sWebTypeCase), LCase("S1W1")) > 0 Then
        If bTopFlange Then
            pSelectorLogic.Add "M_Web_Weld_ST"
            pSelectorLogic.Add "MP_Web_Weld_ST"
        ElseIf oBoundingData.MemberPart.IsPrismatic Then
            pSelectorLogic.Add "M_Web_Weld_ST1"
            pSelectorLogic.Add "MP_Web_Weld_ST"
        Else
            pSelectorLogic.Add "MP_Web_Weld_ST"
            pSelectorLogic.Add "M_Web_Weld_ST1"
        End If
    
    ' For Web Top is Weld Type (Straight), Web Bottom is Snipe Type (Angled)
    ElseIf InStr(LCase(sWebTypeCase), LCase("W1S1")) > 0 Then
        If bBottomFlange Then
            pSelectorLogic.Add "M_Web_Weld_SB"
            pSelectorLogic.Add "MP_Web_Weld_SB"
        ElseIf oBoundingData.MemberPart.IsPrismatic Then
            pSelectorLogic.Add "M_Web_Weld_SB1"
            pSelectorLogic.Add "MP_Web_Weld_SB"
        Else
            pSelectorLogic.Add "MP_Web_Weld_SB"
            pSelectorLogic.Add "M_Web_Weld_SB1"
        End If
    
    ' For Web Top is Snipe Type (Angled), Web Bottom is Snipe Type (Angled)
    ElseIf InStr(LCase(sWebTypeCase), LCase("S1S1")) > 0 Then
        If bFlanges Then
            pSelectorLogic.Add "M_Web_Weld_SS"
            pSelectorLogic.Add "MP_Web_Weld_SS"
        ElseIf oBoundingData.MemberPart.IsPrismatic Then
            pSelectorLogic.Add "M_Web_Weld_SS1"
            pSelectorLogic.Add "MP_Web_Weld_SS"
        Else
            pSelectorLogic.Add "MP_Web_Weld_SS"
            pSelectorLogic.Add "M_Web_Weld_SS1"
        End If
    
    ' For Web Top is Cope Type (Curved), Web Bottom is Weld Type (Straight)
    ElseIf InStr(LCase(sWebTypeCase), LCase("C1W1")) > 0 Then
        If oBoundingData.MemberPart.IsPrismatic Then
            pSelectorLogic.Add "M_Web_Weld_CT"
            pSelectorLogic.Add "MP_Web_Weld_CT"
        Else
            pSelectorLogic.Add "MP_Web_Weld_CT"
            pSelectorLogic.Add "M_Web_Weld_CT"
        End If
    
    ' For Web Top is Weld Type (Straight), Web Bottom is Cope Type (Curved)
    ElseIf InStr(LCase(sWebTypeCase), LCase("W1C1")) > 0 Then
        If oBoundingData.MemberPart.IsPrismatic Then
            pSelectorLogic.Add "M_Web_Weld_CB"
            pSelectorLogic.Add "MP_Web_Weld_CB"
        Else
            pSelectorLogic.Add "MP_Web_Weld_CB"
            pSelectorLogic.Add "M_Web_Weld_CB"
        End If
    
    ' For Web Top is Cope Type (Curved), Web Bottom is Cope Type (Curved)
    ElseIf InStr(LCase(sWebTypeCase), LCase("C1C1")) > 0 Then
        If oBoundingData.MemberPart.IsPrismatic Then
            pSelectorLogic.Add "M_Web_Weld_CC"
            pSelectorLogic.Add "MP_Web_Weld_CC"
        Else
            pSelectorLogic.Add "MP_Web_Weld_CC"
            pSelectorLogic.Add "M_Web_Weld_CC"
        End If

    ' For Web Top is Snipe Type (Angled), Web Bottom is Cope Type (Curved)
    ElseIf InStr(LCase(sWebTypeCase), LCase("S1C1")) > 0 Then
        pSelectorLogic.Add "MP_Web_Weld_SC"
    
    ' For Web Top is Cope Type (Curved), Web Bottom is Snipe Type (Angled)
    ElseIf InStr(LCase(sWebTypeCase), LCase("C1S1")) > 0 Then
        pSelectorLogic.Add "MP_Web_Weld_CS"
    End If


    Exit Sub
    
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
End Sub

'*************************************************************************
'Function
'FlangeType_Select
'
'Abstract
'   Based on the Bounded and Bounding Data:
'       Cross Section Types, Idealized Boundary, Orientation
'   Determine the actual Symbol File used to create the Flange Cut
'
'input
'   pSelectorLogic
'   oBoundedData
'   oBoundingData
'   sWebTypeCase
'
'Return
'
'Exceptions
'
'Notes:
'
'***************************************************************************
Public Sub FlangeType_Select(pSelectorLogic As IJDSelectorLogic, _
                             oBoundedData As MemberConnectionData, _
                             oBoundingData As MemberConnectionData, _
                             sFlangeTypeCase As String, _
                             bEndToEndCase As Boolean)
Const METHOD = "MbrEndCutUtilities::FlangeType_Select"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim lStatus As Long
    
    Dim sConfig As String
    Dim sIdealizedBounded As String
    Dim sIdealizedBoundary As String
    
    Dim bFlanges As Boolean
    Dim bTopFlange As Boolean
    Dim bBottomFlange As Boolean
    
    Dim oEndCutObject As Object
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    
    ' Retreive/Calculate/Determine the Idealized Boundary
    sMsg = "Unknown Error"
    CheckIdealizedBoundary oBoundedData, oBoundingData, sIdealizedBoundary
    CheckIdealizedBoundary oBoundingData, oBoundedData, sIdealizedBounded
    
    ' Determine Configuration between Bounded and Bounding Member
    EndCut_WebFlangeConfig oBoundedData, oBoundingData, sConfig
    
    ' Base on the Orientation between the Bounding and Bounded Members
    ' Determine if Bounding Member Top and Bottom Flanges are to be considered
    bFlanges = False
    bTopFlange = False
    bBottomFlange = False
    
    If Trim(LCase(sConfig)) = LCase("Left_Top") Or _
       Trim(LCase(sConfig)) = LCase("Right_Top") Then
        EndCut_IdealizedWebFlanges oBoundingData.MemberPart, _
                                   sIdealizedBoundary, _
                                   bTopFlange, bBottomFlange
        If Trim(LCase(sConfig)) = LCase("Bottom_Top") Then
            bFlanges = bTopFlange
            bTopFlange = bBottomFlange
            bBottomFlange = bFlanges
        End If
    
        ' Set flag indicating that both Top and Bottom Flange sections exist
        bFlanges = False
        If bTopFlange Then
            If bBottomFlange Then
                bFlanges = True
            End If
        End If
    End If
    
    ' If Bounding Cross Section is Tubular or
    ' if Bounded Cross Section is Tubular
    ' No Flange Cut can be applied
    If Trim(LCase(sIdealizedBoundary)) = LCase(eIdealized_BoundingTube) Then
        Exit Sub
    
    ElseIf Trim(LCase(sIdealizedBounded)) = LCase(eIdealized_BoundingTube) Then
        Exit Sub
    End If
    
    ' Currently, Always Default to the Weld (Straight) case
    ' Use must manually Select other cases
    pSelectorLogic.Add "M_Flange_Weld_01"
    
    ' For both Flange Left/Right are Weld Type (Straight)
    If InStr(LCase(sFlangeTypeCase), LCase("W1W1")) > 0 Then
        ' already added
        ' pSelectorLogic.Add "M_Flange_Weld_01"
        
    ' Flange Left is Snipe Type (Angled), Flange Right is Weld Type (Straight)
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("S1W1")) > 0 Then
        pSelectorLogic.Add "M_Flange_Weld_SL"
        pSelectorLogic.Add "MP_Flange_Free_SL"
        pSelectorLogic.Add "MP_Flange_Free_SS"
    
    ' Flange Left is Weld Type (Straight), Flange Right is Snipe Type (Angled)
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("W1S1")) > 0 Then
        pSelectorLogic.Add "M_Flange_Weld_SR"
        pSelectorLogic.Add "MP_Flange_Free_SR"
        pSelectorLogic.Add "MP_Flange_Free_SS"
    
    ' Flange Left is Snipe Type (Angled), Flange Right is Snipe Type (Angled)
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("S1S1")) > 0 Then
'ToDo        pSelectorLogic.Add "M_Flange_Weld_SS"
        pSelectorLogic.Add "MP_Flange_Free_SS"
    
    ' Flange Left is Cope Type (Curved), Flange Right is Weld Type (Straight)
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("C1W1")) > 0 Then
'ToDo        pSelectorLogic.Add "M_Flange_Weld_CL"
    
    ' Flange Left is Weld Type (Straight), Flange Right is Cope Type (Curved)
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("W1C1")) > 0 Then
'ToDo        pSelectorLogic.Add "M_Flange_Weld_CR"
    
    ' Flange Left is Cope Type (Curved), Flange Right is Cope Type (Curved)
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("C1C1")) > 0 Then
'ToDo        pSelectorLogic.Add "M_Flange_Weld_CC"

    ' Flange Left is Snipe Type (Angled), Flange Right is Cope Type (Curved)
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("S1C1")) > 0 Then
'ToDo        pSelectorLogic.Add "M_Flange_Weld_SC"
    
    ' Flange Left is Cope Type (Curved), Flange Right is Snipe Type (Angled)
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("C1S1")) > 0 Then
'ToDo        pSelectorLogic.Add "M_Flange_Weld_CS"
    End If


    Exit Sub
    
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
End Sub

'*************************************************************************
'Function
'SetMatlGradeThickness
'
'Abstract
'   Based on the Bounded and Bounding Data:
'       Cross Section Types, Idealized Boundary, Orientation
'   Determine the actual Symbol File used to create the Web Cut
'
'input
'   oStructMaterial
'   strMatl
'   strGrade
'   dThickness
'
'Return
'
'Exceptions
'
'Notes:
'
'***************************************************************************
Public Sub SetMatlGradeThickness(oStructMaterial As IJStructureMaterial, _
                                 strMatl As String, strGrade As String, _
                                 dThickness As Double)
Const METHOD = "MbrEndCutUtilities::SetMatlGradeThickness"
  
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim nCount As Long
    Dim oPlate As IJPlate
    Dim oMatlObj As IJDMaterial
    Dim oPlateDims As IJDPlateDimensions
    Dim oRefDataQuery As RefDataMiddleServices.RefdataSOMMiddleServices
    Dim matlThickCol As IJDCollection
    Dim oMatProxy As Object
    Dim oPlateDimProxy As Object
    Dim oResMgr As IJDPOM
    
    Set oResMgr = GetResourceMgr
    
    Set oRefDataQuery = New RefDataMiddleServices.RefdataSOMMiddleServices
    Set oMatlObj = oRefDataQuery.GetMaterialByGrade(strMatl, strGrade)
    Set oMatProxy = oResMgr.GetProxy(oMatlObj)
    
    oStructMaterial.Material = oMatProxy

    If Not TypeOf oStructMaterial Is IJPlate Then
        Exit Sub
    End If
    
    Set oPlate = oStructMaterial
    Set matlThickCol = oRefDataQuery.GetPlateDimensions(oMatlObj.MaterialType, oMatlObj.MaterialGrade)
  
    nCount = matlThickCol.Size
    For iIndex = 1 To nCount
        Set oPlateDims = matlThickCol.Item(iIndex)
        If Abs(oPlateDims.thickness - dThickness) < 0.000005 Then
            Exit For
        End If
        Set oPlateDims = Nothing
    Next
    
    If oPlateDims Is Nothing Then
        Set oPlateDims = matlThickCol.Item(1)
    End If
    
    Set oPlateDimProxy = oResMgr.GetProxy(oPlateDims)
    oPlate.Dimensions = oPlateDimProxy
    
    Set oPlateDims = Nothing
    Set oPlateDimProxy = Nothing
    Set oRefDataQuery = Nothing
  
  Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'GetParentAnswer
'
'Abstract
'   Gets the Question from from the Smart Occurenence Parent
'
'input
'   oSmartObject
'   sQuestion
'
'Return
'   sAnswer
'
'Exceptions
'
'Notes:
'
'***************************************************************************
Public Sub GetParentAnswer(oSmartObject As Object, sQuestion As String, _
                           sAnswer As String)
Const METHOD = "MbrEndCutUtilities::GetParentAnswer"
    On Error GoTo ErrorHandler
    Dim sMsg As String
     
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
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'SetNamingRule
'
'Abstract
'   Based on the Bounded and Bounding Data:
'       Cross Section Types, Idealized Boundary, Orientation
'   Determine the actual Symbol File used to create the Web Cut
'
'input
'   oBearingPlate
'
'Return
'
'Exceptions
'
'Notes:
'
'***************************************************************************
Public Sub SetBearingPlateNamingRule(oBearingPlate As Object)
Const METHOD = "MbrEndCutUtilities::SetBearingPlateNamingRule"
  
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim idx As Long
    Dim lPriority() As Long
    Dim strLongNames() As String
    Dim strShortNames() As String
    
    Dim oRules As IJElements
    Dim oPlate As IJPlate
    Dim oDummyAE As IJNameRuleAE
    Dim oQueryUtil As IJMetaDataCategoryQuery
    Dim oNamingObject As IJDNamingRulesHelper

    Dim oJDObject As IJDObject
    Dim oResourceManager As IUnknown
    
    'Retrieve first default naming rule
    Set oNamingObject = New NamingRulesHelper
    oNamingObject.GetEntityNamingRulesGivenName "CPlatePart", oRules
    If oRules.Count >= 1 Then
        oNamingObject.AddNamingRelations oBearingPlate, oRules.Item(1), oDummyAE
    End If
    Set oDummyAE = Nothing
    Set oNamingObject = Nothing
    
    ' Default naming category to first non-negative value
    Set oJDObject = oBearingPlate
    Set oResourceManager = oJDObject.ResourceManager

    Set oQueryUtil = New CMetaDataCategoryQuery
    oQueryUtil.GetCategoryInfo oResourceManager, _
                               "BracketCategory", _
                               strLongNames, _
                               strShortNames, _
                               lPriority
    Set oQueryUtil = Nothing

    Set oPlate = oBearingPlate
    oPlate.NamingCategory = -1
    For idx = LBound(lPriority) To UBound(lPriority)
        If lPriority(idx) >= 0 Then
            oPlate.NamingCategory = lPriority(idx)
            Exit For
        End If
    Next idx

    Set oPlate = Nothing
    Erase strLongNames
    Erase strShortNames
    Erase lPriority
  
  Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'***********************************************************************
' METHOD:  Get_PortFaceType
'
' DESCRIPTION:
'       Given an IJStructPort
'       Determine if the Port is 'Base', 'Offset', 'WebLeft', etc.
'
'   inputs:
'       oPort As IJPort
'
'   outputs:
'       GetPortBaseOffsetSide as String
'           "Base"      given Port is "Base" port
'           "Offset"    given Port is "Offset" port
'           "Lateral"   given Port is "Lateral" port
'           "TOP"       given Port is Profile JXSEC_TOP port
'           "BOTTOM"    given Port is Profile JXSEC_BOTTOM port
'           "WEB_LEFT"  given Port is Profile JXSEC_WEB_LEFT port
'           "WEB_RIGHT" given Port is Profile JXSEC_WEB_RIGHT port
'
'***********************************************************************
Public Function Get_PortFaceType(oPortObject As Object) As String
Const METHOD = "Get_PortFaceType"
On Error GoTo ErrorHandler
    
    On Error Resume Next
    Dim sMsg As String
    
    Dim lBaseCheck As Long
    Dim lNplusCheck As Long
    Dim lNminusCheck As Long
    Dim lOffsetCheck As Long
    Dim lLateralCheck As Long
    
    Dim lContextID As Long
    Dim lOperatorID As Long
    Dim lOperatationID As Long
    
    Dim ePortType As JS_TOPOLOGY_PROXY_TYPE
    
    Dim sPortSide As String
    Dim eContextID As GSCADSDCreateModifyUtilities.eUSER_CTX_FLAGS
    
    Dim oPort As IJPort
    Dim oConnectable As IJConnectable
    Dim oStructPort As IMSStructConnection.IJStructPort
    
    Dim eAxisPortIndex As SPSMemberAxisPortIndex
    Dim oSplitAxisPort As ISPSSplitAxisPort
    
    Dim oMemberFactory As SPSMembers.SPSMemberFactory
    Dim oMemberConnectionServices As SPSMembers.ISPSMemberConnectionServices
    
    If TypeOf oPortObject Is IJPort Then
        Set oPort = oPortObject
        Set oConnectable = oPort.Connectable
    End If
    
    
    ' SP3D Member Object Ports do not implement IJStructPort interface
    sPortSide = ""
    If oConnectable Is Nothing Then
    
    ElseIf TypeOf oConnectable Is ISPSMemberPartPrismatic Then
        ' SP3D Member Object Solid Port
        Set oMemberFactory = New SPSMembers.SPSMemberFactory
        Set oMemberConnectionServices = oMemberFactory.CreateConnectionServices
        
        oMemberConnectionServices.GetStructPortInfo oPortObject, ePortType, _
                                                    lContextID, lOperatationID, lOperatorID
        If lOperatorID = JXSEC_TOP Then
            sPortSide = C_Port_Top
        ElseIf lOperatorID = JXSEC_BOTTOM Then
            sPortSide = C_Port_Bottom
        ElseIf lOperatorID = JXSEC_WEB_LEFT Then
            sPortSide = C_Port_WebLeft
        ElseIf lOperatorID = JXSEC_WEB_RIGHT Then
            sPortSide = C_Port_WebRight
        Else
            lBaseCheck = eContextID And GSCADSDCreateModifyUtilities.CTX_BASE
            lNplusCheck = eContextID And GSCADSDCreateModifyUtilities.CTX_NPLUS
            lNminusCheck = eContextID And GSCADSDCreateModifyUtilities.CTX_NMINUS
            lOffsetCheck = eContextID And GSCADSDCreateModifyUtilities.CTX_OFFSET
            lLateralCheck = eContextID And GSCADSDCreateModifyUtilities.CTX_LATERAL
        
            If lBaseCheck <> 0 Then
                sPortSide = C_Port_Base
                
            ElseIf lOffsetCheck <> 0 Then
                sPortSide = C_Port_Offset
                    
            ElseIf lNplusCheck <> 0 Then
                sPortSide = C_Port_Base
                    
            ElseIf lNminusCheck <> 0 Then
                sPortSide = C_Port_Offset
            
            ElseIf lLateralCheck <> 0 Then
                sPortSide = C_Port_Lateral
            Else
                sPortSide = C_Port_Lateral
            End If
        End If
        
    ElseIf TypeOf oPortObject Is IMSStructConnection.IJStructPort Then
        Set oStructPort = oPortObject
        eContextID = oStructPort.ContextID
        lBaseCheck = eContextID And GSCADSDCreateModifyUtilities.CTX_BASE
        lNplusCheck = eContextID And GSCADSDCreateModifyUtilities.CTX_NPLUS
        lNminusCheck = eContextID And GSCADSDCreateModifyUtilities.CTX_NMINUS
        lOffsetCheck = eContextID And GSCADSDCreateModifyUtilities.CTX_OFFSET
        lLateralCheck = eContextID And GSCADSDCreateModifyUtilities.CTX_LATERAL
        
        If lBaseCheck <> 0 Then
            sPortSide = C_Port_Base
            
        ElseIf lOffsetCheck <> 0 Then
            sPortSide = C_Port_Offset
                
        ElseIf lNplusCheck <> 0 Then
            sPortSide = C_Port_Base
                
        ElseIf lNminusCheck <> 0 Then
            sPortSide = C_Port_Offset
        
        Else
            ' Connecting Port is not Base or Offset port
            ' check if Profile Top/Bottom/Left/Right port
            lOperatorID = oStructPort.OperatorID
            If lOperatorID = JXSEC_TOP Then
                sPortSide = C_Port_Top
            ElseIf lOperatorID = JXSEC_BOTTOM Then
                sPortSide = C_Port_Bottom
            ElseIf lOperatorID = JXSEC_WEB_LEFT Then
                sPortSide = C_Port_WebLeft
            ElseIf lOperatorID = JXSEC_WEB_RIGHT Then
                sPortSide = C_Port_WebRight
            
            ElseIf lLateralCheck <> 0 Then
                sPortSide = C_Port_Lateral
            
            Else
                sPortSide = C_Port_Lateral
            End If
            
        End If
    
    ElseIf TypeOf oPortObject Is ISPSSplitAxisPort Then
        ' SP3D Member Object Axis Port
        Set oSplitAxisPort = oPortObject
        eAxisPortIndex = oSplitAxisPort.portIndex
    
        If eAxisPortIndex = SPSMemberAxisStart Then
            sPortSide = C_Port_Base
        ElseIf eAxisPortIndex = SPSMemberAxisEnd Then
            sPortSide = C_Port_Offset
        ElseIf eAxisPortIndex = SPSMemberAxisAlong Then
            sPortSide = C_Port_Lateral
        Else
            sPortSide = ""
        End If

        
    End If
    
    Get_PortFaceType = sPortSide
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function
Private Function GetResourceMgr() As IJDPOM

    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim pConnMiddle As IJDConnectMiddle
    Dim pAccessMiddle As IJDAccessMiddle
    
    Dim jContext As IJContext
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
 
    Set pConnMiddle = jContext.GetService("ConnectMiddle")
 
    Set pAccessMiddle = pConnMiddle
 
    Dim strModelDB As String
    strModelDB = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set GetResourceMgr = pAccessMiddle.GetResourceManager(strModelDB)
  
      
    Set jContext = Nothing
    Set oDBTypeConfig = Nothing
    Set pConnMiddle = Nothing
    Set pAccessMiddle = Nothing
End Function
