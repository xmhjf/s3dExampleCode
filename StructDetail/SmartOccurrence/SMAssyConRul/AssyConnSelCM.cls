VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssyConnSelCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'**********************************************************************************************************************
'  Copyright (C) 2011-15, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : EndCutDefCM.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'  CR-CP-273575:  Profile Connection type is used to set 'EndCutType' answer
'**********************************************************************************************************************

Const MODULE = "AssyConnSelCM"

Private sError As String
Private sMethod As String
Const PI = 3.14159265358979


Public Sub ChamferQuestion(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
    ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
   
    ' Get string to set the answer to Based on the Connecting Ports
    Dim strNewAnswer As String
    strNewAnswer = ChamferSelector(pSL.SmartOccurrence)
    
    pSL.Answer("ChamferType") = strNewAnswer
              
    Exit Sub
End Sub


Public Function ChamferSelector(oAC_Object As Object) As String
    On Error GoTo ErrorHandler

    Dim sChamferType As String

    Dim lStatusCode As Long

    Dim dEdgeAngle As Double
    Dim dBaseDelta As Double
    Dim dOffsetDelta As Double
    
    Dim bBaseToBase As Boolean
    Dim bNormalConsistent As Boolean
    
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    Dim oRootParent1 As Object
    Dim oRootParent2 As Object
    
    Dim e_tolerance As Double
    
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oChamferPart1 As StructDetailObjects.PlatePart
    Dim oChamferPart2 As StructDetailObjects.PlatePart
    
    Dim PrimaryMemberSystem As ISPSMemberSystem
    Set PrimaryMemberSystem = Nothing
    
    Dim bIsCan As Boolean
     
' Check/Verify type of object for Input1
    ChamferSelector = "None"
    If oAC_Object Is Nothing Then
        Exit Function
        
    ElseIf TypeOf oAC_Object Is IJAssemblyConnection Then
        Set oAssyConn = New StructDetailObjects.AssemblyConn
        Set oAssyConn.object = oAC_Object

        ' Check for Special case of Plate Edge to Plate Edge
        ' Builtup Members have connection cases where Plate Edge to Plate Edge
        ' are "Tee" type connections (not "Butt" connections)
        ' Chamfers are not valid for these "Tee" type connections
        If oAssyConn.IsEdgeToEdgeTeeConnection Then
       
            Exit Function
        End If

        Set oPort1 = oAssyConn.Port1
        Set oChamferPart1 = New StructDetailObjects.PlatePart
        Set oChamferPart1.object = oPort1.Connectable
    
        Set oPort2 = oAssyConn.Port2
        Set oChamferPart2 = New StructDetailObjects.PlatePart
        Set oChamferPart2.object = oPort2.Connectable
    
    Else
        Exit Function
    End If
    'code to check if the length edge of the built up is involved.
    'no chamfers in this case.
    
    Dim bBoolean As Boolean
    Dim mBoolean As Boolean
    Dim oBuiltupMember As ISPSDesignedMember
    
    IsPortFromBuiltUpMember oPort1, bBoolean, oBuiltupMember
    If bBoolean Then
        IsLengthEdgeOfBuiltup oPort1, oBuiltupMember, mBoolean
        If mBoolean Then
            If oChamferPart1.PlateThickness > oChamferPart2.PlateThickness Then
                'Get the thickness of the two parts and compare.
                sChamferType = "None"
                Exit Function
            Else
                'continue with further implementation
            End If
        End If
    End If

    IsPortFromBuiltUpMember oPort2, bBoolean, oBuiltupMember
    If bBoolean Then
        IsLengthEdgeOfBuiltup oPort2, oBuiltupMember, mBoolean
        If mBoolean Then
            If oChamferPart2.PlateThickness > oChamferPart1.PlateThickness Then
                'Get the thickness of the two parts and compare.
                sChamferType = "None"
                Exit Function
            Else
                'continue with further implementation
            End If
        End If
    End If

    'code to check if the chamfer is between APS and a secondary structure
    'Secondary structures are not identifiable. so, web plates, flange plates are identified and
    'chamfer is none if it is between APS and a plate(not built up)
    If oChamferPart1.plateType = GeneralPlate Then
        If Not oChamferPart2.plateType = WebPlate And Not oChamferPart2.plateType = FlangePlate Then
            sChamferType = "None"
            Exit Function
        End If
    End If
     If oChamferPart2.plateType = GeneralPlate Then
        If Not oChamferPart1.plateType = WebPlate And Not oChamferPart1.plateType = FlangePlate Then
            sChamferType = "None"
            Exit Function
        End If
    End If
    ' Know that we are called only from the PlateEdgeToPlateEdge selector
    ' Therefore we know the Ports are Edge face Ports from Plate Parts
    ' Check if the Plate parts are Leaf Parts from the same Plate System
    ' There is less Performance Issues calculating the Thickness differences
    ' using the Plate Parts Thickness/Offset Properties
    Set oSDO_Helper = New StructDetailObjects.Helper
    Set oRootParent1 = oSDO_Helper.Object_RootParentSystem(oChamferPart1.object)
    Set oRootParent2 = oSDO_Helper.Object_RootParentSystem(oChamferPart2.object)
        
    Dim oShpStrConnProperties As IJShpStrConnProperties
    Set oShpStrConnProperties = oAssyConn.object
    Dim dThicknessDiff As Double
    Dim dThicknessDiff_TN As Double
    Dim dAngle As Double
    Dim dAngle_unknown As Double
    dThicknessDiff = Abs(oChamferPart1.PlateThickness - oChamferPart2.PlateThickness)
    e_tolerance = 0.005
    
    If oRootParent1 Is oRootParent2 Then
        ' Given Plate Parts are from the same Root Plate System
        ' Therefore there is a common Reference Surface and
        ' can use the Plate Thickness data to calculate thickness differences
        'However, check to see if either part is portion of a CAN
        
        ChamferOffsetNormalData oChamferPart1, oChamferPart2, _
                                dBaseDelta, dOffsetDelta, bNormalConsistent
        
        'check to see if a CAN is involved.
        bIsCan = IsCanInvolved(oPort1, oPort2, PrimaryMemberSystem)
        If bIsCan Then
            oShpStrConnProperties.GetMountingAngle oPort1, oPort2, oAssyConn.ConnectedObject1, dAngle, dAngle_unknown
            If dAngle > PI Then
                dAngle = ((2 * PI) - dAngle)
            End If
            'dThicknessDiff_TN ---- thickness difference along tube normal
            'take normals of chamferpart1 and chamferpart2. compare the normals with normal of
            'built up can. take the part which is inclined i.e., cone part and get its adjusted thickness
            GetThicknessAlongTubeNormal oChamferPart1, oChamferPart2, PrimaryMemberSystem, dAngle, dThicknessDiff_TN
             If Abs(dThicknessDiff_TN) > 0.00001 Then
                e_tolerance = 0.000001
            End If
            dBaseDelta = dThicknessDiff_TN
        End If
        
        If dBaseDelta < e_tolerance And dBaseDelta > -e_tolerance Then
            dBaseDelta = 0  'no chamfer on base side
        End If
        
        If dOffsetDelta < e_tolerance And dOffsetDelta > -e_tolerance Then
            dOffsetDelta = 0 ' no chamfer on offset side
        End If
        
        If dOffsetDelta = 0 And dBaseDelta = 0 Then
                sChamferType = "None"
                
        ElseIf dOffsetDelta = 0 And dBaseDelta < 0 Then
                sChamferType = "Obj1Base"
        ElseIf dOffsetDelta > 0 And dBaseDelta = 0 Then
                sChamferType = "Obj1Offset"
        ElseIf dOffsetDelta > 0 And dBaseDelta < 0 Then
                sChamferType = "Obj1Double"
                
        ElseIf dOffsetDelta = 0 And dBaseDelta > 0 Then
                sChamferType = "Obj2Base"
        ElseIf dOffsetDelta < 0 And dBaseDelta = 0 Then
                sChamferType = "Obj2Offset"
        ElseIf dOffsetDelta < 0 And dBaseDelta > 0 Then
                sChamferType = "Obj2Double"
                
        ElseIf dOffsetDelta < 0 And dBaseDelta < 0 Then
                sChamferType = "Obj1BaseObj2Offset"
        ElseIf dOffsetDelta > 0 And dBaseDelta > 0 Then
                sChamferType = "Obj1OffsetObj2Base"
        End If
    
        ChamferSelector = sChamferType
        Exit Function
    End If
    
    ' (This might have an Impact on Performance)
    ' The Plate Parts are from different Root Parent Systems
    ' (Plate parts are from BuiltUps or Manual Logical Connections)
    ' Therefore: there is no common Reference Surface
    ' Calculate thickness differences using the Chamfer Sketch Plane
    
     oAssyConn.GetChamferBaseOffsetData dBaseDelta, dOffsetDelta, _
                                       bBaseToBase, dEdgeAngle, lStatusCode
    
      'check to see if a CAN is involved.
      bIsCan = IsCanInvolved(oPort1, oPort2, PrimaryMemberSystem)
        If bIsCan Then
            oShpStrConnProperties.GetMountingAngle oPort1, oPort2, oAssyConn.ConnectedObject1, dAngle, dAngle_unknown
            If dAngle > PI Then
                dAngle = ((2 * PI) - dAngle)
            End If
            GetThicknessAlongTubeNormal oChamferPart1, oChamferPart2, PrimaryMemberSystem, dAngle, dThicknessDiff_TN
            If Abs(dThicknessDiff_TN) > 0.00001 Then
                e_tolerance = 0.000001
            End If
        End If
    
   
    ' Check/Verify if Chamfer data is valid
    If lStatusCode <> 0 Then
        Exit Function
    End If
    
    ' if Deltas less then e_tolerance, no Chamfer will be placed
    If Abs(dBaseDelta) < e_tolerance Then
        If Abs(dOffsetDelta) < e_tolerance Then
            Exit Function
        End If
        dBaseDelta = 0#
    End If
    
    If Abs(dOffsetDelta) < e_tolerance Then
        dOffsetDelta = 0#
    End If
    
    If bBaseToBase Then
        ' The two parts are oriented the same
        ' dBaseDelta < 0: BasePort1 is below BasePort2 (Chamfer BasePort1)
        ' dBaseDelta > 0: BasePort1 is above BasePort2 (Chamfer BasePort2)
        ' dOffsetDelta < 0: dOffsetPort1 is below OffsetPort2 (Chamfer OffsetPort2)
        ' dOffsetDelta > 0: dOffsetPort1 is above OffsetPort2 (Chamfer OffsetPort1)
        If Abs(dOffsetDelta) < e_tolerance Then
            ' Offset deltas are equal
            If dBaseDelta < 0# Then
                sChamferType = "Obj1Base"
            Else
                sChamferType = "Obj2Base"
            End If
            
        ElseIf dOffsetDelta > 0# Then
            If Abs(dBaseDelta) < e_tolerance Then
                sChamferType = "Obj1Offset"
            ElseIf dBaseDelta < 0# Then
                sChamferType = "Obj1Double"
            Else
                sChamferType = "Obj1OffsetObj2Base"
            End If
        
        Else    '...(If dOffsetDelta < 0# Then)
            If Abs(dBaseDelta) < e_tolerance Then
                sChamferType = "Obj2Offset"
            ElseIf dBaseDelta < 0# Then
                sChamferType = "Obj1BaseObj2Offset"
            Else
                sChamferType = "Obj2Double"
            End If
        End If
    Else
        ' The two parts are NOT oriented the same
        ' The two parts are oriented the same
        ' dBaseDelta < 0: BasePort1 is below OffsetPort2 (Chamfer BasePort1)
        ' dBaseDelta > 0: BasePort1 is above OffsetPort2 (Chamfer OffsetPort2)
        ' dOffsetDelta < 0: dOffsetPort1 is below BasePort2 (Chamfer BasePort2)
        ' dOffsetDelta > 0: dOffsetPort1 is above BasePort2 (Chamfer OffsetPort1)
        If Abs(dOffsetDelta) < e_tolerance Then
            ' Offset deltas are equal
            If dBaseDelta < 0# Then
                sChamferType = "Obj1Base"
            Else
                sChamferType = "Obj2Offset"
            End If
            
        ElseIf dOffsetDelta > 0# Then
            If Abs(dBaseDelta) < e_tolerance Then
                sChamferType = "Obj1Offset"
            ElseIf dBaseDelta < 0# Then
                sChamferType = "Obj1Double"
            Else
                sChamferType = "Obj1OffsetObj2Offset"
            End If
        
        Else    '...(If dOffsetDelta < 0# Then)
            If Abs(dBaseDelta) < e_tolerance Then
                sChamferType = "Obj2Base"
            ElseIf dBaseDelta < 0# Then
                sChamferType = "Obj1BaseObj2Base"
            Else
                sChamferType = "Obj2Double"
            End If
        End If
    End If
    
    ' Return the Chamfer that is required
    ChamferSelector = sChamferType
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ChamferSelector").Number
End Function


Public Sub SetEndToEndDefault(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                              ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    pSL.Answer("EndToEndCase") = "NoAngle"
    Exit Sub
    
    ' Create Assembly Connection Object
    Dim oAssyConn As New StructDetailObjects.AssemblyConn
    Set oAssyConn.object = pSL.SmartOccurrence
 
    Dim sSectionType As String

    If oAssyConn.ConnectedObject1Type = SDOBJECT_STIFFENER Then
        Dim oSDO_BoundingProfile As StructDetailObjects.ProfilePart
        Set oSDO_BoundingProfile = New StructDetailObjects.ProfilePart
        Set oSDO_BoundingProfile.object = oAssyConn.ConnectedObject1
        sSectionType = oSDO_BoundingProfile.sectionType
        
    ElseIf oAssyConn.ConnectedObject1Type = SDOBJECT_BEAM Then
        Dim oSDO_BoundingBeam As StructDetailObjects.BeamPart
        Set oSDO_BoundingBeam = New StructDetailObjects.BeamPart
        Set oSDO_BoundingBeam.object = oAssyConn.ConnectedObject1
        sSectionType = oSDO_BoundingBeam.sectionType
        
    Else
        sSectionType = "???"
    End If
zMsgBox "AssyConnSelCM::SetEndToEndDefault" & vbCrLf & _
        "    sSectionType = " & sSectionType

    ' check the type of seam used to create the assy conn
    Dim sCase As String
    
    If oAssyConn.FromDesignSeam Then
            sCase = "Case1"
        
    ElseIf oAssyConn.FromIntersectionSeam Then
            sCase = "Case2"

    ElseIf oAssyConn.FromPlanningSeam Then
        sCase = "Case4"
    
    ElseIf oAssyConn.FromStrakingSeam Then
        sCase = "NoAngle"
    
    Else
        sCase = "NoAngle"
    End If
    
    pSL.Answer("EndToEndCase") = sCase
              
zMsgBox "AssyConnSelCM::SetEndToEndDefault" & vbCrLf & _
        "    pSL.Answer(EndToEndCase) = " & sCase
              
    Exit Sub
End Sub

Public Sub SetChamferValue(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                           ByRef ppArgument As Object)
    Dim strError As String
    On Error GoTo ErrorHandler
    
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    Dim iIndex As Long
    Dim dThickness As Double
    Dim oPartOnOtherSide As IJElements
    
    Dim pSL As IJDSelectorLogic
    Dim oSDO_AssyConn As StructDetailObjects.AssemblyConn
    Dim oSDO_PlatePart1 As StructDetailObjects.PlatePart
    Dim oSDO_PlatePart2 As StructDetailObjects.PlatePart
    
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    ' Create Assembly Connection Object
    Set oSDO_AssyConn = New StructDetailObjects.AssemblyConn
    Set oSDO_AssyConn.object = pSL.SmartOccurrence
    
    'Get object 1
    Set oSDO_PlatePart1 = New StructDetailObjects.PlatePart
    Set oSDO_PlatePart1.object = oSDO_AssyConn.ConnectedObject1
    
    ' Know that this Plate Edge Bounded by Plate Face case
    ' Want to determine if the Bounded Plate is Split by the Bounding Plate
    ' such that Leaf Part(s) with the same Root Part as the bounded Part
    ' exist on the "other" side of the Bounding Plate
    ' (only return Parts on Other Side that are in Range of bounded Part)
    Dim bPartsOnOtherSide As Boolean
    Set oPartOnOtherSide = oSDO_AssyConn.PartsOnOtherSide(False)
    If oPartOnOtherSide Is Nothing Then
        bPartsOnOtherSide = False
    ElseIf oPartOnOtherSide.Count < 1 Then
        bPartsOnOtherSide = False
    Else
        bPartsOnOtherSide = True
    End If
    
    'check if the assy conn came from an Split type seam
    If bPartsOnOtherSide Then
        'get the thickness of the part on the opposite side of the plate
        ' Use the First Part Thickness
        ' for Multiple Parts:
        '   could cycle thru all Parts, using the largest Thickness found
        ' Or
        '   cycle thru all Parts, finding the Part(s) that overlap only
        '
        Set oSDO_PlatePart2 = New StructDetailObjects.PlatePart
        Set oSDO_PlatePart2.object = oPartOnOtherSide.Item(1)
        dThickness = oSDO_PlatePart2.PlateThickness
        
        ' Search All Parts returned for the Thickest Part
        If oPartOnOtherSide.Count > 1 Then
            For iIndex = 2 To oPartOnOtherSide.Count
                Set oSDO_PlatePart2.object = oPartOnOtherSide.Item(iIndex)
                If oSDO_PlatePart2.PlateThickness > dThickness Then
                    dThickness = oSDO_PlatePart2.PlateThickness
                End If
            Next iIndex
        End If
        
    Else
        'use the thickness of the bounding object instead
        Set oSDO_PlatePart2 = New StructDetailObjects.PlatePart
        Set oSDO_PlatePart2.object = oSDO_AssyConn.ConnectedObject2
        
        'get the thickness and use it to set the chamfer default
        dThickness = oSDO_PlatePart2.PlateThickness
    End If
        
    pSL.Answer("ChamferThickness") = Trim(Str(dThickness))
                 
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetChamferValue", strError).Number
End Sub

Public Sub SetEndCutType(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    Dim oAC As New StructDetailObjects.AssemblyConn
    Set oAC.object = pSL.SmartOccurrence
    
    Dim oBoundedPart As Object
    Set oBoundedPart = oAC.ConnectedObject1
    
    If TypeOf oBoundedPart Is IJEndCutAttributes Then
    
        Dim oStructPort As IJStructPort
        Set oStructPort = oAC.Port1
                            
        Dim oEndCutAttributes As IJEndCutAttributes
        Set oEndCutAttributes = oBoundedPart
                        
        Dim lCodeListValue As Long
            
        'Get Stiffener connection type
        If oStructPort.ContextID = CTX_BASE Then
            lCodeListValue = oEndCutAttributes.StartConnectionType
        ElseIf oStructPort.ContextID = CTX_OFFSET Then
            lCodeListValue = oEndCutAttributes.EndConnectionType
        End If
            
        'Zero indicates "By Rule"
        If lCodeListValue > 0 Then
        
            Dim OSD_Helper As StructDetailObjects.Helper
            Set OSD_Helper = New StructDetailObjects.Helper
            
            Dim sShotDesc As String
            Dim sLongDesc As String
        
            OSD_Helper.GetCustomAttributeCodeListDescription "StiffenerConnectionType", lCodeListValue, _
                                                sShotDesc, sLongDesc
            pSL.Answer("EndCutType") = sShotDesc
            Exit Sub
        End If
    End If
    
    If oAC.ConnectedObject1Type = SDOBJECT_STIFFENER Then

        'If the stiffened plate is a bracket, then set the EndCutType to Snip
        Dim oPlate As Object
        Dim oStiffener As IJStiffener
        Set oStiffener = oAC.ConnectedObject1
        
        Set oPlate = oStiffener.PlateSystem
        
        If TypeOf oPlate Is IJBracketPlateSystem Then
            pSL.Answer("EndCutType") = "Snip"
        Else
            pSL.Answer("EndCutType") = "Welded"
        End If

    Else
        pSL.Answer("EndCutType") = "Welded"
    End If
    
    Exit Sub


ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetEndCutType").Number
End Sub

Public Sub PlaceBracketCM(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    'get the answer to the EndCutType question
    Dim sEndCutType As String
    sEndCutType = pSL.SelectorAnswer(CUSTOMERID + "AssyConRul.StiffEndByStiffFaceSel", "EndCutType")
    
    If sEndCutType = "Bracketed" Then
        pSL.Answer("PlaceBracket") = "Yes"
    Else
        pSL.Answer("PlaceBracket") = "No"
    End If
        

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "PlaceBracketCM").Number
End Sub

Public Sub SetShipyardName(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class

    On Error GoTo ErrorHandler

    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition

    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    pSL.Answer("Shipyard") = "DefaultShipyard"

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetShipyardName").Number
End Sub

Public Sub SetWeldingType(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class

    On Error GoTo ErrorHandler

    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition

    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    pSL.Answer("WeldingType") = "Manual"


    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetWeldingType").Number
End Sub

'***********************************************************************
' METHOD:  PlaceCornerFeatureCM
'
' DESCRIPTION:  It is the Custom method which sets the default answer
'               to the question "AddCornerFeature" for the Tripping Bracket when
'               it is bounded by plate face or stiffener face
'
' Condition :   The default answer is set "Yes" only when...
'
'     both the connected plates must be detailed
'     the corners of bracket should be intersecting
'        AC on support1 has responsibility to create corner between 1 and 2 or
'        AC on support3 has responsibility to create corner between 1 and 3
'            If both the connected objects of the Tripp. Bracket,
'            the connection Behaviour betweem them (commom AC of the connected objects)
'            is penetration (which indicates there is a existence of slot)
'                      only then answer is set to "Yes"
'
'
'***********************************************************************

Public Sub PlaceCornerFeatureCM(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    Dim oSDO_AC As New StructDetailObjects.AssemblyConn
    Set oSDO_AC.object = pSL.SmartOccurrence
    
    
    pSL.Answer("AddCornerFeature") = "No"
   'Logic whether corner feature is required or not should be dealt from here
    
    ' with the help of AC we are trying to get the both the connected objects
    ' one would be the Bracket and other the Support of the Bracket
    Dim oBracketPort As IJPort
    Dim oSupportPort As IJPort
    Dim oBracketPart As Object
    Dim oSupportRoot As Object
    
    Dim oPart1 As Object
    Dim oPart2 As Object
    
    Set oPart1 = oSDO_AC.ConnectedObject1
    Set oPart2 = oSDO_AC.ConnectedObject2
    
    Set oBracketPort = Nothing
    Set oSupportPort = Nothing
    Set oBracketPart = Nothing
    Set oSupportRoot = Nothing
    
    If Not IsBracket(oPart1) And Not IsBracket(oPart2) Then
        Exit Sub
    End If
    
    ' Find the bounded bracket (both parts could be brackets, if one is a boundary for the other)
    If IsBracket(oPart1) And Not IsBracket(oPart2) Then
        Set oBracketPort = oSDO_AC.Port1
        Set oSupportPort = oSDO_AC.Port2
    ElseIf Not IsBracket(oPart1) And IsBracket(oPart2) Then
        Set oBracketPort = oSDO_AC.Port2
        Set oSupportPort = oSDO_AC.Port1
    Else
        Dim oPort1 As IJStructPort
        Dim oPort2 As IJStructPort
        Set oPort1 = oSDO_AC.Port1
        Set oPort2 = oSDO_AC.Port2
        
        If (oPort1.ContextID And CTX_LATERAL) And Not (oPort2.ContextID And CTX_LATERAL) Then
            Set oBracketPort = oSDO_AC.Port1
            Set oSupportPort = oSDO_AC.Port2
        ElseIf Not (oPort1.ContextID And CTX_LATERAL) And (oPort2.ContextID And CTX_LATERAL) Then
            Set oBracketPort = oSDO_AC.Port2
            Set oSupportPort = oSDO_AC.Port1
        Else
            ' this must be a seam in the bracket.  We're not interested in these.
            Exit Sub
        End If
    End If
    
    Set oBracketPart = oBracketPort.Connectable
        
    Dim oStructDetailHelper As StructDetailHelper
    Set oStructDetailHelper = New StructDetailHelper
    
    ' getting the system part of the Support
    oStructDetailHelper.IsPartDerivedFromSystem oSupportPort.Connectable, oSupportRoot, True
   
    
    If oSupportRoot Is Nothing Then
        Set oSupportRoot = oSupportPort.Connectable
    End If
   
   Dim oS1 As Object
   Dim oS2 As Object
   Dim oS3 As Object
   Dim oS4 As Object
   Dim oS5 As Object
   
   Dim oS1Port As IJPort
   Dim oS2Port As IJPort
   Dim oS3Port As IJPort
   Dim oS4Port As IJPort
   Dim oS5Port As IJPort
   
   GetBracketSupports oBracketPart, oS1, oS2, oS3, oS4, oS5, oS1Port, oS2Port, oS3Port, oS4Port, oS5Port
   
   Dim oPosition_between_S1S2 As IJDPosition
   Dim oPosition_between_S1S3 As IJDPosition
   Dim oPosition_between_S2S4 As IJDPosition
   Dim oPosition_between_S3S5 As IJDPosition
   
   Set oPosition_between_S1S2 = Nothing
   Set oPosition_between_S1S3 = Nothing
   Set oPosition_between_S2S4 = Nothing
   Set oPosition_between_S3S5 = Nothing
  
   Dim oEdgePort1ForCut As IJPort
   Dim oEdgePort2ForCut As IJPort
   Dim oEdgePort3ForCut As IJPort
   Dim oEdgePort4ForCut As IJPort
   Dim oEdgePort5ForCut As IJPort
   
   If Not oS1Port Is Nothing Then
         Set oEdgePort1ForCut = RelatedPortBeforeCut(oS1Port)
   Else
         Set oEdgePort1ForCut = Nothing
   End If
   
   If Not oS2Port Is Nothing Then
       Set oEdgePort2ForCut = RelatedPortBeforeCut(oS2Port)
   Else
         Set oEdgePort2ForCut = Nothing
   End If
   
   If Not oS3Port Is Nothing Then
       Set oEdgePort3ForCut = RelatedPortBeforeCut(oS3Port)
   Else
         Set oEdgePort3ForCut = Nothing
   End If
   
   If Not oS4Port Is Nothing Then
       Set oEdgePort4ForCut = RelatedPortBeforeCut(oS4Port)
   Else
         Set oEdgePort4ForCut = Nothing
   End If
   
   If Not oS5Port Is Nothing Then
       Set oEdgePort5ForCut = RelatedPortBeforeCut(oS5Port)
   Else
         Set oEdgePort5ForCut = Nothing
   End If
    
    Dim oIntersect As TopologyLocate

    Set oIntersect = New TopologyLocate
'   Detemines if there is an intersection between both the Bracket ports of S1 and S2
    If Not oEdgePort1ForCut Is Nothing And _
       Not oEdgePort2ForCut Is Nothing Then
           Set oPosition_between_S1S2 = oIntersect.FindIntersectionPoint(oEdgePort1ForCut, oEdgePort2ForCut)
    Else
          Set oPosition_between_S1S2 = Nothing
    End If
    
    If Not oEdgePort1ForCut Is Nothing And _
       Not oEdgePort3ForCut Is Nothing Then
          Set oPosition_between_S1S3 = oIntersect.FindIntersectionPoint(oEdgePort1ForCut, oEdgePort3ForCut)
   Else
          Set oPosition_between_S1S3 = Nothing
   End If
   
   If Not oEdgePort2ForCut Is Nothing And _
      Not oEdgePort4ForCut Is Nothing Then
          Set oPosition_between_S2S4 = oIntersect.FindIntersectionPoint(oEdgePort2ForCut, oEdgePort4ForCut)
   Else
          Set oPosition_between_S2S4 = Nothing
   End If
   
   If Not oEdgePort3ForCut Is Nothing And _
      Not oEdgePort5ForCut Is Nothing Then
          Set oPosition_between_S3S5 = oIntersect.FindIntersectionPoint(oEdgePort3ForCut, oEdgePort5ForCut)
   Else
          Set oPosition_between_S3S5 = Nothing
   End If
   
   'CMT changed oS1 to oS2 so that S2 is responsible for creating the corner feature between S1 and S2.
   'Was getting two corner features at teh S1S3 intersection, because both S1 and S3 were creating corner features
   If IsSupportOrPlateForStiffenerSupport(oSupportRoot, oS2) And _
       oSupportRoot Is oS2 And _
           Not oPosition_between_S1S2 Is Nothing And _
       Not oS1Port Is Nothing And _
       Not oS2Port Is Nothing _
   Or _
      IsSupportOrPlateForStiffenerSupport(oSupportRoot, oS3) And _
        oSupportRoot Is oS3 And _
                Not oPosition_between_S1S3 Is Nothing And _
        Not oS1Port Is Nothing And _
        Not oS3Port Is Nothing _
    Or _
      IsSupportOrPlateForStiffenerSupport(oSupportRoot, oS4) And _
        oSupportRoot Is oS4 And _
                Not oPosition_between_S2S4 Is Nothing And _
        Not oS2Port Is Nothing And _
        Not oS4Port Is Nothing _
    Or _
       IsSupportOrPlateForStiffenerSupport(oSupportRoot, oS5) And _
        oSupportRoot Is oS5 And _
                Not oPosition_between_S3S5 Is Nothing And _
        Not oS3Port Is Nothing And _
        Not oS5Port Is Nothing _
    Then
         
          pSL.Answer("AddCornerFeature") = "Yes"
          
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "PlaceCornerFeatureCM").Number
End Sub

Public Sub ChamferQuestionOnER(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
    ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    ' Get string to set the answer to Based on the Connecting Ports
    Dim strNewAnswer As String
    strNewAnswer = ChamferSelectorOnER(pSL.SmartOccurrence)
    
    pSL.Answer("ChamferType") = strNewAnswer
    
    Exit Sub
End Sub

Public Function ChamferSelectorOnER(oAC_Object As Object) As String
    On Error GoTo ErrorHandler
    
    Dim sChamferType As String
    Dim sSectionType As String
    
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oProfilePart1 As StructDetailObjects.ProfilePart
    Dim oProfilePart2 As StructDetailObjects.ProfilePart
    
    Dim oRootParent1 As Object
    Dim oRootParent2 As Object
    
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    Dim dThicknessDiff As Double

    sChamferType = "None"
    ChamferSelectorOnER = sChamferType
    
    If oAC_Object Is Nothing Then
     Exit Function
     
    ElseIf TypeOf oAC_Object Is IJAssemblyConnection Then
        Set oAssyConn = New StructDetailObjects.AssemblyConn
        Set oAssyConn.object = oAC_Object
        
        Set oPort1 = oAssyConn.Port1
        Set oProfilePart1 = New StructDetailObjects.ProfilePart
        Set oProfilePart1.object = oPort1.Connectable
        sSectionType = oProfilePart1.sectionType
        
        Set oPort2 = oAssyConn.Port2
        Set oProfilePart2 = New StructDetailObjects.ProfilePart
        Set oProfilePart2.object = oPort2.Connectable
     Else
        Exit Function
     End If
     
    If Not TypeOf oProfilePart1.object Is IJProfileER Then Exit Function
    
    Set oSDO_Helper = New StructDetailObjects.Helper
    
    Set oRootParent1 = oSDO_Helper.Object_RootParentSystem(oProfilePart1.object)
    Set oRootParent2 = oSDO_Helper.Object_RootParentSystem(oProfilePart2.object)
        
    If Not oRootParent1 Is oRootParent2 Then Exit Function
    
    dThicknessDiff = Abs(oProfilePart1.webThickness - oProfilePart2.webThickness)
    
    If (sSectionType = "FB") Then
        If dThicknessDiff > 0.005 Then
            Dim oERPort As IJPort
            Dim oStructPort As IJStructPort
      
            If oProfilePart1.webThickness > oProfilePart2.webThickness Then
                Set oERPort = oProfilePart1.MountingFacePort
                Set oStructPort = oERPort
                
                If oStructPort.OperatorID = JXSEC_WEB_LEFT Then
                    sChamferType = "Obj1Offset"
                ElseIf oStructPort.OperatorID = JXSEC_WEB_RIGHT Then
                    sChamferType = "Obj1Base"
                End If
            ElseIf oProfilePart2.webThickness > oProfilePart1.webThickness Then
                Set oERPort = oProfilePart2.MountingFacePort
                Set oStructPort = oERPort
                
                If oStructPort.OperatorID = JXSEC_WEB_LEFT Then
                    sChamferType = "Obj2Offset"
                ElseIf oStructPort.OperatorID = JXSEC_WEB_RIGHT Then
                    sChamferType = "Obj2Base"
                End If
            End If
        End If
    Else
    'Presently not handling other type of cross sections.
    End If
      
    ChamferSelectorOnER = sChamferType
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ChamferSelectorOnER").Number
End Function


