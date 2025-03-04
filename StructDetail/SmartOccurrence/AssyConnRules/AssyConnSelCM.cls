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

Const MODULE = "AssyConnSelCM"

Private sError As String
Private sMETHOD As String

Public Sub ChamferQuestion(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
    ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.Definition
    
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
    
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oChamferPart1 As StructDetailObjects.PlatePart
    Dim oChamferPart2 As StructDetailObjects.PlatePart
    
    
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
    
    ' Know that we are called only from the PlateEdgeToPlateEdge selector
    ' Therefore we know the Ports are Edge face Ports from Plate Parts
    ' Check if the Plate parts are Leaf Parts from the same Plate System
    ' There is less Performance Issues calculating the Thickness differences
    ' using the Plate Parts Thickness/Offset Properties
    Set oSDO_Helper = New StructDetailObjects.Helper
    Set oRootParent1 = oSDO_Helper.Object_RootParentSystem(oChamferPart1.object)
    Set oRootParent2 = oSDO_Helper.Object_RootParentSystem(oChamferPart2.object)
    
    If oRootParent1 Is oRootParent2 Then
        ' Given Plate Parts are from the same Root Plate System
        ' Therefore there is a common Reference Surface and
        ' can use the Plate Thickness data to calculate thickness differences
        ChamferOffsetNormalData oChamferPart1, oChamferPart2, _
                                dBaseDelta, dOffsetDelta, bNormalConsistent

        If dBaseDelta < 0.005 And dBaseDelta > -0.005 Then
            dBaseDelta = 0  'no chamfer on base side
        End If
        
        If dOffsetDelta < 0.005 And dOffsetDelta > -0.005 Then
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
    
    ' Check/Verify if Chamfer data is valid
    If lStatusCode <> 0 Then
        Exit Function
    End If
    
    ' if Deltas less then 5mm, no Chamfer will be placed
    If Abs(dBaseDelta) < 0.005 Then
        If Abs(dOffsetDelta) < 0.005 Then
            Exit Function
        End If
        dBaseDelta = 0#
    End If
    
    If Abs(dOffsetDelta) < 0.005 Then
        dOffsetDelta = 0#
    End If
    
    If bBaseToBase Then
        ' The two parts are oriented the same
        ' dBaseDelta < 0: BasePort1 is below BasePort2 (Chamfer BasePort1)
        ' dBaseDelta > 0: BasePort1 is above BasePort2 (Chamfer BasePort2)
        ' dOffsetDelta < 0: dOffsetPort1 is below OffsetPort2 (Chamfer OffsetPort2)
        ' dOffsetDelta > 0: dOffsetPort1 is above OffsetPort2 (Chamfer OffsetPort1)
        If Abs(dOffsetDelta) < 0.005 Then
            ' Offset deltas are equal
            If dBaseDelta < 0# Then
                sChamferType = "Obj1Base"
            Else
                sChamferType = "Obj2Base"
            End If
            
        ElseIf dOffsetDelta > 0# Then
            If Abs(dBaseDelta) < 0.005 Then
                sChamferType = "Obj1Offset"
            ElseIf dBaseDelta < 0# Then
                sChamferType = "Obj1Double"
            Else
                sChamferType = "Obj1OffsetObj2Base"
            End If
        
        Else    '...(If dOffsetDelta < 0# Then)
            If Abs(dBaseDelta) < 0.005 Then
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
        If Abs(dOffsetDelta) < 0.005 Then
            ' Offset deltas are equal
            If dBaseDelta < 0# Then
                sChamferType = "Obj1Base"
            Else
                sChamferType = "Obj2Offset"
            End If
            
        ElseIf dOffsetDelta > 0# Then
            If Abs(dBaseDelta) < 0.005 Then
                sChamferType = "Obj1Offset"
            ElseIf dBaseDelta < 0# Then
                sChamferType = "Obj1Double"
            Else
                sChamferType = "Obj1OffsetObj2Offset"
            End If
        
        Else    '...(If dOffsetDelta < 0# Then)
            If Abs(dBaseDelta) < 0.005 Then
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
    Set oSymbolDefinition = oInputDG.Definition
    
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
        sSectionType = oSDO_BoundingProfile.SectionType
        
    ElseIf oAssyConn.ConnectedObject1Type = SDOBJECT_BEAM Then
        Dim oSDO_BoundingBeam As StructDetailObjects.BeamPart
        Set oSDO_BoundingBeam = New StructDetailObjects.BeamPart
        Set oSDO_BoundingBeam.object = oAssyConn.ConnectedObject1
        sSectionType = oSDO_BoundingBeam.SectionType
        
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
    Set oSymbolDefinition = oInputDG.Definition
    
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




