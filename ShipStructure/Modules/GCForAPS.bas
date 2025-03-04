Attribute VB_Name = "GCForAPS"
Option Explicit
Function CreateGeometricConstructionMacroOfExtractPorts(pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory, pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                                                        oMember As Object, sKeyOfMember As String, _
                                                        oLineOfNormal As Object) As IJGeometricConstructionMacro
    Call DebugIn("CreateGeometricConstructionMacroOfExtractPorts")
    
    ' create GCMacro transient
    Dim pGeometricConstructionTransient As IJGeometricConstruction
    Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "ExtractPorts", pPOM)
        
    ' feed GCMacro with inputs and parameters
    pGeometricConstructionTransient.Inputs("MemberPart").Add oMember
    pGeometricConstructionTransient.Inputs("ZAxis").Add oLineOfNormal
    Dim oPort As Object:
    On Error Resume Next: Set oPort = Nothing: Set oPort = pGeometricConstruction.ControlledInputs("AxisPort1")(sKeyOfMember): On Error GoTo 0
    If Not oPort Is Nothing Then _
        pGeometricConstructionTransient.ControlledInputs("AxisPort").Add pGeometricConstruction.ControlledInputs("AxisPort1")(sKeyOfMember)
    On Error Resume Next: Set oPort = Nothing: Set oPort = pGeometricConstruction.ControlledInputs("FacePort1")(sKeyOfMember): On Error GoTo 0
    If Not oPort Is Nothing Then _
        pGeometricConstructionTransient.ControlledInputs("FacePort").Add pGeometricConstruction.ControlledInputs("FacePort1")(sKeyOfMember)
        
    'If mirror was along x or y, the order of ports is reversed, we should reverse the support parameter
    Dim pGCPrivateAccess As IJGCPrivateAccess: Set pGCPrivateAccess = pGeometricConstruction
    Dim pGCType As IJGeometricConstructionType: Set pGCType = pGCPrivateAccess.GeometricConstructionType
    If pGCType.IsParameterBulkloaded("Orientation") = True Then
        If pGeometricConstruction.Parameter("Orientation") = GCIndirect Then
            'Reverse support parameter
            Select Case pGeometricConstruction.Parameter("Support")
                Case 1:
                    pGeometricConstructionTransient.Parameter("Support") = 2
                Case 2:
                    pGeometricConstructionTransient.Parameter("Support") = 1
            End Select
        Else
            pGeometricConstructionTransient.Parameter("Support") = pGeometricConstruction.Parameter("Support")
        End If
    Else
        pGeometricConstructionTransient.Parameter("Support") = pGeometricConstruction.Parameter("Support")
    End If
        
    ' evaluate GCMacro
    pGeometricConstructionTransient.Public = True
    pGeometricConstructionTransient.Evaluate
    Call Elements_ReplaceElementsWithKey(pGeometricConstruction.ControlledInputs("AxisPort1"), pGeometricConstructionTransient.ControlledInputs("AxisPort"), sKeyOfMember)
    On Error Resume Next
    Call Elements_ReplaceElementsWithKey(pGeometricConstruction.ControlledInputs("FacePort1"), pGeometricConstructionTransient.ControlledInputs("FacePort"), sKeyOfMember)
    On Error GoTo 0
    
    ' return result
    Set CreateGeometricConstructionMacroOfExtractPorts = pGeometricConstructionTransient
    Call DebugOut
End Function
Function CreateGeometricConstructionMacroOfExtractCircle(pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory, pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction) As IJGeometricConstructionMacro
    Call DebugIn("CreateGeometricConstructionMacroOfExtractCircle")
     
     ' create GCMacro transient
    Dim pGeometricConstructionTransient As IJGeometricConstruction
    Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "ExtractCircle", pPOM)
        
    ' feed GCMacro with inputs and parameters
    pGeometricConstructionTransient.Inputs("MemberPart1").Add pGeometricConstruction.Input("OrthogonalMemberPart")
    pGeometricConstructionTransient.Inputs("MemberPart2").Add pGeometricConstruction.Inputs("MemberParts")(1)
    pGeometricConstructionTransient.ControlledInputs("AxisPort1").AddElements pGeometricConstruction.ControlledInputs("AxisPort2")
    pGeometricConstructionTransient.ControlledInputs("FacePort1").AddElements pGeometricConstruction.ControlledInputs("FacePort2")
    
    ' evaluate GCMacro
    pGeometricConstructionTransient.Public = True
    Call GCEvaluate(pGeometricConstructionTransient, pGeometricConstruction, -1, -1)
    Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("AxisPort2"), pGeometricConstructionTransient.ControlledInputs("AxisPort1"))
    Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort2"), pGeometricConstructionTransient.ControlledInputs("FacePort1"))

    ' return result
    Set CreateGeometricConstructionMacroOfExtractCircle = pGeometricConstructionTransient
    Call DebugOut
End Function
Function CreateGeometricConstructionMacroOfStraightEdgeAround(pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory, pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                                                              oMember As Object, _
                                                              pGeometricConstructionMacroOfExtractPort1 As IJGeometricConstructionMacro, _
                                                              pGeometricConstructionMacroOfExtractPort2 As IJGeometricConstructionMacro, _
                                                              pGeometricConstructionMacroOfExtractCircle As IJGeometricConstructionMacro, _
                                                              pLineOfNormal As IJLine, dAngle As Double) As IJGeometricConstructionMacro
    Call DebugIn("CreateGeometricConstructionMacroOfStraightEdgeAround")
     
     ' create GCMacro transient
    Dim pGeometricConstructionTransient As IJGeometricConstruction
    Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "StraightEdgeAround", pPOM)
            
    ' feed GCMacro with inputs and parameters
    pGeometricConstructionTransient.Inputs("AxisLine1").Add pGeometricConstructionMacroOfExtractPort1.Outputs("MemberAxis").Item(1)
    pGeometricConstructionTransient.Inputs("AxisLine2").Add pGeometricConstructionMacroOfExtractPort2.Outputs("MemberAxis").Item(1)
    pGeometricConstructionTransient.Inputs("EdgeLine1").Add pGeometricConstructionMacroOfExtractPort1.Outputs("CurveLeft").Item(1)
    pGeometricConstructionTransient.Inputs("EdgeLine2").Add pGeometricConstructionMacroOfExtractPort2.Outputs("CurveRight").Item(1)
    pGeometricConstructionTransient.Inputs("Circle").Add pGeometricConstructionMacroOfExtractCircle.Outputs("Circle").Item(1)
    pGeometricConstructionTransient.Inputs("NormalLine").Add pLineOfNormal
    pGeometricConstructionTransient.Parameter("WeldToe") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
    pGeometricConstructionTransient.Parameter("Offset") = CDbl(pGeometricConstruction.Parameter("Offset"))
    pGeometricConstructionTransient.Parameter("Fillet") = CDbl(pGeometricConstruction.Parameter("Fillet"))
    pGeometricConstructionTransient.Parameter("CutBack") = CDbl(pGeometricConstruction.Parameter("CutBack"))
    pGeometricConstructionTransient.Parameter("RibbonHeight") = 2 * MemberPart_GetMaximumSize(oMember) 'dDepth / 2 'dDepth + 2 * pGeometricConstruction.Parameter("OverSize")
            
    ' evaluate GCMacro
    pGeometricConstructionTransient.Public = True
    Call GCEvaluate(pGeometricConstructionTransient, pGeometricConstruction, -1, -1)

    ' return result
    Set CreateGeometricConstructionMacroOfStraightEdgeAround = pGeometricConstructionTransient
    Call DebugOut
End Function
Function CreateGeometricConstructionMacroOfStraightEdgeAround2(pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory, pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                                                              oMember As Object, sCutBack1 As String, sCutBack2 As String, _
                                                              pGeometricConstructionMacroOfExtractPort1 As IJGeometricConstructionMacro, _
                                                              pGeometricConstructionMacroOfExtractPort2 As IJGeometricConstructionMacro, _
                                                              pGeometricConstructionMacroOfExtractCircle As IJGeometricConstructionMacro, _
                                                              pLineOfNormal As IJLine, dAngle As Double) As IJGeometricConstructionMacro
    Call DebugIn("CreateGeometricConstructionMacroOfStraightEdgeAround2")
     
     ' create GCMacro transient
    Dim pGeometricConstructionTransient As IJGeometricConstruction
    Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "StraightEdgeAroun2", pPOM)
            
    ' feed GCMacro with inputs and parameters
    pGeometricConstructionTransient.Inputs("AxisLine1").Add pGeometricConstructionMacroOfExtractPort1.Outputs("MemberAxis").Item(1)
    pGeometricConstructionTransient.Inputs("AxisLine2").Add pGeometricConstructionMacroOfExtractPort2.Outputs("MemberAxis").Item(1)
    pGeometricConstructionTransient.Inputs("EdgeLine1").Add pGeometricConstructionMacroOfExtractPort1.Outputs("CurveLeft").Item(1)
    pGeometricConstructionTransient.Inputs("EdgeLine2").Add pGeometricConstructionMacroOfExtractPort2.Outputs("CurveRight").Item(1)
    pGeometricConstructionTransient.Inputs("Circle").Add pGeometricConstructionMacroOfExtractCircle.Outputs("Circle").Item(1)
    pGeometricConstructionTransient.Inputs("NormalLine").Add pLineOfNormal
    pGeometricConstructionTransient.Parameter("WeldToe") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
    pGeometricConstructionTransient.Parameter("Offset") = CDbl(pGeometricConstruction.Parameter("Offset"))
    pGeometricConstructionTransient.Parameter("Fillet") = CDbl(pGeometricConstruction.Parameter("Fillet"))
    pGeometricConstructionTransient.Parameter("CutBack") = CDbl(pGeometricConstruction.Parameter(sCutBack1))
    pGeometricConstructionTransient.Parameter("CutBack2") = CDbl(pGeometricConstruction.Parameter(sCutBack2))
    pGeometricConstructionTransient.Parameter("RibbonHeight") = 2 * MemberPart_GetMaximumSize(oMember) 'dDepth / 2 'dDepth + 2 * pGeometricConstruction.Parameter("OverSize")
            
    ' evaluate GCMacro
    pGeometricConstructionTransient.Public = True
    Call GCEvaluate(pGeometricConstructionTransient, pGeometricConstruction, -1, -1)

    ' return result
    Set CreateGeometricConstructionMacroOfStraightEdgeAround2 = pGeometricConstructionTransient
    Call DebugOut
End Function
Function CreateGeometricConstructionMacroOfStraightEdgeAround3(pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory, pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                                                              oMember As Object, _
                                                              pGeometricConstructionMacroOfExtractPorts1 As IJGeometricConstructionMacro, _
                                                              pGeometricConstructionMacroOfExtractPorts2 As IJGeometricConstructionMacro, _
                                                              pGeometricConstructionMacroOfExtractCircle As IJGeometricConstructionMacro, _
                                                              pLineOfNormal As IJLine, dAngle As Double) As IJGeometricConstructionMacro
     ' create GCMacro transient
    Dim pGeometricConstructionTransient As IJGeometricConstruction
    If dAngle <= 175 Then
        Set pGeometricConstructionTransient = pGeometricConstructionFactory.CreateEntity("FreeEdgeF14", pPOM)
                
        ' feed GCMacro with inputs and parameters
        pGeometricConstructionTransient.Inputs("AxisLine1").Add pGeometricConstructionMacroOfExtractPorts1.Outputs("CurveLeft")(1)
        pGeometricConstructionTransient.Inputs("AxisLine2").Add pGeometricConstructionMacroOfExtractPorts2.Outputs("CurveRight")(1)
        pGeometricConstructionTransient.Inputs("Circle").Add pGeometricConstructionMacroOfExtractCircle.Outputs("Circle")(1)
        pGeometricConstructionTransient.Parameter("Offset") = CDbl(pGeometricConstruction.Parameter("Offset"))
        pGeometricConstructionTransient.Parameter("CutBack") = CDbl(pGeometricConstruction.Parameter("CutBack"))
        pGeometricConstructionTransient.Parameter("RibbonHeight") = 2 * MemberPart_GetMaximumSize(oMember) 'dDepth / 2 'dDepth + 2 * pGeometricConstruction.Parameter("OverSize")
            
        ' evaluate GCMacro
        pGeometricConstructionTransient.Public = True
        Call GCEvaluate(pGeometricConstructionTransient, pGeometricConstruction, -1, -1)
    Else
        Set pGeometricConstructionTransient = pGeometricConstructionFactory.CreateEntity("StraightEdgeAround", pPOM)
                
        ' feed GCMacro with inputs and parameters
        pGeometricConstructionTransient.Inputs("AxisLine1").Add pGeometricConstructionMacroOfExtractPorts1.Outputs("MemberAxis").Item(1)
        pGeometricConstructionTransient.Inputs("AxisLine2").Add pGeometricConstructionMacroOfExtractPorts2.Outputs("MemberAxis").Item(1)
        pGeometricConstructionTransient.Inputs("EdgeLine1").Add pGeometricConstructionMacroOfExtractPorts1.Outputs("CurveLeft").Item(1)
        pGeometricConstructionTransient.Inputs("EdgeLine2").Add pGeometricConstructionMacroOfExtractPorts2.Outputs("CurveRight").Item(1)
        pGeometricConstructionTransient.Inputs("Circle").Add pGeometricConstructionMacroOfExtractCircle.Outputs("Circle").Item(1)
        pGeometricConstructionTransient.Inputs("NormalLine").Add pLineOfNormal
        pGeometricConstructionTransient.Parameter("WeldToe") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
        pGeometricConstructionTransient.Parameter("Offset") = CDbl(pGeometricConstruction.Parameter("Offset"))
        pGeometricConstructionTransient.Parameter("Fillet") = CDbl(pGeometricConstruction.Parameter("Fillet"))
        pGeometricConstructionTransient.Parameter("CutBack") = CDbl(pGeometricConstruction.Parameter("CutBack"))
        pGeometricConstructionTransient.Parameter("RibbonHeight") = 2 * MemberPart_GetMaximumSize(oMember) 'dDepth / 2 'dDepth + 2 * pGeometricConstruction.Parameter("OverSize")
                
        ' evaluate GCMacro
        pGeometricConstructionTransient.Public = True
        Call GCEvaluate(pGeometricConstructionTransient, pGeometricConstruction, -1, -1)
    End If

    ' return result
    Set CreateGeometricConstructionMacroOfStraightEdgeAround3 = pGeometricConstructionTransient
End Function
Function CreateGeometricConstructionMacroOfFreeEdge(pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory, pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                                                    oMember As Object, _
                                                    pGeometricConstructionMacrosOfExtractPorts1 As IJGeometricConstructionMacro, _
                                                    pGeometricConstructionMacrosOfExtractPorts2 As IJGeometricConstructionMacro, _
                                                    dAngle As Double) As IJGeometricConstructionMacro
    Call DebugIn("CreateGeometricConstructionMacroOfStraightEdgeAround2")
        
    Dim pGeometricConstructionTransient As IJGeometricConstruction
    If dAngle <= 135 Then
        ' create GCMacro transient
        Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "FreeEdgeF13", pPOM)
            
        ' feed GCMacro with inputs and parameters
        pGeometricConstructionTransient.Inputs("AxisLine1").Add pGeometricConstructionMacrosOfExtractPorts1.Outputs("CurveLeft")(1)
        pGeometricConstructionTransient.Inputs("AxisLine2").Add pGeometricConstructionMacrosOfExtractPorts2.Outputs("CurveRight")(1)
        pGeometricConstructionTransient.Parameter("WeldToe") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
        pGeometricConstructionTransient.Parameter("Fillet") = CDbl(pGeometricConstruction.Parameter("Fillet"))
        pGeometricConstructionTransient.Parameter("CutBack") = CDbl(pGeometricConstruction.Parameter("CutBack"))
        pGeometricConstructionTransient.Parameter("RibbonHeight") = 2 * MemberPart_GetMaximumSize(oMember) 'dDepth / 2 'dDepth + 2 * pGeometricConstruction.Parameter("OverSize")
        pGeometricConstructionTransient.Parameter("Side") = 1
            
        ' evaluate GCMacro
        pGeometricConstructionTransient.Public = True
        Call GCEvaluate(pGeometricConstructionTransient, pGeometricConstruction, -1, -1)
    ElseIf dAngle <= 175 Then
        ' create GCMacro transient
        Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "FreeEdgeF12", pPOM)
            
        ' feed GCMacro with inputs and parameters
        pGeometricConstructionTransient.Inputs("AxisLine1").Add pGeometricConstructionMacrosOfExtractPorts1.Outputs("CurveLeft")(1)
        pGeometricConstructionTransient.Inputs("AxisLine2").Add pGeometricConstructionMacrosOfExtractPorts2.Outputs("CurveRight")(1)
        pGeometricConstructionTransient.Parameter("WeldToe") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
        pGeometricConstructionTransient.Parameter("Fillet") = CDbl(pGeometricConstruction.Parameter("Fillet"))
        pGeometricConstructionTransient.Parameter("CutBack") = CDbl(pGeometricConstruction.Parameter("CutBack"))
        pGeometricConstructionTransient.Parameter("ArmLength") = CDbl(pGeometricConstruction.Parameter("ArmLength"))
        pGeometricConstructionTransient.Parameter("RibbonHeight") = 2 * MemberPart_GetMaximumSize(oMember) 'dDepth / 2 'dDepth + 2 * pGeometricConstruction.Parameter("OverSize")
            
        ' evaluate GCMacro
        pGeometricConstructionTransient.Public = True
        Call GCEvaluate(pGeometricConstructionTransient, pGeometricConstruction, -1, -1)
    ElseIf dAngle <= 185 Then
        ' create GCMacro transient
        Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "StraightEdge", pPOM)
            
        ' feed GCMacro with inputs and parameters
        pGeometricConstructionTransient.Inputs("Support1").Add pGeometricConstructionMacrosOfExtractPorts1.Outputs("Support")(1)
        pGeometricConstructionTransient.Inputs("Support2").Add pGeometricConstructionMacrosOfExtractPorts2.Outputs("Support")(1)
        pGeometricConstructionTransient.Inputs("MemberAxis1").Add pGeometricConstructionMacrosOfExtractPorts1.Outputs("MemberAxis")(1)
        pGeometricConstructionTransient.Inputs("MemberAxis2").Add pGeometricConstructionMacrosOfExtractPorts2.Outputs("MemberAxis")(1)
        pGeometricConstructionTransient.Inputs("EdgePort1").Add pGeometricConstructionMacrosOfExtractPorts1.Outputs("CurveLeft")(1)
        pGeometricConstructionTransient.Inputs("EdgePort2").Add pGeometricConstructionMacrosOfExtractPorts2.Outputs("CurveRight")(1)
        pGeometricConstructionTransient.Parameter("WeldToe") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
        pGeometricConstructionTransient.Parameter("CutBack") = CDbl(pGeometricConstruction.Parameter("CutBack"))
        pGeometricConstructionTransient.Parameter("RibbonHeight") = 2 * MemberPart_GetMaximumSize(oMember) 'dDepth / 2 'dDepth + 2 * pGeometricConstruction.Parameter("OverSize")
            
        ' evaluate GCMacro
        pGeometricConstructionTransient.Public = True
        Call GCEvaluate(pGeometricConstructionTransient, pGeometricConstruction, -1, -1)
    Else
        ' create GCMacro transient
        Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "FreeEdgeF13", pPOM)
            
        ' feed GCMacro with inputs and parameters
        pGeometricConstructionTransient.Inputs("AxisLine1").Add pGeometricConstructionMacrosOfExtractPorts1.Outputs("CurveLeft")(1)
        pGeometricConstructionTransient.Inputs("AxisLine2").Add pGeometricConstructionMacrosOfExtractPorts2.Outputs("CurveRight")(1)
        pGeometricConstructionTransient.Parameter("WeldToe") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
        pGeometricConstructionTransient.Parameter("Fillet") = 0.05
        pGeometricConstructionTransient.Parameter("CutBack") = CDbl(pGeometricConstruction.Parameter("CutBack"))
        pGeometricConstructionTransient.Parameter("RibbonHeight") = 2 * MemberPart_GetMaximumSize(oMember) 'dDepth / 2 'dDepth + 2 * pGeometricConstruction.Parameter("OverSize")
        pGeometricConstructionTransient.Parameter("Side") = 2
            
        ' evaluate GCMacro
        pGeometricConstructionTransient.Public = True
        Call GCEvaluate(pGeometricConstructionTransient, pGeometricConstruction, -1, -1)
    End If
    
    ' return result
    Set CreateGeometricConstructionMacroOfFreeEdge = pGeometricConstructionTransient
    Call DebugOut
End Function
Function CreateGeometricConstructionMacroOfFreeEdge2(pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory, pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                                                    oMember As Object, sCutBack1 As String, sCutBack2 As String, _
                                                    pGeometricConstructionMacrosOfExtractPorts1 As IJGeometricConstructionMacro, _
                                                    pGeometricConstructionMacrosOfExtractPorts2 As IJGeometricConstructionMacro, _
                                                    dAngle As Double) As IJGeometricConstructionMacro
    Call DebugIn("CreateGeometricConstructionMacroOfFreeEdge2")
        
    If dAngle <= 175 Then
        ' create GCMacro transient
        Dim pGeometricConstructionTransient As IJGeometricConstruction
        Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "FreeEdgeF13b", pPOM)
                
        ' feed GCMacro with inputs and parameters
        pGeometricConstructionTransient.Inputs("AxisLine1").Add pGeometricConstructionMacrosOfExtractPorts1.Outputs("CurveLeft")(1)
        pGeometricConstructionTransient.Inputs("AxisLine2").Add pGeometricConstructionMacrosOfExtractPorts2.Outputs("CurveRight")(1)
        pGeometricConstructionTransient.Parameter("WeldToe") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
        pGeometricConstructionTransient.Parameter("Fillet") = CDbl(pGeometricConstruction.Parameter("Fillet"))
        pGeometricConstructionTransient.Parameter("CutBack") = CDbl(pGeometricConstruction.Parameter(sCutBack1))
        pGeometricConstructionTransient.Parameter("CutBack2") = CDbl(pGeometricConstruction.Parameter(sCutBack2))
        pGeometricConstructionTransient.Parameter("RibbonHeight") = 2 * MemberPart_GetMaximumSize(oMember) 'dDepth / 2 'dDepth + 2 * pGeometricConstruction.Parameter("OverSize")
        pGeometricConstructionTransient.Parameter("Side") = 1
                
        ' evaluate GCMacro
        pGeometricConstructionTransient.Public = True
        Call GCEvaluate(pGeometricConstructionTransient, pGeometricConstruction, -1, -1)
    ElseIf dAngle <= 185 Then
        ' create GCMacro transient
        Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "StraightEdge2", pPOM)
            
        ' feed GCMacro with inputs and parameters
        pGeometricConstructionTransient.Inputs("Support1").Add pGeometricConstructionMacrosOfExtractPorts1.Outputs("Support")(1)
        pGeometricConstructionTransient.Inputs("Support2").Add pGeometricConstructionMacrosOfExtractPorts2.Outputs("Support")(1)
        pGeometricConstructionTransient.Inputs("MemberAxis1").Add pGeometricConstructionMacrosOfExtractPorts1.Outputs("MemberAxis")(1)
        pGeometricConstructionTransient.Inputs("MemberAxis2").Add pGeometricConstructionMacrosOfExtractPorts2.Outputs("MemberAxis")(1)
        pGeometricConstructionTransient.Inputs("EdgePort1").Add pGeometricConstructionMacrosOfExtractPorts1.Outputs("CurveLeft")(1)
        pGeometricConstructionTransient.Inputs("EdgePort2").Add pGeometricConstructionMacrosOfExtractPorts2.Outputs("CurveRight")(1)
        pGeometricConstructionTransient.Parameter("WeldToe") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
        pGeometricConstructionTransient.Parameter("CutBack") = CDbl(pGeometricConstruction.Parameter(sCutBack1))
        pGeometricConstructionTransient.Parameter("CutBack2") = CDbl(pGeometricConstruction.Parameter(sCutBack2))
        pGeometricConstructionTransient.Parameter("RibbonHeight") = 2 * MemberPart_GetMaximumSize(oMember) 'dDepth / 2 'dDepth + 2 * pGeometricConstruction.Parameter("OverSize")
            
        ' evaluate GCMacro
        pGeometricConstructionTransient.Public = True
        Call GCEvaluate(pGeometricConstructionTransient, pGeometricConstruction, -1, -1)
    Else
        ' create GCMacro transient
        Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "FreeEdgeF13b", pPOM)
                
        ' feed GCMacro with inputs and parameters
        pGeometricConstructionTransient.Inputs("AxisLine1").Add pGeometricConstructionMacrosOfExtractPorts1.Outputs("CurveLeft")(1)
        pGeometricConstructionTransient.Inputs("AxisLine2").Add pGeometricConstructionMacrosOfExtractPorts2.Outputs("CurveRight")(1)
        pGeometricConstructionTransient.Parameter("WeldToe") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
        pGeometricConstructionTransient.Parameter("Fillet") = CDbl(pGeometricConstruction.Parameter("Fillet"))
        pGeometricConstructionTransient.Parameter("CutBack") = CDbl(pGeometricConstruction.Parameter(sCutBack1))
        pGeometricConstructionTransient.Parameter("CutBack2") = CDbl(pGeometricConstruction.Parameter(sCutBack2))
        pGeometricConstructionTransient.Parameter("RibbonHeight") = 2 * MemberPart_GetMaximumSize(oMember) 'dDepth / 2 'dDepth + 2 * pGeometricConstruction.Parameter("OverSize")
        pGeometricConstructionTransient.Parameter("Side") = 2
                
        ' evaluate GCMacro
        pGeometricConstructionTransient.Public = True
        Call GCEvaluate(pGeometricConstructionTransient, pGeometricConstruction, -1, -1)
    End If
    
    ' return result
    Set CreateGeometricConstructionMacroOfFreeEdge2 = pGeometricConstructionTransient
    Call DebugOut
End Function
Function CreateGeometricConstructionMacroOfTrimmingPoint(pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory, pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                                                          pGeometricConstructionMacrosOfExtractPorts As IJGeometricConstructionMacro, _
                                                          pGeometricConstructionMacrosOfFreeEdge1 As IJGeometricConstructionMacro, _
                                                          pGeometricConstructionMacrosOfFreeEdge2 As IJGeometricConstructionMacro) As IJGeometricConstructionMacro
    Call DebugIn("CreateGeometricConstructionMacroOfTrimmingPoint")
    
    ' create GCMacro transient
    Dim pGeometricConstructionTransient As IJGeometricConstruction
    Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "TrimmingPoint", pPOM)
        
    ' feed GCMacro with inputs
    pGeometricConstructionTransient.Inputs("MemberAxis").Add pGeometricConstructionMacrosOfExtractPorts.Outputs("MemberAxis")(1)
    pGeometricConstructionTransient.Inputs("Point1").Add pGeometricConstructionMacrosOfFreeEdge1.Outputs("PointX")(1)
    pGeometricConstructionTransient.Inputs("Point2").Add pGeometricConstructionMacrosOfFreeEdge2.Outputs("PointY")(1)
        
    ' evaluate GCMacro
    pGeometricConstructionTransient.Public = True
    pGeometricConstructionTransient.Evaluate
        
    ' return result
    Set CreateGeometricConstructionMacroOfTrimmingPoint = pGeometricConstructionTransient
    Call DebugOut
End Function
Function CreateGeometricConstructionMacroOfTrimmingPoint2(pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory, pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                                                          pGeometricConstructionMacrosOfExtractPorts As IJGeometricConstructionMacro, _
                                                          pGeometricConstructionMacrosOfFreeEdge1 As IJGeometricConstructionMacro, _
                                                          pGeometricConstructionMacrosOfFreeEdge2 As IJGeometricConstructionMacro) As IJGeometricConstructionMacro
    Call DebugIn("CreateGeometricConstructionMacroOfTrimmingPoint2")
    
    ' create GCMacro transient
    Dim pGeometricConstructionTransient As IJGeometricConstruction
    Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "TrimmingPoint2", pPOM)
        
    ' feed GCMacro with inputs
    pGeometricConstructionTransient.Inputs("MemberAxis").Add pGeometricConstructionMacrosOfExtractPorts.Outputs("MemberAxis")(1)
    pGeometricConstructionTransient.Inputs("Point1").Add pGeometricConstructionMacrosOfFreeEdge1.Outputs("PointX")(1)
    pGeometricConstructionTransient.Inputs("Point2").Add pGeometricConstructionMacrosOfFreeEdge2.Outputs("PointY")(1)
    pGeometricConstructionTransient.Parameter("WeldToe") = pGeometricConstruction.Parameter("WeldToe")
    pGeometricConstructionTransient.Parameter("SlopeRatio") = pGeometricConstruction.Parameter("SlopeRatio")
        
    ' evaluate GCMacro
    pGeometricConstructionTransient.Public = True
    pGeometricConstructionTransient.Evaluate
        
    ' return result
    Set CreateGeometricConstructionMacroOfTrimmingPoint2 = pGeometricConstructionTransient
    Call DebugOut
End Function
Function CreateGeometricConstructionMacroOfTrimmingPlane(pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory, pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                                                         oMember As Object, _
                                                         pGeometricConstructionMacroOfTrimmingPoint As IJGeometricConstructionMacro) As IJGeometricConstructionMacro
    Call DebugIn("CreateGeometricConstructionMacroOfTrimmingPlane")
    
    ' create GCMacro transient
    Dim pGeometricConstructionTransient As IJGeometricConstruction
    Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "TrimmingPlane", pPOM)
        
    ' feed GCMacro with inputs and parameters
    pGeometricConstructionTransient.Inputs("MemberPart").Add oMember
    pGeometricConstructionTransient.Inputs("Point").Add pGeometricConstructionMacroOfTrimmingPoint.Outputs("Point")(1)
    pGeometricConstructionTransient.Parameter("Extension") = CDbl(pGeometricConstruction.Parameter("Extension")) + CDbl(pGeometricConstruction.Parameter("WeldToe"))
        
    ' evaluate GCMacro
    pGeometricConstructionTransient.Public = True
    pGeometricConstructionTransient.Evaluate

    ' return result
    Set CreateGeometricConstructionMacroOfTrimmingPlane = pGeometricConstructionTransient
    Call DebugOut
End Function
Function CreateGeometricConstructionMacroOfSnipedSurfaces(pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory, pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                                                         oMember As Object, _
                                                         pGeometricConstructionMacroOfExtractPorts As IJGeometricConstructionMacro, _
                                                         pGeometricConstructionMacroOfTrimmingPoint As IJGeometricConstructionMacro) As IJGeometricConstructionMacro
    Call DebugIn("CreateGeometricConstructionMacroOfSnipedSurfaces")
    
    ' create GCMacro transient
    Dim pGeometricConstructionTransient As IJGeometricConstruction
    Set pGeometricConstructionTransient = CreateEntity(pGeometricConstructionFactory, "SnipedSurfaces", pPOM)
        
    ' feed GCMacro with inputs and parameters
    pGeometricConstructionTransient.Inputs("Support").Add pGeometricConstructionMacroOfExtractPorts.Outputs("Support")(1)
    pGeometricConstructionTransient.Inputs("CurveLeft").Add pGeometricConstructionMacroOfExtractPorts.Outputs("CurveLeft")(1)
    pGeometricConstructionTransient.Inputs("CurveRight").Add pGeometricConstructionMacroOfExtractPorts.Outputs("CurveRight")(1)
    pGeometricConstructionTransient.Inputs("MemberAxis").Add pGeometricConstructionMacroOfExtractPorts.Outputs("MemberAxis")(1)
    pGeometricConstructionTransient.Inputs("TrimmingPoint").Add pGeometricConstructionMacroOfTrimmingPoint.Outputs("Point")(1)
    pGeometricConstructionTransient.Parameter("WeldToe") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
    pGeometricConstructionTransient.Parameter("SlopeRatio") = CDbl(pGeometricConstruction.Parameter("SlopeRatio"))
    pGeometricConstructionTransient.Parameter("Extension") = CDbl(pGeometricConstruction.Parameter("Extension"))
    pGeometricConstructionTransient.Parameter("RibbonHeight") = 2 * MemberPart_GetMaximumSize(oMember)
        
    ' evaluate GCMacro
    pGeometricConstructionTransient.Public = True
    pGeometricConstructionTransient.Evaluate

    ' return result
    Set CreateGeometricConstructionMacroOfSnipedSurfaces = pGeometricConstructionTransient
    Call DebugOut
End Function
Public Sub CreatePseudoBoundaries(pPOM As IJDPOM, pGeometricConstruction As IJGeometricConstruction, _
                                   iCountOfSecondaryMembers As Integer, pElementsOfSecondaryMembers As IJElements, sKeysOfSecondaryMembers() As String, _
                                   iCountOfBoundaries As Integer, pElementsOfBoundaries As IJElements, ByRef sKeysOfBoundaries() As String)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "CreatePseudoBoundaries"
    Call DebugIn("CreatePseudoBoundaries")
    
    ' retrieve GCMacro
    Dim pGeometricConstructionMacro As IJGeometricConstructionMacro: Set pGeometricConstructionMacro = pGeometricConstruction
    
    ' retrieve the input index of the secondary membert parts
    Dim iInputIndexOfSecondaryMemberParts As Integer: Let iInputIndexOfSecondaryMemberParts = MacroDefinition_GetInputIndex(pGeometricConstruction.definition, "SecondaryMemberParts")
    
    ' instantiate GCFactory
    Dim pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory: Set pGeometricConstructionFactory = New GeometricConstructionEntitiesFactory
    
    ' compute dummy position
    Dim pPosition As IJDPosition
    If True Then
        Set pPosition = New DPosition
        Call pPosition.Set(0#, 0#, 0#)
    End If
    
    Let sComment = "loop on secondary members"
    Dim i As Integer
    For i = 1 To iCountOfSecondaryMembers
        ' set controlled input
        If True Then
            ' extract the axis port
            Dim oAxisPort As Object
            If True Then
                Dim oLineAxisPortExtractor0 As GeometricConstruction
                Set oLineAxisPortExtractor0 = CreateEntity(pGeometricConstructionFactory, "LineAxisPortExtractor", pPOM)
                oLineAxisPortExtractor0.Inputs("MemberPart").Add pElementsOfSecondaryMembers(i)
                oLineAxisPortExtractor0.Evaluate
                Set oAxisPort = oLineAxisPortExtractor0.ControlledInputs("Port")(1)
            End If
            
            ' update the corresponding controlled input
            If True Then
                Dim pElementsOfAxisPorts As IJElements: Set pElementsOfAxisPorts = pGeometricConstruction.ControlledInputs("AxisPort" + CStr(iInputIndexOfSecondaryMemberParts))
                If Not pElementsOfAxisPorts.Contains(oAxisPort) Then
                    Call pElementsOfAxisPorts.Add(oAxisPort, sKeysOfSecondaryMembers(i))
                End If
            End If
        End If
        
        ' build key
        Dim sKeyOfPseudoBoundary As String: Let sKeyOfPseudoBoundary = "PB" + CStr(iInputIndexOfSecondaryMemberParts) + "." + sKeysOfSecondaryMembers(i)
        ' assign a dummy position to each pseudo boundary
        pGeometricConstructionMacro.Output("PseudoBoundary", sKeyOfPseudoBoundary) = Point_FromPosition(pPosition)
    
         If iCountOfBoundaries > 0 Then
            Call Keys_RemoveKey(sKeysOfBoundaries, sKeyOfPseudoBoundary)
        End If
    Next
    
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = "Common" + "::" + sFunction
    Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
Function MemberPart_GetMaximumSize(oMemberPart As Object) As Double

    Dim dWidth As Double, dDepth As Double
    
    If Not TypeOf oMemberPart Is IJStiffenerSystem Then
        ' retrieve cross-section
        Dim pCrossSection As ISPSCrossSection
        If True Then
            If TypeOf oMemberPart Is ISPSMemberPartPrismatic Then
                Dim pMemberPartPrismatic As ISPSMemberPartPrismatic
                Set pMemberPartPrismatic = oMemberPart
                Set pCrossSection = pMemberPartPrismatic.CrossSection
            Else
                Set pCrossSection = oMemberPart
            End If
        End If
            
        ' retrieve nominal cross-section size
        If True Then
            Dim pPositionOfOrigin As IJDPosition: Set pPositionOfOrigin = Position_FromLine(MemberPart_GetLine(oMemberPart), 0)
            Call pCrossSection.GetNominalSectionSize(pPositionOfOrigin, dWidth, dDepth)
        End If
    Else
        dWidth = 1
        dDepth = 1
    End If
    
    ' return result
    If dWidth > dDepth Then Let MemberPart_GetMaximumSize = dWidth Else Let MemberPart_GetMaximumSize = dDepth
End Function
Sub GeometricConstructionMacro_CreateCoordinateSystemOfSupport(pGeometricConstructionMacro As IJGeometricConstructionMacro, pPOM As IJDPOM, pGCGeomFactory As IJGCGeomFactory, pGCGeomFactory2 As IJGCGeomFactory2, pPositionOfNode As IJDPosition)
    ' get the CS without controlling its origin
    Dim pLocalCoordinateSystem As IJLocalCoordinateSystem
    Set pLocalCoordinateSystem = pGCGeomFactory2.CSByPlane(pPOM, pGeometricConstructionMacro.Outputs("Support")(1), Nothing)
    
    ' compute position of projected node
    Dim pPositionOfProjectedNode As IJDPosition:
    Set pPositionOfProjectedNode = pPositionOfNode.Offset(Vector_Scale(pLocalCoordinateSystem.ZAxis, -pPositionOfNode.Subtract(pLocalCoordinateSystem.Position).Dot(pLocalCoordinateSystem.ZAxis)))
    
    ' set the macro output
    pGeometricConstructionMacro.Output("CoordinateSystem", 1) = pGCGeomFactory2.CSByPlane(pPOM, pGeometricConstructionMacro.Outputs("Support")(1), Point_FromPosition(pPositionOfProjectedNode))
End Sub
Sub GeometricConstructionMacro_CreateNode(pGeometricConstructionMacro As IJGeometricConstructionMacro, pPositionOfNode As IJDPosition)
    ' set the macro output
    pGeometricConstructionMacro.Output("Node", 1) = Point_FromPosition(pPositionOfNode)
End Sub
Function GeometricConstructionMacro_RetrievePositionOfNode(pGeometricConstructionMacro As IJGeometricConstructionMacro) As IJDPosition
    ' prepare result
    Dim pPositionOfNode As IJDPosition
        
    If pGeometricConstructionMacro.Outputs("Node").Count = 1 Then
        Set pPositionOfNode = Position_FromPoint(pGeometricConstructionMacro.Outputs("Node")(1))
    Else
        ' the "Node" output might not exist for old APSs
        Set pPositionOfNode = New DPosition
    End If
    
    ' return result
    Set GeometricConstructionMacro_RetrievePositionOfNode = pPositionOfNode
End Function
Sub GeometricConstruction_MigrateInputs(ByVal MyGC As IJGeometricConstruction, ByVal pMigrateHelper As IJGCMigrateHelper, ByVal sName As String, ByVal pPositionOfNode As IJDPosition)
    Call DebugIn("GeometricConstruction_MigrateInputs")
    Call DebugInput("Name", sName)
    Call DebugInput("PositionOfNode", pPositionOfNode)
           
    Dim i As Integer
    For i = 1 To MyGC.Inputs(sName).Count
        Call DebugMsg("Processing input #" + CStr(i) + "/" + CStr(MyGC.Inputs(sName).Count))
        
        ' retrieve the current input
        Dim oMemberPartOriginal As Object: Set oMemberPartOriginal = MyGC.Inputs(sName)(i)
        Call DebugValue("MemberPartOriginal", oMemberPartOriginal)
        
        ' get the positions of the joints and ends of the replacing member part
        Dim pPositionsAtJointsAndEndsOfMemberPartOriginal() As IJDPosition:
        Let pPositionsAtJointsAndEndsOfMemberPartOriginal = MemberPart_GetPositionsAtJointsAndEnds(oMemberPartOriginal)
     
        ' retrieve the replacing objects, if they exist
        Dim pElementsOfReplacingMemberParts As IJElements
        If True Then
            Dim bIsDeleted As Boolean
            Call pMigrateHelper.ObjectsReplacing(oMemberPartOriginal, pElementsOfReplacingMemberParts, bIsDeleted)
            Call DebugValue("IsDeleted", bIsDeleted)
        End If
        
        ' check if current input has been replaced
        If Not pElementsOfReplacingMemberParts Is Nothing Then
            'find the end of the original member part closer to the node
            Dim dDistanceMini As Double: Let dDistanceMini = 1000000#
            Dim k As Integer, k2 As Integer
            For k = 2 To 3
                Dim dDistance As Double: Let dDistance = pPositionOfNode.DistPt(pPositionsAtJointsAndEndsOfMemberPartOriginal(k))
                Call DebugMsg("distance for end (" + CStr(k) + ")= " + CStr(Round(dDistance, 6)))
                If dDistance < dDistanceMini Then
                    Let dDistanceMini = dDistance
                    Let k2 = k
                End If
            Next
            Call DebugMsg("coincidence found for k2= " + CStr(k2))
            
            Let dDistanceMini = 1000000#
            Dim j As Integer: Dim j2 As Integer
            For j = 1 To pElementsOfReplacingMemberParts.Count
                Call DebugMsg("Processing replacing member #" + CStr(j) + "/" + CStr(pElementsOfReplacingMemberParts.Count))
                
               ' retrieve the replacing member part
                Dim oMemberPartReplacing As Object: Set oMemberPartReplacing = pElementsOfReplacingMemberParts(j)
                Call DebugValue("MemberPartReplacing", oMemberPartReplacing)
            
                ' get the positions of the joints and ends of the replacing member part
                Dim pPositionsAtJointsAndEndsOfMemberPartReplacing() As IJDPosition:
                Let pPositionsAtJointsAndEndsOfMemberPartReplacing = MemberPart_GetPositionsAtJointsAndEnds(oMemberPartReplacing)
                
                Let dDistance = pPositionOfNode.DistPt(pPositionsAtJointsAndEndsOfMemberPartReplacing(k2))
                Call DebugMsg("distance for end (" + CStr(k2) + ")= " + CStr(Round(dDistance, 6)))
                If dDistance < dDistanceMini Then
                    Let dDistanceMini = dDistance
                    Let j2 = j
                End If
            Next
            
            Call DebugMsg("coincidence found for j2= " + CStr(j2))
            Call Elements_ReplaceElementWithSameKey(MyGC.Inputs(sName), oMemberPartOriginal, pElementsOfReplacingMemberParts(j2))
            
            ' disconnect corresponding ports
            Dim pElementsOfControlledInputs As IJElements: Set pElementsOfControlledInputs = GeometricConstruction_GetElementsOfAllControlledInputsByNameWithNameAsKey(MyGC, "")
            For k = 1 To pElementsOfControlledInputs.Count
                Dim oControlledInput As Object: Set oControlledInput = pElementsOfControlledInputs(k)
                If TypeOf oControlledInput Is IJPort Then
                    Dim pPort As IJPort: Set pPort = oControlledInput
                    Dim oConnectable As Object: Set oConnectable = pPort.Connectable
                    If TypeOf oConnectable Is IJPlateSystem Then
                        Set oConnectable = PlateSystem_GetBuiltUp(oConnectable)
                    End If
                    If oConnectable Is oMemberPartOriginal Then
                        Dim sKey As String: Let sKey = pElementsOfControlledInputs.GetKey(oControlledInput)
                        Call DebugValue("Disconnect port with key", sKey)
                        
                        Dim iDelimiter As Integer: Let iDelimiter = InStr(1, sKey, ".", vbTextCompare)
                        Dim sNameOfControlledInput As String
                        Dim sKeyOfControlledInput As String
                        If iDelimiter > 0 Then
                            Let sNameOfControlledInput = Mid(sKey, 1, iDelimiter - 1)
                            Let sKeyOfControlledInput = Mid(sKey, iDelimiter + 1)
                        Else
                            Let sNameOfControlledInput = sKey
                            Let sKeyOfControlledInput = ""
                        End If
                        If sKeyOfControlledInput <> "" Then
                            MyGC.ControlledInputs(sNameOfControlledInput).Remove (sKeyOfControlledInput)
                        Else
                            MyGC.ControlledInputs(sNameOfControlledInput).Remove (1)
                        End If
                    End If
                End If
            Next
            
            Call Elements_RenameKeyOfElement(MyGC.ControlledInputs("AdvancedPlateSystem"), 1, "JustMigrated")
        End If
    Next
            
    ' clear MemberSystemReference property
    Call MemberSystemReference_Set(Nothing)
    
    'Clear CTL_FLAG_SKIP_EVALUATE in case of migrate in MDR to get GCCompute error in TDL
    Dim oControlFlags As IJControlFlags: Set oControlFlags = MyGC
    If oControlFlags.ControlFlags(CTL_FLAG_SKIP_EVALUATE) = CTL_FLAG_SKIP_EVALUATE Then oControlFlags.ControlFlags(CTL_FLAG_SKIP_EVALUATE) = 0

    Call DebugOut
End Sub
Public Sub GetTrimmingPlanes(ByVal pPOM As IJDPOM, ByVal oSupportingMember1 As Object, ByVal oSupportingMember2 As Object, ByVal oUpperPlate As Object, ByVal oBelowPlate As Object, ByRef oTrimmingPlane1 As Object, ByRef oTrimmingPlane2 As Object)
    
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim oFacePortExtractor01 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor01 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "0001-FacePortExtractor0")
    oFacePortExtractor01.Inputs("Connectable").Add oUpperPlate, "1"
    oFacePortExtractor01.Parameter("Offset") = 0#
    oFacePortExtractor01.Parameter("GeometrySelector") = 2
    oFacePortExtractor01.Evaluate

    Dim oFacePortExtractor02 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor02 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "0002-FacePortExtractor0")
    oFacePortExtractor02.Inputs("Connectable").Add oBelowPlate, "2"
    oFacePortExtractor02.Parameter("Offset") = 0#
    oFacePortExtractor02.Parameter("GeometrySelector") = 2
    oFacePortExtractor02.Evaluate

    Dim oAxisPortExtractor23 As SP3DGeometricConstruction.GeometricConstruction
    Set oAxisPortExtractor23 = oGCFactory.CreateEntity("AxisPortExtractor2", pPOM, "0003-AxisPortExtractor2")
    oAxisPortExtractor23.Inputs("Connectable").Add oSupportingMember1, "1"
    oAxisPortExtractor23.Parameter("GeometrySelector") = 2
    oAxisPortExtractor23.Evaluate


    Dim oAxisPortExtractor24 As SP3DGeometricConstruction.GeometricConstruction
    Set oAxisPortExtractor24 = oGCFactory.CreateEntity("AxisPortExtractor2", pPOM, "0004-AxisPortExtractor2")
    oAxisPortExtractor24.Inputs("Connectable").Add oSupportingMember2, "1"
    oAxisPortExtractor24.Parameter("GeometrySelector") = 2
    oAxisPortExtractor24.Evaluate
    
    Dim oCSByPlane4 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane4 = oGCFactory.CreateEntity("CSByPlane", pPOM, "0004-CSByPlane")
    oCSByPlane4.Inputs("Plane").Add oFacePortExtractor01, "1"
    oCSByPlane4.Evaluate

    Dim oCurveByCurves6 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves6 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "0006-CurveByCurves")
    oCurveByCurves6.Inputs("Curves").Add oAxisPortExtractor23, "1"
    oCurveByCurves6.Inputs("Curves").Add oAxisPortExtractor24, "2"
    oCurveByCurves6.Parameter("ConstructionSurface") = 0
    oCurveByCurves6.Evaluate
    
    Dim oPointAtCurveMiddle7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle7 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0007-PointAtCurveMiddle")
    oPointAtCurveMiddle7.Inputs("Curve").Add oAxisPortExtractor23, "1"
    oPointAtCurveMiddle7.Evaluate
    
    Dim oLineFromCS5 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS5 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0005-LineFromCS")
    oLineFromCS5.Inputs("CoordinateSystem").Add oCSByPlane4, "1"
    oLineFromCS5.Parameter("LookingAxis") = 3
    oLineFromCS5.Parameter("Length") = 1#
    oLineFromCS5.Parameter("CSOrientation") = 1
    oLineFromCS5.Parameter("LineJustification") = 1
    oLineFromCS5.Evaluate

    Dim oCSByLines6 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines6 = oGCFactory.CreateEntity("CSByLines", pPOM, "0006-CSByLines")
    oCSByLines6.Inputs("AxisLine1").Add oAxisPortExtractor23, "1"
    oCSByLines6.Inputs("AxisLine2").Add oLineFromCS5, "2"
    oCSByLines6.Parameter("AxesRoles") = 1
    oCSByLines6.Parameter("CSOrientation") = 1
    oCSByLines6.Parameter("TrackFlag") = 1
    oCSByLines6.Evaluate

    Dim oPointsAsRangeBox7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointsAsRangeBox7 = oGCFactory.CreateEntity("PointsAsRangeBox", pPOM, "0007-PointsAsRangeBox")
    oPointsAsRangeBox7.Inputs("CoordinateSystem").Add oCSByLines6, "1"
    oPointsAsRangeBox7.Inputs("Geometries").Add oFacePortExtractor01, "1"
    oPointsAsRangeBox7.Inputs("Geometries").Add oFacePortExtractor02, "2"
    oPointsAsRangeBox7.Evaluate

    Dim oPointByProjectOnCurve10 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnCurve10 = oGCFactory.CreateEntity("PointByProjectOnCurve", pPOM, "0010-PointByProjectOnCurve")
    oPointByProjectOnCurve10.Inputs("Point").Add oPointsAsRangeBox7.Output("Points", 1), "1"
    oPointByProjectOnCurve10.Inputs("Curve").Add oCurveByCurves6, "2"
    oPointByProjectOnCurve10.Evaluate

    Dim oPointByProjectOnCurve11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnCurve11 = oGCFactory.CreateEntity("PointByProjectOnCurve", pPOM, "0011-PointByProjectOnCurve")
    oPointByProjectOnCurve11.Inputs("Point").Add oPointsAsRangeBox7.Output("Points", 2), "1"
    oPointByProjectOnCurve11.Inputs("Curve").Add oCurveByCurves6, "1"
    oPointByProjectOnCurve11.Evaluate
    
    Dim oPlaneByPointNormal8 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal8 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0008-PlaneByPointNormal")
    oPlaneByPointNormal8.Inputs("Point").Add oPointByProjectOnCurve10, "1"
    oPlaneByPointNormal8.Inputs("Line").Add oAxisPortExtractor23, "1"
    oPlaneByPointNormal8.Parameter("Range") = 3#
    oPlaneByPointNormal8.Evaluate

    Dim oPlaneByPointNormal9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal9 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0009-PlaneByPointNormal")
    oPlaneByPointNormal9.Inputs("Point").Add oPointByProjectOnCurve11, "2"
    oPlaneByPointNormal9.Inputs("Line").Add oAxisPortExtractor23, "1"
    oPlaneByPointNormal9.Parameter("Range") = 3#
    oPlaneByPointNormal9.Evaluate

    Dim oSurfFromGType10 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType10 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0010-SurfFromGType")
    oSurfFromGType10.Inputs("Surface").Add oPlaneByPointNormal8, "1"
    oSurfFromGType10.Evaluate

    Dim oSurfFromGType11 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType11 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0011-SurfFromGType")
    oSurfFromGType11.Inputs("Surface").Add oPlaneByPointNormal9, "2"
    oSurfFromGType11.Evaluate
    
    ' Now determine which plane is the closest from the mid point
    Dim dPlane1, dPlane2 As Double
    dPlane1 = Distance_PointToPlane(oPointAtCurveMiddle7, oPlaneByPointNormal8)
    dPlane2 = Distance_PointToPlane(oPointAtCurveMiddle7, oPlaneByPointNormal9)
    
    If (dPlane1 < dPlane2) Then
        Set oTrimmingPlane1 = oSurfFromGType10
        Set oTrimmingPlane2 = oSurfFromGType11
    Else
        Set oTrimmingPlane1 = oSurfFromGType11
        Set oTrimmingPlane2 = oSurfFromGType10
    End If
End Sub
Public Sub GetTrimmingPlaneAndNodePosition(ByVal pPOM As IJDPOM, ByVal oSupportingMember As Object, ByVal oUpperPlate As Object, ByVal oBelowPlate As Object, ByRef oTrimmingPlane As Object, ByRef oPosition As IJDPosition)
    
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oFacePortExtractor01 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor01 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "0001-FacePortExtractor0")
    oFacePortExtractor01.Inputs("Connectable").Add oUpperPlate, "1"
    oFacePortExtractor01.Parameter("Offset") = 0#
    oFacePortExtractor01.Parameter("GeometrySelector") = 2
    oFacePortExtractor01.Evaluate

    Dim oFacePortExtractor02 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor02 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "0002-FacePortExtractor0")
    oFacePortExtractor02.Inputs("Connectable").Add oBelowPlate, "1"
    oFacePortExtractor02.Parameter("Offset") = 0#
    oFacePortExtractor02.Parameter("GeometrySelector") = 2
    oFacePortExtractor02.Evaluate

    Dim oAxisPortExtractor23 As SP3DGeometricConstruction.GeometricConstruction
    Set oAxisPortExtractor23 = oGCFactory.CreateEntity("AxisPortExtractor2", pPOM, "0003-AxisPortExtractor2")
    oAxisPortExtractor23.Inputs("Connectable").Add oSupportingMember, "1"
    oAxisPortExtractor23.Parameter("GeometrySelector") = 2
    oAxisPortExtractor23.Evaluate

    Dim oCSByPlane4 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane4 = oGCFactory.CreateEntity("CSByPlane", pPOM, "0004-CSByPlane")
    oCSByPlane4.Inputs("Plane").Add oFacePortExtractor01, "1"
    oCSByPlane4.Evaluate

    Dim oPointAtCurveMiddle5 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle5 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0005-PointAtCurveMiddle")
    oPointAtCurveMiddle5.Inputs("Curve").Add oAxisPortExtractor23, "1"
    oPointAtCurveMiddle5.Evaluate

    Dim oLineFromCS6 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS6 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0006-LineFromCS")
    oLineFromCS6.Inputs("CoordinateSystem").Add oCSByPlane4, "1"
    oLineFromCS6.Parameter("LookingAxis") = 3
    oLineFromCS6.Parameter("Length") = 1#
    oLineFromCS6.Parameter("CSOrientation") = 1
    oLineFromCS6.Parameter("LineJustification") = 1
    oLineFromCS6.Evaluate

    Dim oCSByLines7 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines7 = oGCFactory.CreateEntity("CSByLines", pPOM, "0007-CSByLines")
    oCSByLines7.Inputs("AxisLine1").Add oAxisPortExtractor23, "1"
    oCSByLines7.Inputs("AxisLine2").Add oLineFromCS6, "1"
    oCSByLines7.Parameter("AxesRoles") = 1
    oCSByLines7.Parameter("CSOrientation") = 1
    oCSByLines7.Parameter("TrackFlag") = 1
    oCSByLines7.Evaluate

    Dim oLineFromCS8 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS8 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0008-LineFromCS")
    oLineFromCS8.Inputs("CoordinateSystem").Add oCSByLines7, "1"
    oLineFromCS8.Parameter("LookingAxis") = 1
    oLineFromCS8.Parameter("Length") = 1#
    oLineFromCS8.Parameter("CSOrientation") = 1
    oLineFromCS8.Parameter("LineJustification") = 1
    oLineFromCS8.Evaluate

    Dim oPointsAsRangeBox9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointsAsRangeBox9 = oGCFactory.CreateEntity("PointsAsRangeBox", pPOM, "0009-PointsAsRangeBox")
    oPointsAsRangeBox9.Inputs("CoordinateSystem").Add oCSByLines7, "1"
    oPointsAsRangeBox9.Inputs("Geometries").Add oFacePortExtractor01, "1"
    oPointsAsRangeBox9.Inputs("Geometries").Add oFacePortExtractor02, "2"
    oPointsAsRangeBox9.Evaluate
    
    'It's time to determine the right point to project
    Dim dDistance1 As Double: Let dDistance1 = 0
    Dim dDistance2 As Double: Let dDistance2 = 0
    
    Dim pPosition1 As IJDPosition: Set pPosition1 = Position_FromPoint(oPointsAsRangeBox9.Output("Points", 1))
    Dim pPosition2 As IJDPosition: Set pPosition2 = Position_FromPoint(oPointsAsRangeBox9.Output("Points", 2))
    Dim pPositionMiddle As IJDPosition: Set pPositionMiddle = Position_FromPoint(oPointAtCurveMiddle5)
    
    Let dDistance1 = pPositionMiddle.DistPt(pPosition1)
    Let dDistance2 = pPositionMiddle.DistPt(pPosition2)
    
    Dim index As Long
    If dDistance1 < dDistance2 Then
        index = 1
    Else
        index = 2
    End If

    Dim oPointByProjectOnCurve10 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnCurve10 = oGCFactory.CreateEntity("PointByProjectOnCurve", pPOM, "0010-PointByProjectOnCurve")
    oPointByProjectOnCurve10.Inputs("Point").Add oPointsAsRangeBox9.Output("Points", index), "1"
    oPointByProjectOnCurve10.Inputs("Curve").Add oAxisPortExtractor23, "1"
    oPointByProjectOnCurve10.Evaluate

    'Compute position
    Dim oPointAtCurveStart As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "0005-PointAtCurveStart")
    oPointAtCurveStart.Inputs("Curve").Add oAxisPortExtractor23, "1"
    oPointAtCurveStart.Evaluate
    
    Dim oPointAtCurveEnd As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "0005-PointAtCurveEnd")
    oPointAtCurveEnd.Inputs("Curve").Add oAxisPortExtractor23, "1"
    oPointAtCurveEnd.Evaluate
    
    Set pPosition1 = Position_FromPoint(oPointAtCurveStart)
    Set pPosition2 = Position_FromPoint(oPointAtCurveEnd)
    Dim pProjectedPosition As IJDPosition: Set pProjectedPosition = Position_FromPoint(oPointByProjectOnCurve10)
    
    Let dDistance1 = pProjectedPosition.DistPt(pPosition1)
    Let dDistance2 = pProjectedPosition.DistPt(pPosition2)
    
    If dDistance1 < dDistance2 Then
        Set oPosition = pPosition1
    Else
        Set oPosition = pPosition2
    End If
    
    Dim oPlaneByPointNormal11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal11 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0011-PlaneByPointNormal")
    oPlaneByPointNormal11.Inputs("Point").Add oPointByProjectOnCurve10, "3"
    oPlaneByPointNormal11.Inputs("Line").Add oLineFromCS8, "1"
    oPlaneByPointNormal11.Parameter("Range") = 3#
    oPlaneByPointNormal11.Evaluate

    Dim oSurfFromGType12 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType12 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0012-SurfFromGType")
    oSurfFromGType12.Inputs("Surface").Add oPlaneByPointNormal11, "1"
    oSurfFromGType12.Evaluate
 
    Set oTrimmingPlane = oSurfFromGType12
End Sub
