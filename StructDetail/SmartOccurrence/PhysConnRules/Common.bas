Attribute VB_Name = "Common"
Option Explicit

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\PhysConnRules\Common.bas"

Public Function GetRefSide(oObject As Object) As String

    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim sMoldedSide As String
    Dim typProfileSystemType As StructProfileType

    Dim pHelper As New StructDetailObjects.Helper

    Select Case pHelper.ObjectType(oObject)
        Case SDOBJECT_PLATE

            Dim oWeldPlate As New StructDetailObjects.PlatePart
            Set oWeldPlate.object = oObject
            sMoldedSide = oWeldPlate.MoldedSide

        Case SDOBJECT_STIFFENER
        'get the secondary orientation of the profile part
            Dim oWeldProfile As New StructDetailObjects.ProfilePart
            Set oWeldProfile.object = oObject
            Dim oWeldProfileSecond As StructMoldedOrientation
            Dim bflag As Boolean
            oWeldProfile.Get_SecondOrientation oWeldProfileSecond, bflag
            
            'get the profile type of the profile part
             typProfileSystemType = oWeldProfile.ProfileType
            
            Select Case typProfileSystemType
                Case sptVertical
                    Select Case oWeldProfileSecond
                        Case InboardOrient
                            sMoldedSide = "Outboard"
                        Case OutboardOrient
                            sMoldedSide = "Inboard"
                        Case PortOrient
                            sMoldedSide = "Starboard"
                        Case StarboardOrient
                            sMoldedSide = "Port"
                    End Select
                Case sptLongitudinal
                    Select Case oWeldProfileSecond
                        Case InboardOrient
                            sMoldedSide = "Outboard"
                        Case OutboardOrient
                            sMoldedSide = "Inboard"
                        Case PortOrient
                            sMoldedSide = "Starboard"
                        Case StarboardOrient
                            sMoldedSide = "Port"
                    End Select
                Case sptTransversal
                    Select Case oWeldProfileSecond
                        Case ForeOrient
                            sMoldedSide = "Aft"
                        Case AftOrient
                            sMoldedSide = "Fore"
                        Case BelowOrient
                            sMoldedSide = "Above"
                        Case AboveOrient
                            sMoldedSide = "Below"
                    End Select
                    
                ' Add Case for Edge Reinforcement type
                ' Set "TowardPlate" or "FromPlate" as the ReferenceSide
                ' the ReferenceSide value is also in:
                ' see: StructDetail\Data\Symbols\WeldSymbols\ButtWeld.cls(817):
                Case sptEdgeReinforcement
                    Select Case oWeldProfileSecond
                        Case ER_TowardPlate
                            sMoldedSide = "FromPlate"
                        Case ER_FromPlate
                            sMoldedSide = "TowardPlate"
                    End Select
                    
                End Select
        Case SDOBJECT_BEAM
            sMoldedSide = "Web_Left"
    End Select
    
'    End If
    
    GetRefSide = sMoldedSide

  Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetRefSide", strError).Number
End Function

'------------------------------------------------------------------------------------------------------------
' Procedure (Function):
'     DegToRad (Double)
'
' Description:
'     Converts an angle measure in degrees to its equivalent in radians.
'
' Arguments:
'     dAngle    Double    The angle measure in degrees.
'------------------------------------------------------------------------------------------------------------
Public Function DegToRad(dAngle As Double) As Double

    Const PI As Double = 3.141592654

    DegToRad = dAngle * PI / 180  'Radians=Degrees*Pi/180
End Function

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     Get_ParameterRuleData
'
' Description:
'   Retrieves Parameter Rule data required by
'       all Tee Welds, Chain Weld, and ZigZag Weld Parameter Rules
'
'   The Data is retrieved from
'       the "Standard" Tee Weld selectors
'   Or
'       the "Chamfer" Tee Weld selectors
'
' Arguments:
'------------------------------------------------------------------------------------------------------------
Public Sub Get_ParameterRuleData(pPRL As IJDParameterLogic, _
                                 sStandardItemName As String, _
                                 sClassSociety As String, _
                                 sCategory As String, _
                                 sBevelMethod As String, _
                                 dThickness1 As Double, _
                                 dThickness2 As Double)
    On Error GoTo ErrorHandler
    Dim dChamferThickness As Double
  
    ' Get Class Arguments
    Dim oPhysConn As StructDetailObjects.PhysicalConn
    Set oPhysConn = New StructDetailObjects.PhysicalConn
    Set oPhysConn.object = pPRL.SmartOccurrence
  
    ' Check where the Parameter Rule data is to be Retrieve From
    ' Is this a "Standard" Physical Connection or a Chamfer Physical Connection
    If LCase(Trim(pPRL.SmartItem.Name)) = LCase(Trim(sStandardItemName)) Then
        ' Get data from "Standard" selector
        sCategory = pPRL.SelectorAnswer("PhysConnRules.TeeWeldSel", "Category")
        sBevelMethod = pPRL.SelectorAnswer("PhysConnRules.TeeWeldSel", "BevelAngleMethod")
        sClassSociety = pPRL.SelectorAnswer("PhysConnRules.TeeWeldSel", "ClassSociety")
        dThickness1 = oPhysConn.Object1Thickness
    Else
        ' Get data from "Chamfer" selector
        sCategory = pPRL.SelectorAnswer("PhysConnRules.ChamferTeeWeldSel", "Category")
        sBevelMethod = pPRL.SelectorAnswer("PhysConnRules.ChamferTeeWeldSel", "BevelAngleMethod")
        sClassSociety = pPRL.SelectorAnswer("PhysConnRules.ChamferTeeWeldSel", "ClassSociety")
        dChamferThickness = pPRL.SelectorAnswer("PhysConnRules.ChamferTeeWeldSel", "ChamferThickness")
        dThickness1 = dChamferThickness
    End If

    dThickness2 = oPhysConn.Object2Thickness

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "Get_ParameterRuleData").Number
End Sub

Public Function GetParentPhysicalConnectionSelectorDefinition(ByRef oPCParent As Object) As IJDSymbolDefinition
    
    On Error GoTo ErrorHandler
    
    If Not oPCParent Is Nothing Then
    
        If TypeOf oPCParent Is IJStructPhysicalConnection Then
    
            If TypeOf oPCParent Is IJSmartOccurrence Then
                
                Dim oSmartItem As IJSmartItem
                Dim oSmartParent As Object
                Dim oSmartOccurrence As IJSmartOccurrence
                
                Set oSmartOccurrence = oPCParent
                Set oSmartItem = oSmartOccurrence.ItemObject
                Set oSmartParent = oSmartItem.Parent
                
                If TypeOf oSmartParent Is IJSmartClass Then
                    
                    Dim osmartclass As IJSmartClass
                    Dim oSelectorDefinition As IJDSymbolDefinition
                    
                    Set osmartclass = oSmartParent
                    Set osmartclass = osmartclass.Parent
                    If Not osmartclass Is Nothing Then
                        
                        Set oSelectorDefinition = osmartclass.SelectionRuleDef
                        
                        If Not oSelectorDefinition Is Nothing Then
                            
                            Set GetParentPhysicalConnectionSelectorDefinition = oSelectorDefinition
                            
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetParentPhysicalConnectionSelectorDefinition").Number
End Function
Public Sub SetAnswerFromParentPhysicalConnection(ByRef oSymbolDefinition As IJDSymbolDefinition, ByVal AnswerName As String)
    
    On Error GoTo ErrorHandler
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
        
    ' Get the parent PC
    Dim oPCParent As Object
    Dim oSystemChild As IJSystemChild
    
    Set oSystemChild = pSL.SmartOccurrence
    
    Set oPCParent = oSystemChild.GetParent
    
    If Not oPCParent Is Nothing Then
        Dim oSelectorDefinition As IJDSymbolDefinition
        
        Set oSelectorDefinition = GetParentPhysicalConnectionSelectorDefinition(oPCParent)
        
        If Not oSelectorDefinition Is Nothing Then
            Dim oCommonHelper As DefinitionHlprs.CommonHelper
            Set oCommonHelper = New DefinitionHlprs.CommonHelper
            
            pSL.Answer(AnswerName) = oCommonHelper.GetAnswer(oPCParent, oSelectorDefinition, AnswerName)
            
        End If
        
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetAnswerFromParentPhysicalConnection").Number
End Sub
