VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MI_DGC_PAIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MI_DGC_PAIR.cls
'   ProgID:         HS_Hilti_MI.MI_DGC_PAIR
'   Author:         SN
'   Creation Date:  October 23, 2007

'   Description:
'       MI-DGC 90/120 to MI
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts  
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "MI_DGC_PAIR"  'Used for error messages
Private Const m_progID = "HS_Hilti_MI.MI_DGC_PAIR"

Private m_SymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
Private m_SymbolDefinition As IMSSymbolEntities.DSymbolDefinition
Private m_outputColl As IJDOutputCollection
Dim oCodeListData As IJDCodeListMetaData

Dim sPartNumber As String

Implements IJDUserSymbolServices
Implements IJHgrSymbolBOMServices


'********************************************************************
' Function: IJDUserSymbolServices_EditOccurence
' This method is called each time the user wants to edit the symbol
'   occurrence (Modify its parameter values or its representation for
'   instance). Prefer to register a custom command to fulfill this
'   functionality, into the symbol definition.
' Currently obsolete (as of Oct 99).
'********************************************************************
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
        IJDUserSymbolServices_EditOccurence = False
End Function

'********************************************************************
' Sub: IJDUserSymbolServices_InitializeSymbolDefinition
' This procedure is called to set up the Inputs and outputs of the Symbol
'
' Parameters: pSymbolDefinition - IMSSymbolEntities.IJDSymbolDefinition
'********************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrHandler

        'Inputs
        Const iNumInputs As Integer = 13
        Dim sInputName(iNumInputs) As String
        Dim sDataType(iNumInputs) As String
        
        'outputs
        Const iNumOutputs As Integer = 10
        Dim sOutputName(iNumOutputs) As String
        
        'No Caching
        'pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED

        'Inputs
        sInputName(2) = "PlateThick"
        sDataType(2) = "Double"

        sInputName(3) = "ItemNo"
        sDataType(3) = "String"

        sInputName(4) = "Finish"
        sDataType(4) = "Long"
        
        sInputName(5) = "Group"
        sDataType(5) = "String"
        
        sInputName(6) = "W"
        sDataType(6) = "Double"
        
        sInputName(7) = "D"
        sDataType(7) = "Double"
        
        sInputName(8) = "PlateDepth"
        sDataType(8) = "Double"
        
        sInputName(9) = "PlateWidth"
        sDataType(9) = "Double"
                
        sInputName(10) = "RodDia"
        sDataType(10) = "Double"
        
        sInputName(11) = "RodOffset"
        sDataType(11) = "Double"
                
        sInputName(12) = "PortOffset"
        sDataType(12) = "Double"
        
        sInputName(13) = "NumberClamps"
        sDataType(13) = "Long"
                
        'Outputs
        sOutputName(1) = "StructureInt"
        sOutputName(2) = "StructureExt"
        sOutputName(3) = "Plate"
        sOutputName(4) = "RightRod"
        sOutputName(5) = "LeftRod"
        sOutputName(6) = "BottomRod"
        sOutputName(7) = "Plate2"
        sOutputName(8) = "RightRod2"
        sOutputName(9) = "LeftRod2"
        sOutputName(10) = "BottomRod2"
        SetupInputsAndOutputs pSymbolDefinition, iNumInputs, iNumOutputs, sInputName(), sDataType(), sOutputName(), m_progID

        Exit Sub
ErrHandler:
        Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

'********************************************************************
' Function: IJDUserSymbolServices_GetDefinitionName
' Return the Definition Name
'
' Parameters: definitionParamaters - Variant
' Returns: String
'********************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
        IJDUserSymbolServices_GetDefinitionName = m_progID
End Function

'********************************************************************
' Function: IJDUserSymbolServices_InstanciateDefinition
' Instantiates a persistent symbol definition object and initializes
'   it for the first time.
'
' Parameters: CodeBase - String, defParamaters - Variant, ActiveConnection - Object
' Returns: Object
'********************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
        Set m_SymbolDefinition = m_SymbolFactory.CreateEntity(Definition, ActiveConnection)
        IJDUserSymbolServices_InitializeSymbolDefinition m_SymbolDefinition
        Set IJDUserSymbolServices_InstanciateDefinition = m_SymbolDefinition
        m_SymbolDefinition.ProgId = m_progID
        m_SymbolDefinition.CodeBase = CodeBase
        m_SymbolDefinition.name = m_SymbolDefinition.ProgId
End Function

'********************************************************************
' Sub: IJDUserSymbolServices_InvokeRepresentation
' It handles the recompute of  the symbol graphics outputs given the
'   specified representation and set of parameters.
'
' Parameters: sblOcc - Object, repName - String, outputcoll - Object, arrayOfInputs()
'********************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
        Set m_outputColl = outputcoll
        Set oCodeListData = sblOcc
        If StrComp(repName, "Symbolic") = 0 Then
            Call Physical(arrayOfInputs, sblOcc)
        End If
        Set oCodeListData = Nothing
End Sub

'********************************************************************
' Function: Physical
' Handles the recompute of the Physical representation
'
' Parameters: arrayOfInputs()
' Returns: Nothing
'********************************************************************
Private Function Physical(ByRef arrayOfInputs(), ByVal sblOcc As Object)
Const METHOD = "Physical"
        On Error GoTo ErrHandler
               
        Dim oSupportComp As IJHgrSupportComponent
        Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr

        Set oSupportComp = sblOcc
        oSupportComp.GetOccAssembly oICH
        
        Dim oPartOcc As PartOcc
        Dim oPart As IJDPart
        Set oPartOcc = oSupportComp ' The part occurence
        oPartOcc.GetPart oPart
        Dim sPartNumber As String
        sPartNumber = oPart.PartNumber
        
        Set oSupportComp = Nothing
        Dim dSteelDepth As Double
        Dim dSteelWidth As Double
        Dim dFlangeThick As Double
        Dim dWebThick As Double
        Dim dPipeDia As Double
        Dim MIC_C90_AUX As String
        Dim dPlateDepth As Double
        Dim dPlateWidth As Double

        Dim dPlateThick As Double
        Dim sItemNo As String
        Dim lFinish As Long
        Dim sGroup As String
        Dim dPortOffset As Double
        Dim dW As Double
        Dim dD As Double
        Dim dRodDia As Double
        Dim dRodOffset As Double
        Dim dRodL As Double
        Dim lNumberClamps As Long

        dPlateThick = arrayOfInputs(2)
        sItemNo = arrayOfInputs(3)
        lFinish = arrayOfInputs(4)
        sGroup = arrayOfInputs(5)
        dW = arrayOfInputs(6)
        dD = arrayOfInputs(7)
        dPlateDepth = arrayOfInputs(8)
        dPlateWidth = arrayOfInputs(9)
        dRodDia = arrayOfInputs(10)
        dRodOffset = arrayOfInputs(11)
        dPortOffset = arrayOfInputs(12)
        lNumberClamps = arrayOfInputs(13)
       
        If oICH Is Nothing Then  ' The OCH will be null the first time thru until the symbol is computed the first time
            dPipeDia = 0.168275 ' 6" pipe dia
        Else
            Hilti_InitializeMyHH oICH
            dSteelDepth = HH.GetSupportingProperty("Depth", 1)
            dSteelWidth = HH.GetSupportingProperty("Width", 1)
            dFlangeThick = HH.GetSupportingProperty("tf", 1)
            dWebThick = HH.GetSupportingProperty("tw", 1)
            Hilti_DestroyMyHH oICH
                        
            If dFlangeThick < 0.003 Or dFlangeThick > 0.036 Then
                PF_EventHandler "Flange thickness must be between 3 to 36 mm", Err, MODULE, METHOD, True
            End If
            
            Set oICH = Nothing
        End If
       
        'calculate the length of four rods.
        dRodL = dRodOffset + dPlateThick + dD
       
        AddPort "StructureInt", 0, 0, 0, _
                  -1, 0, 0, _
                   0, 0, -1, m_outputColl, "StructureInt", arrayOfInputs(1), "StructureInt"
                   
        AddPort "StructureExt", 0, 0, -dPortOffset, _
                  -1, 0, 0, _
                   0, 0, -1, m_outputColl, "StructureExt", arrayOfInputs(1), "StructureExt"

        If lNumberClamps = 1 Then
            AddBox dPlateDepth, dPlateWidth, dPlateThick, Loc(-dPlateDepth / 2, -dPlateWidth / 2, 0), m_outputColl, "Plate"
            AddCylinder dRodL, dRodDia / 2, Loc(0, dW / 2 - dRodDia / 2, -dD), m_outputColl, "RightRod"
            AddCylinder dRodL, dRodDia / 2, Loc(0, -dW / 2 + dRodDia / 2, -dD), m_outputColl, "LeftRod"
            AddCylinder dW, dRodDia / 2, "ROTX(90) * " + Loc(0, dW / 2, dRodDia / 2 - dD), m_outputColl, "BottomRod"
        Else
            AddBox dPlateDepth, dPlateWidth, dPlateThick, Loc(-dSteelWidth / 2 - dPlateDepth / 2 - dRodDia / 2, -dPlateWidth / 2, 0), m_outputColl, "Plate"
            AddCylinder dRodL, dRodDia / 2, Loc(-dSteelWidth / 2 - dRodDia / 2, dW / 2 - dRodDia / 2, -dD), m_outputColl, "RightRod"
            AddCylinder dRodL, dRodDia / 2, Loc(-dSteelWidth / 2 - dRodDia / 2, -dW / 2 + dRodDia / 2, -dD), m_outputColl, "LeftRod"
            AddCylinder dW, dRodDia / 2, "ROTX(90) * " + Loc(-dSteelWidth / 2 - dRodDia / 2, dW / 2, dRodDia / 2 - dD), m_outputColl, "BottomRod"
            
            AddBox dPlateDepth, dPlateWidth, dPlateThick, Loc(-dPlateDepth / 2 + dSteelWidth / 2 + dRodDia / 2, -dPlateWidth / 2, 0), m_outputColl, "Plate2"
            AddCylinder dRodL, dRodDia / 2, Loc(dSteelWidth / 2 + dRodDia / 2, dW / 2 - dRodDia / 2, -dD), m_outputColl, "RightRod2"
            AddCylinder dRodL, dRodDia / 2, Loc(dSteelWidth / 2 + dRodDia / 2, -dW / 2 + dRodDia / 2, -dD), m_outputColl, "LeftRod2"
            AddCylinder dW, dRodDia / 2, "ROTX(90) * " + Loc(dSteelWidth / 2 + dRodDia / 2, dW / 2, dRodDia / 2 - dD), m_outputColl, "BottomRod2"
        End If
        Exit Function
ErrHandler:
        Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Private Sub IJHgrSymbolBOMServices_EvaluateBOM(ByVal pSupportComp As Object, bstrBOMDesc As String)
Const METHOD = "IJHgrSymbolBOMServices_EvaluateBOM"
    On Error GoTo ErrorHandler
    
    ' <Determine BOM description here>
    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Set oPartOcc = pSupportComp ' The part occurence
    oPartOcc.GetPart oPart      ' The associated catalog part
    
    Dim strFinish As String
    Dim lFinish As Long
    Dim sItemNo As Long

    bstrBOMDesc = Hilti_BuildBom(pSupportComp)
   
    Exit Sub

ErrorHandler:
        Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub






