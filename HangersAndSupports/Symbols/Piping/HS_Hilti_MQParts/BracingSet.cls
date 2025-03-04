VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracingSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQZ_L.cls
'   ProgID:         HS_Hilti_MQParts.MQZ_L
'   Author:         JRK
'   Creation Date:  Monday, November 19, 2007

'   Description:
'       Pre-assembled 90
'
'   Change History:
'   Date            who         change description
'   16-02-2010      VSP         Addition to S3D
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts  
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "BracingSet"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.BracingSet"

Private m_SymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
Private m_SymbolDefinition As IMSSymbolEntities.DSymbolDefinition
Private m_outputColl As IJDOutputCollection
Dim oCodeListData As IJDCodeListMetaData

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

    ' Inputs
    Const iNumInputs As Integer = 9
    Dim sInputName(iNumInputs) As String
    Dim sDataType(iNumInputs) As String

    ' Outputs
    Const iNumOutputs As Integer = 3
    Dim sOutputName(iNumOutputs) As String

    ' Inputs
    sInputName(2) = "Angle1"
    sDataType(2) = "Double"

    sInputName(3) = "Thickness1"
    sDataType(3) = "Double"

    sInputName(4) = "Width1"
    sDataType(4) = "Double"

    sInputName(5) = "DimA"
    sDataType(5) = "Double"

    sInputName(6) = "DimB"
    sDataType(6) = "Double"

    sInputName(7) = "DimC"
    sDataType(7) = "Double"

    sInputName(8) = "ItemNo"
    sDataType(8) = "String"

    sInputName(9) = "Group"
    sDataType(9) = "String"

    sOutputName(1) = "Rod"
    sOutputName(2) = "Rod2"
    sOutputName(3) = "Body"

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

    Dim dAngle1 As Double
    Dim dThickness1 As Double
    Dim dWidth1 As Double
    Dim dDimA As Double
    Dim dDimB As Double
    Dim dDimC As Double
    Dim sItemNo As String
    Dim sGroup As String
    Dim dCalc1 As Double
    Dim dCalc2 As Double
    Dim dCalc3 As Double
    Dim dCalc4 As Double
    Dim dCalc6 As Double
    Dim dCalc7 As Double
    Dim dCalc8 As Double
    Dim dCalc9 As Double
    Dim dCalc10 As Double
    Dim dCalc11 As Double
    Dim dAngle2 As Double
    Dim dXDirX1 As Double
    Dim dXDirY1 As Double
    Dim dXDirZ1 As Double
    Dim dZDirX1 As Double
    Dim dZDirY1 As Double
    Dim dZDirZ1 As Double
    Dim dXDirX2 As Double
    Dim dXDirY2 As Double
    Dim dXDirZ2 As Double
    Dim dZDirX2 As Double
    Dim dZDirY2 As Double
    Dim dZDirZ2 As Double

    dAngle1 = arrayOfInputs(2)
    dThickness1 = arrayOfInputs(3)
    dWidth1 = arrayOfInputs(4)
    dDimA = arrayOfInputs(5)
    dDimB = arrayOfInputs(6)
    dDimC = arrayOfInputs(7)
    sItemNo = arrayOfInputs(8)
    sGroup = arrayOfInputs(9)

    dCalc1 = Sin(dAngle1) * dDimA
    dCalc2 = Cos(dAngle1) * dDimA
    dCalc3 = Sin(dAngle1) * (dDimA / 2)
    dCalc4 = Cos(dAngle1) * (dDimA / 2)
    dCalc6 = Sin(dAngle1) * dThickness1
    dCalc7 = Cos(dAngle1) * dThickness1
    dCalc8 = Sin(dAngle1) * (dThickness1 / 2)
    dCalc9 = Cos(dAngle1) * (dThickness1 / 2)
    dAngle2 = (180 - Deg(dAngle1)) / 2
    dCalc10 = dThickness1 / TanDeg(dAngle2)

    dXDirX1 = Cos(dAngle1) / (Cos(dAngle1) ^ 2 + Sin(dAngle1) ^ 2)
    dXDirY1 = 0
    dXDirZ1 = Sin(dAngle1) / (Cos(dAngle1) ^ 2 + Sin(dAngle1) ^ 2)
    dZDirX1 = Cos(PF_HgrGeneralUtil.Rad(90) + dAngle1) / (Cos(PF_HgrGeneralUtil.Rad(90) + dAngle1) ^ 2 + Sin(PF_HgrGeneralUtil.Rad(90) + dAngle1) ^ 2)
    dZDirY1 = 0
    dZDirZ1 = Sin(PF_HgrGeneralUtil.Rad(90) + dAngle1) / (Cos(PF_HgrGeneralUtil.Rad(90) + dAngle1) ^ 2 + Sin(PF_HgrGeneralUtil.Rad(90) + dAngle1) ^ 2)
    dXDirX2 = Cos(-dAngle1) / (Cos(-dAngle1) ^ 2 + Sin(-dAngle1) ^ 2)
    dXDirY2 = 0
    dXDirZ2 = Sin(-dAngle1) / (Cos(-dAngle1) ^ 2 + Sin(-dAngle1) ^ 2)
    dZDirX2 = Cos(PF_HgrGeneralUtil.Rad(90) - dAngle1) / (Cos(PF_HgrGeneralUtil.Rad(90) - dAngle1) ^ 2 + Sin(PF_HgrGeneralUtil.Rad(90) - dAngle1) ^ 2)
    dZDirY2 = 0
    dZDirZ2 = Sin(PF_HgrGeneralUtil.Rad(90) - dAngle1) / (Cos(PF_HgrGeneralUtil.Rad(90) + dAngle1) ^ 2 + Sin(PF_HgrGeneralUtil.Rad(90) - dAngle1) ^ 2)

    AddPort "Rod", dCalc2 / 2 - dCalc9, 0, dCalc1 / 2 + dCalc8, _
            dXDirX1, dXDirY1, dXDirZ1, _
            -dZDirX1, -dZDirY1, -dZDirZ1, _
            m_outputColl, "Rod", arrayOfInputs(1), "Rod"

    AddPort "Rod2", -dDimB, 0, dDimC / 2, _
            0, 1, 0, _
            -1, 0, 0, m_outputColl, "Rod2", arrayOfInputs(1), "Rod2"

    AddPolygon 8, "EXTRUDED", 0, "Body"
    AddLine 0, -dWidth1 / 2, 0, dCalc2, -dWidth1 / 2, dCalc1, m_outputColl, "", "Body1"
    AddLine dCalc2, -dWidth1 / 2, dCalc1, dCalc2 - dCalc7, -dWidth1 / 2, dCalc1 + dCalc6, m_outputColl, "", "Body2"
    AddLine dCalc2 - dCalc7, -dWidth1 / 2, dCalc1 + dCalc6, -dCalc10, -dWidth1 / 2, dThickness1, m_outputColl, "", "Body3"
    AddLine -dCalc10, -dWidth1 / 2, dThickness1, -dDimB + dThickness1, -dWidth1 / 2, dThickness1, m_outputColl, "", "Body4"
    AddLine -dDimB + dThickness1, -dWidth1 / 2, dThickness1, -dDimB + dThickness1, -dWidth1 / 2, dDimC, m_outputColl, "", "Body5"
    AddLine -dDimB + dThickness1, -dWidth1 / 2, dDimC, -dDimB, -dWidth1 / 2, dDimC, m_outputColl, "", "Body6"
    AddLine -dDimB, -dWidth1 / 2, dDimC, -dDimB, -dWidth1 / 2, 0, m_outputColl, "", "Body7"
    AddLine -dDimB, -dWidth1 / 2, 0, 0, -dWidth1 / 2, 0, m_outputColl, "", "Body8"
    AddExtrusion 0, dWidth1, 0, 1, m_outputColl, "Body"

    Exit Function
ErrHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Private Sub IJHgrSymbolBOMServices_EvaluateBOM(ByVal pSupportComp As Object, bstrBOMDesc As String)
    Const METHOD = "IJHgrSymbolBOMServices_EvaluateBOM"
    On Error GoTo ErrorHandler

    bstrBOMDesc = Hilti_BuildBom(pSupportComp)

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub











