VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlnAssyParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''**************************************************************************************
''  Copyright (C) 2005, Intergraph Corporation.  All rights reserved.
''
''  Project     : PlanningQI
''  File        : CPlnAssyParts.cls
''
''  Description : QueryInterpreter for Planning Assembly and their parts
''
''  Author      : Thandur Raghuveer
'' 10th Nov 2008  Apparao   TR-CP-153738  Generate a report for block assignment engine results
'' 12th Nov 2008  Apparao   TR-CP-149407  Report should not validate against Working Set
''                          TR-CP-149553  Surface Area report is not accumulated
'' 26th April 2009 Apparao  TR-CP-160536  IN Planning report Assembly and parts report spools and valves are not being reported
'' 04th May   2014 Praveen  TR-CP-160536  TR-CP-254129  As part of Oracle Support-Some Planning Reports are not showing the results.
''**************************************************************************************
Option Explicit

Private Const MODULE = "PlanningQI.CPlnAssyParts:"
Private Const DBL_UNDEFINED = 1.6E+308
Private Const IID_IJMfgParent As String = "{07AE45B9-D88F-4DF2-865D-FEBE67D82E32}"
Private Const IID_IJDProductionRouting As String = "{B447C9B4-FB74-11D1-8A49-00A0C9065DF6}"
Private Const IID_IJAssemblyChild As String = "{B447C9B4-FB74-11D1-8A49-00A0C9065DF6}"
Private Const TOLERANCE  As Double = 0.001

Implements IJQueryInterpreter

Private m_bEvaluateOnly             As Boolean
Private m_oADORecordset             As ADODB.Recordset
Private m_oFinalADORecordset        As ADODB.Recordset
Private m_oFilterResult             As Object
Private m_oParameters               As IJQueriesParameters
Private m_oQParams                  As IJQueryParameters
Private m_oCacheController          As IJCacheController
Private m_oMetaDataQuery            As IJDRepositoryQuery
Private m_oDBCommand                As ADODB.Command


Private m_strCommandType            As String
Private m_strCommandArgs            As String
Private m_strSubArgs()              As String
Private m_strOids                   As String
Private m_oPOM                      As IJDPOM

Private m_strListOfAssemblyType()   As String
Private m_strListOfAssemblyStage()  As String

Private m_lListOfAssemblyTypeNumber() As Long
Private m_lListOfAssemblyStageNumber() As Long

Private Type structWeightCG
    Weight       As String ' Double
    CGX          As String 'Double
    CGY          As String 'Double
    CGZ          As String 'Double
End Type

'Arguments

'PARENT_BLOCK
'BLK_NAME_1
'BLK_NAME_2
'BLK_NAME_3
'
'ASSY_NAME_1
'ASSY_NAME_2
'ASSY_NAME_3
'ASSY_NAME_4
'ASSY_NAME_5
'ASSY_NAME_6
'ASSY_NAME_7
'ASSY_NAME_8
'ASSY_NAME_9
'
'ASS_ASSEMBLY_TYPE
'ASS_SEQUENCE
'ASS_COMMENTS
'ASS_TYPE
'ASS_STAGE

'
'ASS_NEXT_ASSEMBLY
'ASS_WORKCENTER
'ASS_NEXT_WORKCENTER
'PART_NAME
'
'OBJ_WEIGHT
'OBJ_COGX
'OBJ_COGY
'OBJ_COGZ
'
'OBJ_AREA
'OBJ_PERMISSION
'OBJ_STATUS
'OBJ_USER

'Command Types
Private Const COMMAND_BLOCK_REPORT = "BLOCK_REPORT"
Private Const COMMAND_ASSEMBLY_REPORT = "ASSEMBLY_REPORT"
Private Const COMMAND_ASSEMBLY_STAGE_REPORT = "ASSEMBLY_STAGE_REPORT"
Private Const COMMAND_ASSEMBLY_ATTR_REPORT = "ASSEMBLY_ATTR_REPORT"
Private Const BLOCK_ASSIGNMENT = "BLOCK_ASSIGNMENT"

Private Const BLOCKASSEMBLYDIMENSIONS = "BLOCK_DIMENSIONS"
Private Const ASSEMBLY_ROUTING_REPORT = "ASSEMBLY_ROUTING_REPORT"
Private Const COMMON_PART_REPORT = "COMMON_PART_REPORT"
Private Const IID_IJCommonPartParent As String = "{8AABB5CC-1F51-4654-A433-5878A3684F5A}"
Private Const IID_IJAssemblyBase As String = "{DC16272D-8B0E-4079-A5C9-795D97E05584}"
Private m_dTotalSurfaceArea       As Double

'/////////////////////////////////////////////////////////////////////////////
'Class Constructor/Destructor
'/////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

   Set m_oADORecordset = Nothing
   Set m_oCacheController = Nothing
   Set m_oMetaDataQuery = Nothing
   Set m_oFinalADORecordset = Nothing
    
End Sub

'/////////////////////////////////////////////////////////////////////////////
' Implementation of IJQueryInterpreter
'/////////////////////////////////////////////////////////////////////////////
Public Property Set IJQueryInterpreter_CacheController _
            (RHS As SP3DReportsRuntimeInterfaces.IJCacheController)
                            
            Set m_oCacheController = RHS
            Set m_oMetaDataQuery = m_oCacheController.MetadataQuery
            
End Property

Public Property Get IJQueryInterpreter_CacheController() _
        As SP3DReportsRuntimeInterfaces.IJCacheController
                
End Property

Public Property Let IJQueryInterpreter_EvaluateOnly(ByVal RHS As Boolean)

            m_bEvaluateOnly = RHS
    
End Property

Public Property Get IJQueryInterpreter_EvaluateOnly() As Boolean
    
    IJQueryInterpreter_EvaluateOnly = m_bEvaluateOnly
        
End Property

Public Sub IJQueryInterpreter_Execute(Command As String, Argument As String)
Const METHOD = "IJQueryInterpreter_Execute"
On Error GoTo ErrorHandler
    
    Dim lIndex              As Long

    'Get the Command Type and Arguments from the input parameters
    m_strCommandType = Command
    m_strCommandArgs = Argument
    m_strSubArgs() = Split(m_strCommandArgs, ",")
    
    Set m_oPOM = GetPOM("Model")
    
    'Build the Recordset with Columns
    BuildRecordSetDef
                  
    If m_bEvaluateOnly = True Then
        Exit Sub
    End If
                                                                                                            
    If Not m_oFilterResult Is Nothing Then  ' User has picked some element through selectset
    
                        
        If m_oFilterResult.Count = 0 Then Exit Sub
        
        'Need to filtered only objects which are having parent child relations
        If m_strCommandType <> COMMAND_ASSEMBLY_STAGE_REPORT And m_strCommandType <> COMMAND_ASSEMBLY_ATTR_REPORT Then
            If Not m_oFilterResult Is Nothing Then
                FilterElementsFromPassedFilter
            End If
        End If
                
    End If
    
    If m_strCommandType <> COMMON_PART_REPORT Then
        GetAssemTypeColl ' Get the Assembly Type Collection
        
        GetAssemStageColl ' Get the Assembly Stage Collection
    End If
           
    If m_strCommandType = COMMAND_BLOCK_REPORT Then   ' Block related Reports
                    
        Call BlockReport
                        
    ElseIf m_strCommandType = COMMAND_ASSEMBLY_REPORT Or m_strCommandType = ASSEMBLY_ROUTING_REPORT Then 'Assembly related Reports
        
        Call AssemblyReport
    
    ElseIf m_strCommandType = COMMAND_ASSEMBLY_STAGE_REPORT Then 'Assembly stage report
    
        Call AssemblyStageReport
    
    ElseIf m_strCommandType = COMMAND_ASSEMBLY_ATTR_REPORT Then 'Assembly Attribute report
            
        Call AssemblyAttributeReport
    ElseIf m_strCommandType = BLOCK_ASSIGNMENT Then 'Block assignment results report
            
        Call BlockAssignmentReport
        
    ElseIf m_strCommandType = COMMON_PART_REPORT Then
        Call CommonPartReport
    ElseIf m_strCommandType = BLOCKASSEMBLYDIMENSIONS Then
       Call BlockReport
    End If
            
CleanUp:
Exit Sub
ErrorHandler:
    GoTo CleanUp
End Sub

Public Property Set IJQueryInterpreter_FilterResult(ByVal RHS As Object)

    Set m_oFilterResult = RHS
    
End Property

Public Property Get IJQueryInterpreter_FilterResult() As Variant

    Set IJQueryInterpreter_FilterResult = m_oFilterResult
    
End Property

Public Property Set IJQueryInterpreter_Parameters _
            (RHS As SP3DReportsObjectsInterfaces.IJQueryParameters)
            
    Set m_oQParams = RHS
    
End Property

Public Property Get IJQueryInterpreter_Parameters() _
            As SP3DReportsObjectsInterfaces.IJQueryParameters
            
    Set IJQueryInterpreter_Parameters = m_oQParams
    
End Property

Public Property Get IJQueryInterpreter_ParticipatingObjects() _
            As SP3DReportsRuntimeInterfaces.IJElements
            
End Property

Public Property Set IJQueryInterpreter_Query _
            (RHS As SP3DReportsObjectsInterfaces.IJQuery)
            
End Property

Public Property Get IJQueryInterpreter_Query() _
            As SP3DReportsObjectsInterfaces.IJQuery
    
End Property

Public Property Set IJQueryInterpreter_QueryFilter _
            (RHS As SP3DReportsObjectsInterfaces.IJQueryFilter)
    
End Property

Public Property Get IJQueryInterpreter_QueryFilter() _
            As SP3DReportsObjectsInterfaces.IJQueryFilter
    
End Property

Public Property Get IJQueryInterpreter_QueryResult() As Variant
    
    Set IJQueryInterpreter_QueryResult = m_oFinalADORecordset
    
End Property


'******************************************************************************
' Method        : CreateOpenRecordset
' Description   : Create Recordset and populate the property names
'
'******************************************************************************
Private Function CreateOpenRecordset() As ADODB.Recordset
Const METHOD = "CreateOpenRecordset"
On Error GoTo ErrorHandler
        
    ' Create the recordset
    Dim oRecordSet As ADODB.Recordset
    Set oRecordSet = New ADODB.Recordset
    
    m_strOids = "B0Oid"
    
    ' Specify the fields in the recordset.
    With oRecordSet.Fields
        .Append m_strOids, adVarWChar, 256, adFldIsNullable
    End With
    
    ' Open the recordset to enable addition of records.
    oRecordSet.Open , , , adLockOptimistic
    
    Set m_oADORecordset = oRecordSet
    
CleanUp:
    Set oRecordSet = Nothing
    Exit Function
ErrorHandler:
    GoTo CleanUp
End Function

Private Function GetRecordSetData(strQuery As String) As Recordset
Const METHOD = "GetRecordSetData"
On Error GoTo ErrorHandler
    
    Set m_oDBCommand = New ADODB.Command
    Set m_oDBCommand.ActiveConnection = m_oCacheController.DBConnection
    m_oDBCommand.CommandText = strQuery
    Set GetRecordSetData = m_oDBCommand.Execute
                
CleanUp:
Exit Function
ErrorHandler:
    GoTo CleanUp
End Function

Private Sub ProcessAssemblyReport(oObject As Object, bNeedToCheck As Boolean)
Const METHOD = "ProcessAssemblyReport"
On Error GoTo ErrorHandler
                    
    Dim strOId                      As String
    Dim oIMoniker                   As IMoniker
    Dim oBlockObj                   As Object
    Dim oBlockAssembly              As IJAssembly
    Dim lIndex                      As Long
    Dim oPOm                        As IJDPOM
    Dim oResObjColl                 As IJDObjectCollection
    Dim lReslIndex                  As Long
    Dim oBlockName                  As IJNamedItem
    Dim oNamedItem                  As IJNamedItem
    Dim oWeightCGGroup              As IJWeightCGgrouping
    Dim enumProperties              As WCGProperties
    Dim dWeightCg                   As structWeightCG
    Dim dCGX                        As Double
    Dim dCGY                        As Double
    Dim dCGZ                        As Double
    Dim strLevel                    As String
    Dim oSelectedObj                As Object
    Dim oAssemblyBase               As IJAssemblyBase
    Dim oBlock                      As IJBlock
    Dim oAssembly                   As IJAssembly
    Dim oParentBlockObj             As Object
    Dim oAssemblyAsChild            As IJAssemblyChild
    Dim oMfgChild                   As IJMfgChild
    
    On Error Resume Next
    Set oBlock = oObject
    Set oAssembly = oObject
    On Error GoTo ErrorHandler
    
    If oBlock Is Nothing And Not oAssembly Is Nothing Then  'The Selected one is not a Block, its an Assembly/AssemblyBase
                                                          
        Set oBlockAssembly = oObject
        
        If Not oBlockAssembly Is Nothing Then
            FillRecordsetForAssemblyReport oBlockAssembly, 1, bNeedToCheck
        End If
    
    ElseIf Not oBlock Is Nothing Then   ' The Selected object is a Block
    
        Set oBlockAssembly = oBlock
        FillRecordsetForAssemblyReport oBlockAssembly, 0, bNeedToCheck
    
    End If
        
CleanUp:
    Set oBlockObj = Nothing
    Set oIMoniker = Nothing
    
Exit Sub
ErrorHandler:
    GoTo CleanUp
End Sub

Private Function GetPOM(strDbType As String) As IJDPOM
Const METHOD = "GetPOM"
On Error GoTo ErrHandler
    
    Dim oContext As IJContext
    Dim oAccessMiddle As IJDAccessMiddle
    
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set GetPOM = oAccessMiddle.GetResourceManagerFromType(strDbType)
    
CleanUp:
    Set oContext = Nothing
    Set oAccessMiddle = Nothing
Exit Function
ErrHandler:
    GoTo CleanUp
End Function

Private Sub FillRecordsetForAssemblyReport(oAssembly As IJAssembly, lDepth As Long, bNeedToCheck As Boolean)
Const METHOD = "FillRecordsetForAssemblyReport"
On Error GoTo ErrorHandler

    If oAssembly Is Nothing Then Exit Sub
    

    Dim oTempRecordSet          As ADODB.Recordset
    Dim oRecurRecordSet         As ADODB.Recordset
    Dim oPlnInitHelper          As IJDPlnIntHelper
    Dim oAssyChildColl          As IJElements
    Dim oAssemblyBase           As IJAssemblyBase
    Dim oWeightCGGroup          As IJWeightCGgrouping
    Dim oWeightCG               As IJWeightCG
    Dim enumProperties          As WCGProperties
    Dim dWeight                 As Double
    Dim dCGX                    As Double
    Dim dCGY                    As Double
    Dim dCGZ                    As Double
    Dim oChild                  As Object
    Dim oChildName              As IJNamedItem
    Dim oAssemblyName           As IJNamedItem
    Dim oSequence               As IJAssemblySequence
    Dim oPlnAssociateCatalog    As IJPlnAssociateCatalog
    Dim oParentAssembly         As IJAssembly
    Dim oAssemblyChild          As IJAssemblyChild
    Dim strLevel                As String
    Dim lIndex                  As Long
    Dim oBlock                  As IJBlock
    Dim lSequence               As Long
    Dim dArea                   As Double
    Dim dWeightCg               As structWeightCG
    Dim bWeightCGAvailable      As Boolean

    On Error Resume Next
    Set oBlock = oAssembly
    On Error GoTo ErrorHandler
        
' 1. Update the Recordset with the given Assembly/Block

    If Not oBlock Is Nothing Then ' Passed is a Block
    
        m_oFinalADORecordset.AddNew
        
        'Get the WEight & CG Calculated values
        dWeightCg = GetWeightCG(oBlock)
    
        For lIndex = 0 To UBound(m_strSubArgs)
                
            Select Case m_strSubArgs(lIndex)
            
                Case "PARENT_BLOCK"  'Parent Block Column
                    
                    m_oFinalADORecordset.Fields("PARENT_BLOCK").Value = GetName(oBlock)
                    
                Case "OBJ_WEIGHT"
                    m_oFinalADORecordset.Fields("OBJ_WEIGHT").Value = dWeightCg.Weight
                    
                Case "OBJ_COGX"
                    m_oFinalADORecordset.Fields("OBJ_COGX").Value = dWeightCg.CGX
                    
                Case "OBJ_COGY"
                    m_oFinalADORecordset.Fields("OBJ_COGY").Value = dWeightCg.CGY
                    
                Case "OBJ_COGZ"
                    m_oFinalADORecordset.Fields("OBJ_COGZ").Value = dWeightCg.CGZ
                        
                Case "OBJ_AREA"
                    m_dTotalSurfaceArea = 0
                    m_oFinalADORecordset.Fields("OBJ_AREA").Value = GetSurfaceArea(oBlock)
                
                Case "ASS_TYPE"
                    m_oFinalADORecordset.Fields("ASS_TYPE").Value = GetAssyType(oBlock)
                                        
                Case "ASS_STAGE"
                    m_oFinalADORecordset.Fields("ASS_STAGE").Value = GetAssyStage(oBlock)
                    
                    
                                                
            End Select
            
        Next lIndex
        
        m_oFinalADORecordset.MoveNext
                                            
    Else 'Passed one is an Assembly\Assembly Block
        
        m_oFinalADORecordset.AddNew
                
        ' Fill the name of Assembly in the proper Column
        For lIndex = 1 To 9
            strLevel = "ASSY_NAME_" & lIndex
            If lIndex = lDepth Then
                If Not TypeOf oAssembly Is IJAssemblyBlock Then
                    m_oFinalADORecordset.Fields(strLevel).Value = GetName(oAssembly)
                Else
                    m_oFinalADORecordset.Fields("PARENT_BLOCK").Value = GetName(oAssembly)
                    lDepth = lDepth - 1
                End If
           End If
        Next lIndex
        
        'Get the Weight CG
        dWeightCg = GetWeightCG(oAssembly)
    
        For lIndex = 0 To UBound(m_strSubArgs)
                
            Select Case Trim(m_strSubArgs(lIndex))
               
                Case "ASS_SEQUENCE"
                    m_oFinalADORecordset.Fields("ASS_SEQUENCE").Value = GetSequence(oAssembly)
                    
                Case "ASS_WORKCENTER"
                    m_oFinalADORecordset.Fields("ASS_WORKCENTER").Value = GetWorkCenterName(oAssembly)
                                                
                Case "OBJ_WEIGHT"
                    m_oFinalADORecordset.Fields("OBJ_WEIGHT").Value = dWeightCg.Weight
                    
                Case "OBJ_COGX"
                    m_oFinalADORecordset.Fields("OBJ_COGX").Value = dWeightCg.CGX
                    
                Case "OBJ_COGY"
                    m_oFinalADORecordset.Fields("OBJ_COGY").Value = dWeightCg.CGY
                    
                Case "OBJ_COGZ"
                    m_oFinalADORecordset.Fields("OBJ_COGZ").Value = dWeightCg.CGZ
                        
                Case "OBJ_AREA"
                    m_dTotalSurfaceArea = 0
                    m_oFinalADORecordset.Fields("OBJ_AREA").Value = GetSurfaceArea(oAssembly)
                
                Case "ASS_TYPE"
                    m_oFinalADORecordset.Fields("ASS_TYPE").Value = GetAssyType(oAssembly)
                        
                Case "ASS_ASSEMBLY_TYPE"
                    m_oFinalADORecordset.Fields("ASS_ASSEMBLY_TYPE").Value = GetAssemblyType(oAssembly)
                    
                Case "ASS_STAGE"
                    m_oFinalADORecordset.Fields("ASS_STAGE").Value = GetAssyStage(oAssembly)
                
                Case "ASYM_TYPE"
                    m_oFinalADORecordset.Fields("ASYM_TYPE").Value = GetAssemblyTypeName(oAssembly)
                                                            
            End Select
            
        Next lIndex
        
        m_oFinalADORecordset.MoveNext
        
    End If
                
               
 '2. Get All parts under the given Assembly and update in the Record set
 
    Set oPlnInitHelper = New CPlnIntHelper
    Set oAssyChildColl = oPlnInitHelper.GetHierarchyAssemblyChildren(oAssembly, True)
    
    Dim oSpoolColl As IJDObjectCollection
    Dim oMfgParent As IJMfgParent
    Dim oSpool  As Object
    
    Set oMfgParent = oAssembly
    Set oSpoolColl = New JObjectCollection
    
    oMfgParent.GetChildren oSpoolColl, "IsSpoolChildOf"
    
    If oSpoolColl.Count <> 0 Then
    
        For Each oSpool In oSpoolColl
            oAssyChildColl.Add oSpool
        Next
    End If
    
    Set oSpool = Nothing
    Set oMfgParent = Nothing
    Set oSpoolColl = Nothing
    
    If oAssyChildColl.Count = 0 Then Exit Sub
            
    'First fill all the parts
    For Each oChild In oAssyChildColl
    
        ' If it is other than Assembly/Block/AssemblyBlock/UnProcessed Parts
        If Not TypeOf oChild Is IJAssemblyBase And Not TypeOf oChild Is IJPlnUnprocessedParts Then
        
            'Logic for getting the Production Workcenter and Production Routing stage code for a given part
            
            Dim oPlnProdRouting As PlanningObjects.PlnProdRouting
            Set oPlnProdRouting = New PlanningObjects.PlnProdRouting
            Set oPlnProdRouting.object = oChild
        
            Dim strProdStageWorkCenter As String, strProdStageCode As String, strAction As String, strMachine As String, strCode As String
                
            strProdStageCode = oPlnProdRouting.GetProductionRoutingCode
            strProdStageWorkCenter = oPlnProdRouting.GetWorkcenter
            
            Set oPlnProdRouting = Nothing
            
            'End of the Logic
                
            m_oFinalADORecordset.AddNew
        
            'Get the WEight & CG Calculated values
            dWeightCg = GetWeightCG(oChild)
    
            For lIndex = 0 To UBound(m_strSubArgs)
                    
                Select Case m_strSubArgs(lIndex)
                
                    Case "ASS_SEQUENCE"
                        m_oFinalADORecordset.Fields("ASS_SEQUENCE").Value = GetSequence(oChild)
                    
                    Case "PART_NAME"
                        m_oFinalADORecordset.Fields("PART_NAME").Value = GetName(oChild)
                        
                    Case "PROD_ROUTING_WORKCENTER"
                        m_oFinalADORecordset.Fields("PROD_ROUTING_WORKCENTER").Value = strProdStageWorkCenter
                        
                    Case "PROD_ROUTING_CODE"
                        m_oFinalADORecordset.Fields("PROD_ROUTING_CODE").Value = strProdStageCode
                        
                    Case "OBJ_WEIGHT"
                        m_oFinalADORecordset.Fields("OBJ_WEIGHT").Value = dWeightCg.Weight
                        
                    Case "OBJ_COGX"
                        m_oFinalADORecordset.Fields("OBJ_COGX").Value = dWeightCg.CGX
                        
                    Case "OBJ_COGY"
                        m_oFinalADORecordset.Fields("OBJ_COGY").Value = dWeightCg.CGY
                        
                    Case "OBJ_COGZ"
                        m_oFinalADORecordset.Fields("OBJ_COGZ").Value = dWeightCg.CGZ
                            
                    Case "OBJ_AREA"
                        m_dTotalSurfaceArea = 0
                        m_oFinalADORecordset.Fields("OBJ_AREA").Value = GetSurfaceArea(oChild)
                                                    
                End Select
                
            Next lIndex
            
            m_oFinalADORecordset.MoveNext
            
            If m_strCommandType = ASSEMBLY_ROUTING_REPORT Then
            
                'creating a Production Routing Element for given part
                Dim oProdRouting As IJDProductionRouting
                Set oProdRouting = GetProductionRouting(oChild)
                
                'If Production Routing Command is not applied on given part\ given part is not detailed
                If Not oProdRouting Is Nothing Then
                    Dim oRoutingActionColl As IJElements
                    
                    'getting all the Routing Actions on given part in a collection
                    oProdRouting.GetActions oRoutingActionColl
                
                    If Not oRoutingActionColl Is Nothing Then
                        If oRoutingActionColl.Count > 0 Then
                            For lIndex = 1 To oRoutingActionColl.Count
                                m_oFinalADORecordset.AddNew
        
                                Dim oProdRoutingAction As IJDRoutingAction
                                Set oProdRoutingAction = oRoutingActionColl.Item(lIndex)
                                
                                Dim oCollection As Collection
                                Set oCollection = GetActionAttrs(oProdRoutingAction)
                                
                                Dim oAttribute As IJDAttribute
                                
                                Set oAttribute = oCollection.Item(1)
                                m_oFinalADORecordset.Fields("PROD_ACTION_TYPE").Value = oAttribute.Value
                                Set oAttribute = Nothing
                                
                                Set oAttribute = oCollection.Item(2)
                                m_oFinalADORecordset.Fields("PROD_ACTION_MACHINE").Value = oAttribute.Value
                                Set oAttribute = Nothing
                                
                                Set oAttribute = oCollection.Item(3)
                                m_oFinalADORecordset.Fields("PROD_ACTION_CODE").Value = oAttribute.Value
                                Set oAttribute = Nothing
                                
                                Set oCollection = Nothing
                                Set oProdRoutingAction = Nothing
                                
                                m_oFinalADORecordset.MoveNext
                            Next lIndex
                        End If
                    End If
                End If
                Set oProdRouting = Nothing
            End If
        End If
        
    Set oChild = Nothing
    
    Next oChild
        
        
'3. Get All Assemblies and update Recordset accordingly.
    'Next Fill all Assemblies
    
    For Each oChild In oAssyChildColl
            
        If TypeOf oChild Is IJPlanningAssembly Or TypeOf oChild Is IJAssemblyBlock Then  'For Assembly\AssemblyBlock
        
            FillRecordsetForAssemblyReport oChild, lDepth + 1, True
            
        ElseIf TypeOf oChild Is IJPlnUnprocessedParts Then 'If it is Unprocessed Parts
        
            Dim oUnProcessedPartsAssy As IJAssembly
            Set oUnProcessedPartsAssy = oChild
            FillRecordsetForAssemblyReport oUnProcessedPartsAssy, lDepth + 1, True
            Set oUnProcessedPartsAssy = Nothing
            
        End If
        
        Set oChild = Nothing
        
    Next oChild
    
    
'4  Next Blocks
    For Each oChild In oAssyChildColl
        If TypeOf oChild Is IJBlock Then
            FillRecordsetForAssemblyReport oChild, 0, True
        End If
    Next oChild

CleanUp:
Exit Sub
ErrorHandler:
    GoTo CleanUp
End Sub

Private Sub BlockReport()
Const METHOD = "BlockReport"
On Error GoTo ErrorHandler

    Dim strQuery                    As String
    Dim strOId                      As String
    Dim oIMoniker                   As IMoniker
    Dim oBlockObj                   As Object
    Dim oBlock                      As IJBlock
    Dim oPlnInitHelper              As IJDPlnIntHelper

    If m_oFilterResult Is Nothing Then
                
        Set oPlnInitHelper = New CPlnIntHelper
        
        Set oBlock = oPlnInitHelper.GetTopLevelBlock(m_oPOM)
                
        If oBlock Is Nothing Then Exit Sub
        
        Set m_oFilterResult = New Collection
        
        m_oFilterResult.Add oBlock
        
        ProcessBlockReport False
        
    Else
    
        ProcessBlockReport True
                            
    End If
            
    
    
Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub AssemblyReport()
Const METHOD = "AssemblyReport"
On Error GoTo ErrorHandler
    
    Dim strQuery                    As String
    Dim strOId                      As String
    Dim oIMoniker                   As IMoniker
    Dim oObject                     As Object
    Dim lIndex                      As Long
    Dim lCount                      As Long
    Dim lArgsIndex                  As Long
    Dim oPlnInitHelper              As IJDPlnIntHelper
    Dim oBlock                      As IJBlock
    
    If m_oFilterResult Is Nothing Then
    
        Set oPlnInitHelper = New CPlnIntHelper
        
        Set oBlock = oPlnInitHelper.GetTopLevelBlock(m_oPOM)
        
        If oBlock Is Nothing Then Exit Sub
                        
        Set m_oFilterResult = New Collection
        
        m_oFilterResult.Add oBlock
                
        ProcessAssemblyReport m_oFilterResult.Item(1), False
        
    Else
    
        lCount = m_oFilterResult.Count
        
        For lIndex = 1 To lCount
            ProcessAssemblyReport m_oFilterResult.Item(lIndex), True
        Next lIndex

                    
    End If
        

Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub AssemblyStageReport()
Const METHOD = "AssemblyStageReport"
On Error GoTo ErrorHandler

    Dim strQuery        As String
    Dim lRecCount       As Long
    Dim lIndex          As Long
    Dim oAssemblyObj    As Object
    Dim strOId          As String
    Dim oIMoniker       As IMoniker

    
    If m_oFilterResult Is Nothing Then

        If Not m_oADORecordset Is Nothing Then
            Set m_oADORecordset = Nothing
        End If
    
        Set m_oADORecordset = CreateOpenRecordset    ' Cretae a Temporary Recordset for storing the query result
    
        
        If StrComp(m_oPOM.ConnectionType, "ORACLE", vbTextCompare) = 0 Then
            strQuery = "Select CAST(oid as varchar2(32)) oid from JAssemblyBase"
        Else
            strQuery = "Select oid from JAssemblyBase"
        End If
        
        Set m_oADORecordset = GetRecordSetData(strQuery) 'Get the Query Result
                
        lRecCount = m_oADORecordset.RecordCount
        
    
        For lIndex = 1 To lRecCount
        
            strOId = m_oADORecordset(0)  'Oid
            Set oIMoniker = m_oPOM.MonikerFromDbIdentifier(strOId)
                    
            Set oAssemblyObj = m_oPOM.GetObject(oIMoniker)
            
            If m_oFilterResult Is Nothing Then
                Set m_oFilterResult = New Collection
            End If
            
            m_oFilterResult.Add oAssemblyObj
            
            Set oAssemblyObj = Nothing
            
            m_oADORecordset.MoveNext
            
        Next lIndex
                                            
    End If
    
    ProcessAssemblyStageReport
            
Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub AssemblyAttributeReport()
Const METHOD = "AssemblyAttributeReport"
On Error GoTo ErrorHandler

    Dim strQuery        As String
    Dim lIndex          As Long
    Dim oAssemblyObj    As Object
    Dim strOId          As String
    Dim oIMoniker       As IMoniker
    Dim lRecCount       As Long
    
    If m_oFilterResult Is Nothing Then
    
        If Not m_oADORecordset Is Nothing Then
            Set m_oADORecordset = Nothing
        End If
    
        Set m_oADORecordset = CreateOpenRecordset    ' Cretae a Temporary Recordset for storing the query result
        
        
        If StrComp(m_oPOM.ConnectionType, "ORACLE", vbTextCompare) = 0 Then
            strQuery = "Select CAST(oid as varchar2(32)) oid from JAssemblyBase"
        Else
            strQuery = "Select oid from JAssemblyBase"
        End If

        
        Set m_oADORecordset = GetRecordSetData(strQuery) 'Get the Query Result
        
        lRecCount = m_oADORecordset.RecordCount
    
        For lIndex = 1 To lRecCount
        
            strOId = m_oADORecordset(0)  'Oid
            Set oIMoniker = m_oPOM.MonikerFromDbIdentifier(strOId)
                    
            Set oAssemblyObj = m_oPOM.GetObject(oIMoniker)
            
            If m_oFilterResult Is Nothing Then
                Set m_oFilterResult = New Collection
            End If
            
            m_oFilterResult.Add oAssemblyObj
            
            Set oAssemblyObj = Nothing
            
            m_oADORecordset.MoveNext
            
        Next lIndex
        
    End If
        
    ProcessAssemblyAttributeReport
   
Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub BuildRecordSetDef()
Const METHOD = "BuildRecordSetDef"
On Error GoTo ErrorHandler

    Dim lIndex As Long
    
    Set m_oFinalADORecordset = New Recordset
    
    For lIndex = 0 To UBound(m_strSubArgs)
        m_oFinalADORecordset.Fields.Append Trim(m_strSubArgs(lIndex)), adVarWChar, 256, adFldIsNullable
    Next lIndex
    
    m_oFinalADORecordset.Open

Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Property Get GetSequence(oObject As Object) As Long
Const METHOD = "GetSequence"
On Error GoTo ErrorHandler

    Dim oAssemblyChild          As IJAssemblyChild
    Dim oParentAssembly         As IJAssembly
    Dim oSequence               As IJAssemblySequence
    Dim oAssemblyBase           As IJAssemblySequence
    
    'Sequence Number
    On Error Resume Next
    Set oAssemblyChild = oObject
    Set oParentAssembly = oAssemblyChild.Parent
    Set oAssemblyBase = oParentAssembly
    Set oSequence = oAssemblyBase
    On Error GoTo ErrorHandler
    
    If Not oSequence Is Nothing Then
        GetSequence = oSequence.ChildIndex(oObject)
    Else
        GetSequence = vbNull
    End If
    
    Set oSequence = Nothing
    Set oParentAssembly = Nothing
    Set oAssemblyChild = Nothing
    Set oAssemblyBase = Nothing

Exit Property
ErrorHandler:
    Exit Property
End Property

Private Property Get GetWeightCG(oObject As Object) As structWeightCG
Const METHOD = "GetWeightCG"
On Error GoTo ErrorHandler

    Dim oWeightCGGroup          As IJWeightCGgrouping
    Dim oWeightCG               As IJWeightCG
    Dim enumProperties          As WCGProperties
    Dim dWeight                 As Double
    Dim dCGX                    As Double
    Dim dCGY                    As Double
    Dim dCGZ                    As Double
    Dim strUndefined            As String


    'Weight & CG
    On Error Resume Next
    Set oWeightCGGroup = oObject
    Set oWeightCG = oObject
    On Error GoTo ErrorHandler
    strUndefined = "Undefined"
    
    With GetWeightCG
    
    If Not oWeightCGGroup Is Nothing Then
    
        oWeightCGGroup.UpdateWCG
        oWeightCGGroup.GetWCG WCG_DRY, dWeight, dCGX, dCGY, dCGZ, enumProperties

            If enumProperties = WCGProperties.WCG_UNDEFINED Or enumProperties = WCG_ERROR Then
                If Abs(DBL_UNDEFINED - dWeight) <= TOLERANCE Then
                    .Weight = strUndefined
                Else
                    .Weight = dWeight
                End If
            Else
                .Weight = dWeight
            End If
            
            If Abs(DBL_UNDEFINED - dCGX) <= TOLERANCE Then
                .CGX = strUndefined
            Else
               .CGX = dCGX
            End If
            
            If Abs(DBL_UNDEFINED - dCGY) <= TOLERANCE Then
                .CGY = strUndefined
            Else
                .CGY = dCGY
            End If
            
            If Abs(DBL_UNDEFINED - dCGZ) <= TOLERANCE Then
                .CGZ = strUndefined
            Else
               .CGZ = dCGZ
            End If
        
    ElseIf Not oWeightCG Is Nothing Then
               
        oWeightCG.GetWCG WCG_DRY, dWeight, dCGX, dCGY, dCGZ, enumProperties
                    
            If enumProperties = WCGProperties.WCG_UNDEFINED Or enumProperties = WCG_ERROR Then
             If Abs(DBL_UNDEFINED - dWeight) <= TOLERANCE Then
                    .Weight = strUndefined
                Else
                    .Weight = dWeight
                End If
            Else
                .Weight = dWeight
            End If
            
            If Abs(DBL_UNDEFINED - dCGX) <= TOLERANCE Then
                .CGX = strUndefined
            Else
                .CGX = dCGX
            End If
            
            If Abs(DBL_UNDEFINED - dCGY) <= TOLERANCE Then
                .CGY = strUndefined
            Else
                .CGY = dCGY
            End If
            
            If Abs(DBL_UNDEFINED - dCGZ) <= TOLERANCE Then
               .CGZ = strUndefined
            Else
                .CGZ = dCGZ
            End If
    Else
            .Weight = strUndefined
            .CGX = strUndefined
            .CGY = strUndefined
            .CGZ = strUndefined
    End If
    
    End With
    
    Set oWeightCGGroup = Nothing
    Set oWeightCG = Nothing
        
Exit Property
ErrorHandler:
    Exit Property
End Property

Private Property Get GetName(oObject As Object) As String
Const METHOD = "GetName"
On Error GoTo ErrorHandler

    Dim oNamedItem As IJNamedItem

    On Error Resume Next
    Set oNamedItem = oObject
    Err.Clear
    On Error GoTo ErrorHandler
                
    If Not oNamedItem Is Nothing Then
        GetName = oNamedItem.Name
    Else
        GetName = vbNullString
    End If
                    
    Set oNamedItem = Nothing

Exit Property
ErrorHandler:
    Exit Property
End Property

Private Property Get GetWorkCenterName(oObject As Object) As String
Const METHOD = "GetWorkCenterName"
On Error GoTo ErrorHandler

    Dim oPlnAssociateCatalog As IJPlnAssociateCatalog

    'Get the Assembly Work center
    On Error Resume Next
    Set oPlnAssociateCatalog = oObject
    On Error GoTo ErrorHandler
    
    If Not oPlnAssociateCatalog Is Nothing Then
        GetWorkCenterName = oPlnAssociateCatalog.GetWorkcenter.WorkcenterName
    End If
    
    Set oPlnAssociateCatalog = Nothing

Exit Property
ErrorHandler:
    Exit Property
End Property

Private Property Get GetSurfaceArea(oObject As Object) As Double
Const METHOD = "GetSurfaceArea"
On Error GoTo ErrorHandler

    Dim oAssembly               As IJAssembly
    Dim oAssyUnk                As IUnknown
    Dim oModelBody              As IJDModelBody
    Dim dDummy                  As Double
    Dim dArea                   As Double
    Dim oPlnInitHelper          As IJDPlnIntHelper
    
    Dim oAssyChildColl          As IJElements
    Dim oChild                  As Object
    Dim oPlate                  As IJPlate
    Dim oProfilePart            As IJProfilePart
    
    On Error Resume Next
    Set oAssembly = oObject
    Set oModelBody = oObject
    On Error GoTo ErrorHandler
    
    Dim dTolerance              As Double
    dTolerance = 0.001
    
    If Not oAssembly Is Nothing Then
    
        Set oAssyUnk = m_oPOM.GetObjectMoniker(oAssembly)
        
        'Fill the Area column value after calculating it
        Set oPlnInitHelper = New CPlnIntHelper
        Set oAssyChildColl = oPlnInitHelper.GetStoredProcAssemblyChildren(oAssyUnk, vbNullString, False, m_oPOM, True)
        If oAssyChildColl.Count = 0 Then Exit Property
                    
        For Each oChild In oAssyChildColl
            
            On Error Resume Next
            Set oModelBody = oChild
            On Error GoTo ErrorHandler
            
            If TypeOf oChild Is IJPlate Then
                Set oPlate = oChild
                
                m_dTotalSurfaceArea = m_dTotalSurfaceArea + oPlate.SurfaceArea
                
            ElseIf TypeOf oChild Is IJProfilePart Then
                Set oProfilePart = oChild
                m_dTotalSurfaceArea = m_dTotalSurfaceArea + oProfilePart.Area
          
            Else
            
                If Not oModelBody Is Nothing Then
                    On Error Resume Next
                    oModelBody.GetDimMetrics dTolerance, dDummy, dDummy, dArea, dDummy
                    On Error GoTo ErrorHandler
                    m_dTotalSurfaceArea = m_dTotalSurfaceArea + dArea
                End If
            
            End If
            
            Set oModelBody = Nothing
            Set oChild = Nothing
            Set oProfilePart = Nothing
            Set oPlate = Nothing
            
            dArea = 0
            GetSurfaceArea = m_dTotalSurfaceArea
            
        Next oChild
    Else
        If TypeOf oObject Is IJPlate Then
            Set oPlate = oObject
            
            GetSurfaceArea = m_dTotalSurfaceArea + oPlate.SurfaceArea
        ElseIf TypeOf oObject Is IJProfilePart Then
            Set oProfilePart = oObject
            GetSurfaceArea = m_dTotalSurfaceArea + oProfilePart.Area
        ElseIf Not oModelBody Is Nothing Then 'Part part has been passed
            On Error Resume Next
            oModelBody.GetDimMetrics dTolerance, dDummy, dDummy, dArea, dDummy
            On Error GoTo ErrorHandler
            GetSurfaceArea = dArea
        Else
            GetSurfaceArea = vbNull
        End If
    End If
        
Exit Property
ErrorHandler:
    Exit Property
End Property

Private Sub ProcessBlockReport(bNeedToCheck As Boolean)
Const METHOD = "ProcessBlockReport"
On Error GoTo ErrorHandler

    Dim strOId                      As String
    Dim oIMoniker                   As IMoniker
    Dim oBlockObj                   As Object
    Dim oBlock                      As IJBlock
    Dim oChild                      As Object
    Dim lIndex                      As Long
    Dim lElemCount                  As Long
            
    If m_oFilterResult Is Nothing Then Exit Sub
    
    lElemCount = m_oFilterResult.Count
    
    For lIndex = 1 To lElemCount
        
        Set oBlockObj = m_oFilterResult.Item(lIndex)
            
        On Error Resume Next
        Set oBlock = oBlockObj
        On Error GoTo ErrorHandler
            
        If Not oBlock Is Nothing Then
            ProcessBlocksRecursively oBlock, 1, bNeedToCheck
        End If
                
        Set oBlock = Nothing
            
    Next lIndex
               
Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub FillRecordsetForBlockReport(oAssembly As IJAssembly, lAssyDepth As Long, bNeedToCheck As Boolean, Optional lBlockDepth As Long)
Const METHOD = "FillRecordsetForBlockReport"
On Error GoTo ErrorHandler

    If oAssembly Is Nothing Then Exit Sub
    
    If m_strCommandType = BLOCKASSEMBLYDIMENSIONS Then
        FillRecordSetForAssemblyAttribute oAssembly
        Exit Sub
    End If

    Dim oTempRecordSet          As ADODB.Recordset
    Dim oRecurRecordSet         As ADODB.Recordset
    Dim oPlnInitHelper          As IJDPlnIntHelper
    Dim oAssyChildColl          As IJElements
    Dim oAssemblyBase           As IJAssemblyBase
    Dim oWeightCGGroup          As IJWeightCGgrouping
    Dim oWeightCG               As IJWeightCG
    Dim enumProperties          As WCGProperties
    Dim dWeight                 As Double
    Dim dCGX                    As Double
    Dim dCGY                    As Double
    Dim dCGZ                    As Double
    Dim oChild                  As Object
    Dim oChildName              As IJNamedItem
    Dim oAssemblyName           As IJNamedItem
    Dim oSequence               As IJAssemblySequence
    Dim oPlnAssociateCatalog    As IJPlnAssociateCatalog
    Dim oParentAssembly         As IJAssembly
    Dim oAssemblyChild          As IJAssemblyChild
    Dim strLevel                As String
    Dim lIndex                  As Long
    Dim oBlock                  As IJBlock
    Dim lSequence               As Long
    Dim dArea                   As Double
    Dim dWeightCg               As structWeightCG
    Dim bWeightCGAvailable      As Boolean

    On Error Resume Next
    Set oBlock = oAssembly
    On Error GoTo ErrorHandler
        
' 1. Update the Recordset with the given Assembly/Block

    If Not oBlock Is Nothing Then ' Passed is a Block
    
        m_oFinalADORecordset.AddNew
        
        ' Fill the name of Block in the proper Column
        For lIndex = 1 To 3
            strLevel = "BLK_NAME_" & lIndex
            If lIndex = lBlockDepth Then
                m_oFinalADORecordset.Fields(strLevel).Value = GetName(oBlock)
'            Else
'                m_oFinalADORecordset.Fields(strLevel).Value = vbNullString
            End If
        Next lIndex
        
        'Get the WEight & CG Calculated values
        dWeightCg = GetWeightCG(oBlock)
    
        For lIndex = 0 To UBound(m_strSubArgs)
                
            Select Case m_strSubArgs(lIndex)
                            
                Case "OBJ_WEIGHT"
                    m_oFinalADORecordset.Fields("OBJ_WEIGHT").Value = dWeightCg.Weight
                    
                Case "OBJ_COGX"
                    m_oFinalADORecordset.Fields("OBJ_COGX").Value = dWeightCg.CGX
                    
                Case "OBJ_COGY"
                    m_oFinalADORecordset.Fields("OBJ_COGY").Value = dWeightCg.CGY
                    
                Case "OBJ_COGZ"
                    m_oFinalADORecordset.Fields("OBJ_COGZ").Value = dWeightCg.CGZ
                
                Case "OBJ_AREA"
                    m_dTotalSurfaceArea = 0
                    m_oFinalADORecordset.Fields("OBJ_AREA").Value = GetSurfaceArea(oBlock)
                
                Case "ASS_TYPE"
                    m_oFinalADORecordset.Fields("ASS_TYPE").Value = GetAssyType(oBlock)
                    
                Case "ASS_STAGE"
                    m_oFinalADORecordset.Fields("ASS_STAGE").Value = GetAssyStage(oBlock)
                    
                Case "ASS_ASSEMBLY_TYPE"
                    m_oFinalADORecordset.Fields("ASS_ASSEMBLY_TYPE").Value = GetAssyStage(oBlock)
                   
            End Select
            
        Next lIndex
        
        m_oFinalADORecordset.MoveNext
                                            
    Else 'Passed one is an Assembly\Assembly Block
        
        m_oFinalADORecordset.AddNew
                
        ' Fill the name of Assembly in the proper Column
        For lIndex = 1 To 9
            strLevel = "ASSY_NAME_" & lIndex
            If lIndex = lAssyDepth Then
                m_oFinalADORecordset.Fields(strLevel).Value = GetName(oAssembly)
'            Else
'                m_oFinalADORecordset.Fields(strLevel).Value = vbNullString
            End If
        Next lIndex

        'Get the Weight & CG
        dWeightCg = GetWeightCG(oAssembly)
    
        For lIndex = 0 To UBound(m_strSubArgs)
                
            Select Case Trim(m_strSubArgs(lIndex))
                               
                Case "ASS_WORKCENTER"
                    m_oFinalADORecordset.Fields("ASS_WORKCENTER").Value = GetWorkCenterName(oAssembly)
                                                
                Case "OBJ_WEIGHT"
                    m_oFinalADORecordset.Fields("OBJ_WEIGHT").Value = dWeightCg.Weight
                    
                Case "OBJ_COGX"
                    m_oFinalADORecordset.Fields("OBJ_COGX").Value = dWeightCg.CGX
                    
                Case "OBJ_COGY"
                    m_oFinalADORecordset.Fields("OBJ_COGY").Value = dWeightCg.CGY
                    
                Case "OBJ_COGZ"
                    m_oFinalADORecordset.Fields("OBJ_COGZ").Value = dWeightCg.CGZ
                
                Case "OBJ_AREA"
                    m_dTotalSurfaceArea = 0
                    m_oFinalADORecordset.Fields("OBJ_AREA").Value = GetSurfaceArea(oAssembly)
                
                Case "ASS_TYPE"
                    m_oFinalADORecordset.Fields("ASS_TYPE").Value = GetAssyType(oAssembly)
                    
                Case "ASS_STAGE"
                    m_oFinalADORecordset.Fields("ASS_STAGE").Value = GetAssyStage(oAssembly)
                    
                Case "ASS_ASSEMBLY_TYPE"
                    m_oFinalADORecordset.Fields("ASS_ASSEMBLY_TYPE").Value = GetAssyStage(oAssembly)
                                               
            End Select
            
        Next lIndex
        
        m_oFinalADORecordset.MoveNext
        
    End If
                
               
 '2. Get All parts under the given Assembly and update in the Record set
 
    Set oPlnInitHelper = New CPlnIntHelper
    Set oAssyChildColl = oPlnInitHelper.GetHierarchyAssemblyChildren(oAssembly, True)
    
    Dim oSpoolColl As IJDObjectCollection
    Dim oMfgParent As IJMfgParent
    Dim oSpool  As Object
    
    Set oMfgParent = oAssembly
    Set oSpoolColl = New JObjectCollection
    
    oMfgParent.GetChildren oSpoolColl, "IsSpoolChildOf"
    
    If oSpoolColl.Count <> 0 Then
    
        For Each oSpool In oSpoolColl
            oAssyChildColl.Add oSpool
        Next
    End If
    
    Set oSpool = Nothing
    Set oMfgParent = Nothing
    Set oSpoolColl = Nothing
    
    If oAssyChildColl.Count = 0 Then Exit Sub
            
    'First fill all the parts
    For Each oChild In oAssyChildColl
    
        ' If it is other than Assembly/Block/AssemblyBlock/UnProcessed Parts
        If Not TypeOf oChild Is IJAssemblyBase And Not TypeOf oChild Is IJPlnUnprocessedParts Then
                
            m_oFinalADORecordset.AddNew
        
            'Get the WEight & CG Calculated values
            dWeightCg = GetWeightCG(oChild)
    
            For lIndex = 0 To UBound(m_strSubArgs)
                    
                Select Case m_strSubArgs(lIndex)
                                                    
                    Case "PART_NAME"
                        m_oFinalADORecordset.Fields("PART_NAME").Value = GetName(oChild)
                        
                    Case "OBJ_WEIGHT"
                        m_oFinalADORecordset.Fields("OBJ_WEIGHT").Value = dWeightCg.Weight
                        
                    Case "OBJ_COGX"
                        m_oFinalADORecordset.Fields("OBJ_COGX").Value = dWeightCg.CGX
                        
                    Case "OBJ_COGY"
                        m_oFinalADORecordset.Fields("OBJ_COGY").Value = dWeightCg.CGY
                        
                    Case "OBJ_COGZ"
                        m_oFinalADORecordset.Fields("OBJ_COGZ").Value = dWeightCg.CGZ
                        
                    Case "OBJ_AREA"
                        m_dTotalSurfaceArea = 0
                        m_oFinalADORecordset.Fields("OBJ_AREA").Value = GetSurfaceArea(oChild)
                                        
                                                            
                End Select
                
            Next lIndex
            
            m_oFinalADORecordset.MoveNext
        
        End If
        
    Set oChild = Nothing
    
    Next oChild
        
        
'3. Get All Assemblies and update Recordset accordingly.
    'Next Fill all Assemblies
    
    For Each oChild In oAssyChildColl
        
        If TypeOf oChild Is IJBlock Then ' For Block, the level should start from 0
        
            'FillRecordsetForBlockReport oChild, 0,
            
        ElseIf TypeOf oChild Is IJAssemblyBase And Not TypeOf oChild Is IJBlock Then 'For Assembly\AssemblyBlock
        
            FillRecordsetForBlockReport oChild, lAssyDepth + 1, True
            
        ElseIf TypeOf oChild Is IJPlnUnprocessedParts Then 'If it is Unprocessed Parts
        
            Dim oUnProcessedPartsAssy As IJAssembly
            Set oUnProcessedPartsAssy = oChild
            FillRecordsetForBlockReport oUnProcessedPartsAssy, lAssyDepth + 1, True
            Set oUnProcessedPartsAssy = Nothing
            
        End If
        
        Set oChild = Nothing
        
    Next oChild

CleanUp:
Exit Sub
ErrorHandler:
    GoTo CleanUp
End Sub

Private Sub ProcessBlocksRecursively(oBlock As IJAssembly, lBlockDepth As Long, bNeedToCheck As Boolean)
Const METHOD = "ProcessBlocksRecursively"
On Error GoTo ErrorHandler

    If oBlock Is Nothing Then Exit Sub

    Dim oBlockChildColl             As IJElements
    Dim oBlockAllChildColl          As IJElements
    Dim oPlnInitHelper              As IJDPlnIntHelper
    Dim oChild                      As Object
    
    FillRecordsetForBlockReport oBlock, 0, False, lBlockDepth

    Set oPlnInitHelper = New CPlnIntHelper
    
    Set oBlockAllChildColl = oPlnInitHelper.GetHierarchyAssemblyChildren(oBlock, True)
    
    If oBlockAllChildColl.Count <> 0 Then
        
        For Each oChild In oBlockAllChildColl
            If TypeOf oChild Is IJBlock Then
                ProcessBlocksRecursively oChild, lBlockDepth + 1, True
            ElseIf (TypeOf oChild Is IJAssemblyBase) And (m_strCommandType = BLOCKASSEMBLYDIMENSIONS) Then
                 ProcessBlocksRecursively oChild, 1, True
            End If

        Next oChild
                                                                
    End If
        
    Set oPlnInitHelper = Nothing
    
CleanUp:
Exit Sub
ErrorHandler:
    GoTo CleanUp
End Sub

Private Sub ProcessAssemblyStageReport()
Const METHOD = "ProcessAssemblyStageReport"
On Error GoTo ErrorHandler

    Dim strOId                      As String
    Dim oIMoniker                   As IMoniker
    Dim oBlockObj                   As Object
    Dim oBlock                      As IJBlock
    Dim oChild                      As Object
    Dim lRecCount                   As Long
    Dim lIndex                      As Long
    Dim oAssyBlockObj               As Object
    Dim oAssyBlock                  As IJAssemblyBase
        
    If m_oFilterResult Is Nothing Then Exit Sub
    
    lRecCount = m_oFilterResult.Count
            
    For lIndex = 1 To lRecCount
        
        Set oAssyBlockObj = m_oFilterResult.Item(lIndex)
        
        On Error Resume Next
        Set oAssyBlock = oAssyBlockObj
        On Error GoTo ErrorHandler
        
        If Not oAssyBlock Is Nothing Then
            FillRecordSetForAssemblyStage oAssyBlockObj
        End If
        
        Set oAssyBlockObj = Nothing
        Set oAssyBlock = Nothing
                    
    Next lIndex
    
CleanUp:
Exit Sub
ErrorHandler:
    GoTo CleanUp
End Sub

Private Sub FillRecordSetForAssemblyStage(oAssemblyBlock As IJAssemblyBase)
Const METHOD = "FillRecordSetForAssemblyStage"
On Error GoTo ErrorHandler

    Dim oChildColl          As IJElements
    Dim oChild              As Object
    Dim oPlnInitHelper      As IJDPlnIntHelper
    Dim lIndex              As Long
    Dim oNextAssemblyObj    As Object

    If oAssemblyBlock Is Nothing Then Exit Sub
    
    
     m_oFinalADORecordset.AddNew
             
        For lIndex = 0 To UBound(m_strSubArgs)
        
            Select Case m_strSubArgs(lIndex)
            
                Case "PARENT_BLOCK"
                    m_oFinalADORecordset.Fields("PARENT_BLOCK").Value = GetName(oAssemblyBlock)
                
                Case "ASS_ASSEMBLY_TYPE"
                    m_oFinalADORecordset.Fields("ASS_ASSEMBLY_TYPE").Value = GetAssemblyType(oAssemblyBlock)
                
                Case "ASS_TYPE"
                    m_oFinalADORecordset.Fields("ASS_TYPE").Value = GetAssyType(oAssemblyBlock)
                    
                Case "ASS_STAGE"
                    m_oFinalADORecordset.Fields("ASS_STAGE").Value = GetAssyStage(oAssemblyBlock)
            
            End Select
                        
        Next lIndex
                                                        
    m_oFinalADORecordset.MoveNext
        
    Set oPlnInitHelper = New CPlnIntHelper
    
    Dim oIMoniker       As IMoniker
    Dim oAssyBlockObj   As Object
    
    Set oAssyBlockObj = oAssemblyBlock
    Set oIMoniker = m_oPOM.GetObjectMoniker(oAssyBlockObj)
    Set oChildColl = oPlnInitHelper.GetStoredProcAssemblyChildren(oIMoniker, "IJAssembly", False, m_oPOM, False)
    
    For Each oChild In oChildColl
    
        If TypeOf oChild Is IJAssemblyBase And Not TypeOf oChild Is IJBlock Then ' Assembly\AssemblyBlock
        
        m_oFinalADORecordset.AddNew
        
        Set oNextAssemblyObj = GetNextAssembly(oChild)
        
        For lIndex = 0 To UBound(m_strSubArgs)
        
            Select Case m_strSubArgs(lIndex)
                
                Case "ASSY_NAME_1"
                    m_oFinalADORecordset.Fields("ASSY_NAME_1").Value = GetName(oChild)
                            
                Case "ASS_WORKCENTER"
                    m_oFinalADORecordset.Fields("ASS_WORKCENTER").Value = GetWorkCenterName(oChild)
                                                
                Case "ASS_NEXT_ASSEMBLY"
                
                    If Not oNextAssemblyObj Is Nothing Then
                        m_oFinalADORecordset.Fields("ASS_NEXT_ASSEMBLY").Value = GetName(oNextAssemblyObj)
                    End If
                                
                Case "ASS_NEXT_WORKCENTER"
                
                    If Not oNextAssemblyObj Is Nothing Then
                        m_oFinalADORecordset.Fields("ASS_NEXT_WORKCENTER").Value = GetWorkCenterName(oNextAssemblyObj)
                    End If
                
                Case "ASS_SEQUENCE"
                    m_oFinalADORecordset.Fields("ASS_SEQUENCE").Value = GetSequence(oChild)
                                
                Case "ASS_ASSEMBLY_TYPE"
                    m_oFinalADORecordset.Fields("ASS_ASSEMBLY_TYPE").Value = GetAssemblyType(oChild)
                
                Case "ASS_TYPE"
                    m_oFinalADORecordset.Fields("ASS_TYPE").Value = GetAssyType(oChild)
                    
                Case "ASS_STAGE"
                    m_oFinalADORecordset.Fields("ASS_STAGE").Value = GetAssyStage(oChild)
            
            End Select
                        
        Next lIndex
                                                                
        m_oFinalADORecordset.MoveNext
        End If
        
    Next oChild
        
    Set oChild = Nothing
        
CleanUp:
Exit Sub
ErrorHandler:
    GoTo CleanUp
End Sub

Private Function GetNextAssembly(oAssembly As IJAssembly) As Object
Const METHOD = "GetNextAssembly"
On Error GoTo ErrorHandler

    Dim oParentAssemblyObj          As Object
    Dim oParentAssembly             As IJAssembly
    Dim oMfgChild                   As IJMfgChild
    
    Set GetNextAssembly = Nothing
    
    If oAssembly Is Nothing Then Exit Function
    
    Set oMfgChild = oAssembly
    
    If oMfgChild Is Nothing Then Exit Function
    
    Set oParentAssemblyObj = oMfgChild.GetParent
    
    If Not oParentAssemblyObj Is Nothing Then
        Set GetNextAssembly = oParentAssemblyObj
    End If
                           
CleanUp:
Exit Function
ErrorHandler:
    GoTo CleanUp
End Function

Private Function GetAssemblyType(oAssemblyObj As Object) As String
Const METHOD = "GetAssemblyType"
On Error GoTo ErrorHandler
    
    If TypeOf oAssemblyObj Is IJPlanningAssembly Then
        GetAssemblyType = "Assembly"
    ElseIf TypeOf oAssemblyObj Is IJAssemblyBlock Then
        GetAssemblyType = "AssemblyBlock"
    Else
        GetAssemblyType = vbNullString
    End If
        
CleanUp:
Exit Function
ErrorHandler:
    GoTo CleanUp
End Function

Private Function GetStatus(oObject As Object) As String
Const METHOD = "GetStatus"
On Error GoTo ErrorHandler

    Dim oStatusObj As IJDObject
    GetStatus = vbNullString
    Set oStatusObj = oObject
    
    Dim oIJDCodeListMetaData    As IJDCodeListMetaData
    Dim oContext                As IJContext
    Dim oAccessMiddle           As IJDAccessMiddle
    
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set oIJDCodeListMetaData = oAccessMiddle.GetResourceManagerFromType("Catalog")
    
    GetStatus = oIJDCodeListMetaData.LongStringValue("ApprovalStatus", oStatusObj.ApprovalStatus)
    
    Set oStatusObj = Nothing

CleanUp:
Exit Function
ErrorHandler:
    GoTo CleanUp
End Function

Private Function GetPermissionGroup(oObject As Object) As String
Const METHOD = "GetPermissionGroup"
On Error GoTo ErrorHandler

    Dim oStatusObj      As IJDObject
    Dim strQuery        As String
    Dim oRecordSet      As Recordset
    Dim oFields         As Fields
    Dim oField          As Field
    Dim lPermiID        As Long
    
    GetPermissionGroup = vbNullString
    
    Set oStatusObj = oObject
    
    lPermiID = oStatusObj.PermissionGroup
    
    strQuery = "Select Name from JDPermissionGroup where PermissiongroupID = " + CStr(lPermiID)
    
    Set oRecordSet = GetRecordSetData(strQuery)
        
    If Not oRecordSet.EOF Then
        Set oFields = oRecordSet.Fields
        Set oField = oFields.Item("Name")
        GetPermissionGroup = oField.Value
        Set oField = Nothing
        Set oFields = Nothing
    End If
    
    Set oRecordSet = Nothing
    Set oStatusObj = Nothing

CleanUp:
Exit Function
ErrorHandler:
    GoTo CleanUp
End Function

Private Sub ProcessAssemblyAttributeReport()
Const METHOD = "ProcessAssemblyAttributeReport"
On Error GoTo ErrorHandler

    Dim strOId                      As String
    Dim oIMoniker                   As IMoniker
    Dim oChild                      As Object
    Dim lRecCount                   As Long
    Dim lIndex                      As Long
    Dim oAssemblyObj                As Object
    Dim oAssembly                   As IJAssembly
            
    If m_oFilterResult Is Nothing Then Exit Sub
        
    lRecCount = m_oFilterResult.Count
    
    For lIndex = 1 To lRecCount
                                
        Set oAssemblyObj = m_oFilterResult.Item(lIndex)
        
        On Error Resume Next
        Set oAssembly = oAssemblyObj
        On Error GoTo ErrorHandler
        
        If Not oAssembly Is Nothing And Not TypeOf oAssemblyObj Is IJBlock Then
            FillRecordSetForAssemblyAttribute oAssembly
        End If
        
        Set oAssembly = Nothing
        Set oAssemblyObj = Nothing
                    
    Next lIndex

CleanUp:
Exit Sub
ErrorHandler:
    GoTo CleanUp
End Sub

Private Sub FillRecordSetForAssemblyAttribute(oAssembly As IJAssembly)
Const METHOD = "FillRecordSetForAssemblyAttribute"
On Error GoTo ErrorHandler

    Dim oChildColl          As IJElements
    Dim oChild              As Object
    Dim oPlnInitHelper      As IJDPlnIntHelper
    Dim lIndex              As Long
    Dim oNextAssemblyObj    As Object

    If oAssembly Is Nothing Then Exit Sub
    
    
     m_oFinalADORecordset.AddNew
             
        For lIndex = 0 To UBound(m_strSubArgs)
        
            Select Case m_strSubArgs(lIndex)
                            
                Case "ASSY_NAME_1"
                    m_oFinalADORecordset.Fields("ASSY_NAME_1").Value = GetName(oAssembly)
                    
                Case "ASS_WORKCENTER"
                    m_oFinalADORecordset.Fields("ASS_WORKCENTER").Value = GetWorkCenterName(oAssembly)
                            
                Case "OBJ_PERMISSION"
                    m_oFinalADORecordset.Fields("OBJ_PERMISSION").Value = GetPermissionGroup(oAssembly)
                                
                Case "OBJ_STATUS"
                    m_oFinalADORecordset.Fields("OBJ_STATUS").Value = GetStatus(oAssembly)
                                
                Case "OBJ_USER"
                    m_oFinalADORecordset.Fields("OBJ_USER").Value = GetUser(oAssembly)
                
                Case "ASS_SEQUENCE"
                    m_oFinalADORecordset.Fields("ASS_SEQUENCE").Value = GetSequence(oAssembly)
                                
                Case "ASS_TYPE"
                    m_oFinalADORecordset.Fields("ASS_TYPE").Value = GetAssyType(oAssembly)
                    
                Case "ASS_ASSEMBLY_TYPE"
                    m_oFinalADORecordset.Fields("ASS_ASSEMBLY_TYPE").Value = GetAssemblyType(oAssembly)
                
                Case "ASS_STAGE"
                    m_oFinalADORecordset.Fields("ASS_STAGE").Value = GetAssyStage(oAssembly)
                    
                Case "LEN_ACT"
                    m_oFinalADORecordset.Fields("LEN_ACT").Value = GetDimension(oAssembly, True, "Length Actual")
                Case "WIDTH_ACT"
                    m_oFinalADORecordset.Fields("WIDTH_ACT").Value = GetDimension(oAssembly, True, "Width Actual")
                Case "HEIGHT_ACT"
                    m_oFinalADORecordset.Fields("HEIGHT_ACT").Value = GetDimension(oAssembly, True, "Height Actual")
                Case "LEN_NOM"
                    m_oFinalADORecordset.Fields("LEN_NOM").Value = GetDimension(oAssembly, True, "Length Nominal")
                Case "WIDTH_NOM"
                    m_oFinalADORecordset.Fields("WIDTH_NOM").Value = GetDimension(oAssembly, True, "Width Nominal")
                Case "HEIGHT_NOM"
                    m_oFinalADORecordset.Fields("HEIGHT_NOM").Value = GetDimension(oAssembly, True, "Height Nominal")
                            
            End Select
                        
        Next lIndex
        
        m_oFinalADORecordset.MoveNext

CleanUp:
Exit Sub
ErrorHandler:
    GoTo CleanUp
End Sub

Private Function GetUser(oObject As Object) As String
Const METHOD = "GetUser"
On Error GoTo ErrorHandler

    Dim oJDObject As IJDObject
    
    Set oJDObject = oObject
    
    GetUser = oJDObject.UserCreated
    
    Set oJDObject = Nothing

CleanUp:
Exit Function
ErrorHandler:
    GoTo CleanUp
End Function



Private Function ObjectPassesFilter(unkMoniker As IUnknown, strDBID As String, _
                oFilter As IJSimpleFilter, bFilterBad As Boolean, bFilterMissing As Boolean) As Boolean
    On Error GoTo Errx
    Const METHOD = "ObjectPassesFilter"
    Dim bRetVal As Boolean
    
    bRetVal = False
    
    'Needed because IJSimpleFilter::FilterObjectEX will error if the filter is empty
    On Error Resume Next
    bRetVal = oFilter.FilterObjectEX(unkMoniker, strDBID, DatabaseQueryResult)
    If Err.Number <> 0 Then
        If Err.Number = MF_E_SQLQUERYEXECFAILED Then
            bRetVal = False
            bFilterBad = True
        ElseIf Err.Number = MF_E_FILTERNOTINITIALIZED Then
           'Some some other unexpected error has been returned, report it.
            bRetVal = False
            bFilterBad = True
        ElseIf Err.Number = MF_E_FILTERNOTFOUND Then
            bRetVal = False
            bFilterMissing = True
        Else
          GoTo Errx
        End If
    End If
    On Error GoTo Errx
    
    ObjectPassesFilter = bRetVal
    
    Exit Function
Errx:
Exit Function
End Function

Private Sub FilterElementsFromPassedFilter()
Const METHOD = "FilterElementsFromPassedFilter"
On Error GoTo ErrorHandler
        
    Dim oFilteredElements   As IJElements
    Dim lCount              As Long
    Dim oObject             As Object
    Dim lForIndex           As Long
        
    Set oFilteredElements = New Collection
            
    'Iterate through each element in the hierarchy
    For lCount = 1 To m_oFilterResult.Count
            
        Dim oAssyChild          As IJAssemblyChild
        Dim oImmediateParent    As IJAssembly
        Dim bRootBlock          As Boolean
        Dim bFound              As Boolean
        
        bFound = False
        
        Set oObject = m_oFilterResult.Item(lCount)
        
        If oObject Is Nothing Then GoTo LoopCleanUp
        If Not TypeOf oObject Is IJAssemblyChild Then GoTo LoopCleanUp
        
        Set oAssyChild = oObject
        
        If oAssyChild Is Nothing Then GoTo LoopCleanUp
        
        Set oImmediateParent = oAssyChild.Parent
        
        If oImmediateParent Is Nothing Then GoTo LoopCleanUp
        If Not TypeOf oImmediateParent Is IJAssembly Then GoTo LoopCleanUp
        
                
        For lForIndex = 1 To m_oFilterResult.Count 'Check whether the parent already exists r not
            
            If oImmediateParent Is m_oFilterResult.Item(lForIndex) Then
                oFilteredElements.Add m_oFilterResult.Item(lCount)
                bFound = True
                Exit For
            End If
                
        Next lForIndex
                
        If bFound = False Then
        
            bRootBlock = IsRootBlock(oImmediateParent)
            
            Dim lIndex                  As Long
            lIndex = 1
            
            While (Not bRootBlock)
                      
                Dim oTempAssyChild          As IJAssemblyChild
                Dim oTempParent             As IJAssembly
                Dim oIntermediateParent     As IJAssembly
                                
                If lIndex = 1 Then
                    If Not TypeOf oImmediateParent Is IJAssemblyChild Then GoTo LoopCleanUp
                    Set oTempAssyChild = oImmediateParent
                Else
                    If Not TypeOf oIntermediateParent Is IJAssemblyChild Then GoTo LoopCleanUp
                    Set oTempAssyChild = oIntermediateParent
                End If
                            
                If oTempAssyChild Is Nothing Then GoTo LoopCleanUp
                If Not TypeOf oTempAssyChild Is IJAssemblyChild Then GoTo LoopCleanUp
            
                Set oTempParent = oTempAssyChild.Parent
                
                If oTempParent Is Nothing Then GoTo LoopCleanUp
                If Not TypeOf oTempParent Is IJAssembly Then GoTo LoopCleanUp
                
                
                For lForIndex = 1 To m_oFilterResult.Count  'Check whether the parent already exists r not
            
                    If oTempParent Is m_oFilterResult.Item(lForIndex) Then
                        oFilteredElements.Add m_oFilterResult.Item(lCount)
                        GoTo LoopCleanUp
                    End If
                
                Next lForIndex
                                                
                bRootBlock = IsRootBlock(oTempParent)
                
                Set oIntermediateParent = oTempParent
                
                If oIntermediateParent Is Nothing Then GoTo LoopCleanUp
                If Not TypeOf oIntermediateParent Is IJAssembly Then GoTo LoopCleanUp
                
                Set oTempAssyChild = Nothing
                Set oTempParent = Nothing
                
                lIndex = lIndex + 1
                
            Wend
        
        End If
        
LoopCleanUp:
        Set oObject = Nothing
        Set oAssyChild = Nothing
        Set oImmediateParent = Nothing
    Next lCount
            
    If oFilteredElements.Count <> 0 Then  'Remove the filtered Items from the selectset
        m_oFilterResult.RemoveElements oFilteredElements
    End If
    
    Set oIntermediateParent = Nothing
    Set oFilteredElements = Nothing
            
Exit Sub
ErrorHandler:
    'ReportUnanticipatedError MODULE, METHOD
End Sub

Private Function IsRootBlock(ByVal Elem As Object) As Long
    IsRootBlock = 0
    If TypeOf Elem Is IJPlanningAssembly Then
        IsRootBlock = 0
    ElseIf TypeOf Elem Is IJAssemblyBlock Then
        IsRootBlock = 0
    ElseIf TypeOf Elem Is IJBlock Then
        If TypeOf AsIJAssemblyChild(Elem).Parent Is IJConfigProjectRoot Then 'Parent is Ship
            IsRootBlock = 1
        Else
            IsRootBlock = 0
        End If
    End If
End Function
Private Function AsIJAssemblyChild(ByVal Entity As IUnknown) As IJAssemblyChild
    Set AsIJAssemblyChild = Entity
End Function

Private Function GetAssyType(oAssemblyObj As Object) As String
Const METHOD = "GetAssyType"
On Error GoTo ErrorHandler

    Dim oAssemblyBase           As IJAssemblyBase
    Dim lIndex                  As Long
    Dim lAssyType               As Long
    
    GetAssyType = vbNullString
    
    If Not TypeOf oAssemblyObj Is IJAssemblyBase Then Exit Function
        
    Set oAssemblyBase = oAssemblyObj
    
    lAssyType = oAssemblyBase.Type
        
    For lIndex = 1 To UBound(m_lListOfAssemblyTypeNumber())
        If lAssyType = m_lListOfAssemblyTypeNumber(lIndex) Then
            GetAssyType = m_strListOfAssemblyType(lIndex)
            Exit For
        End If
    Next lIndex
    
    Set oAssemblyBase = Nothing
        
Exit Function
ErrorHandler:
    Exit Function
End Function

'Gets the Assembly types supported by the objects selected
Private Function GetAssemTypeColl() As Collection
Const METHOD = "GetAssemTypeColl"
On Error GoTo ErrorHandler

    Dim objSRDQuery                 As IJSRDQuery
    Dim objWorkcenterQuery          As IJSRDWorkcenterQuery
    Dim strShortDescription1()      As String
    Dim sTableName                  As String
    Dim oAssemTypeColl              As New Collection
    Dim i                           As Long

    ' Create the non-persistent SRDServices.SRDQuery to get the Query Object for Workcenter
    Set objSRDQuery = New SRDQuery
    
    ' Fetch the SRDWorkcenterQuery (persisted in the catalog db) to perform SQL queries on the catalog database
    Set objWorkcenterQuery = objSRDQuery.GetWorkcenterQuery()
    
    If Not objWorkcenterQuery Is Nothing Then
        ' add Assemblytype to Assembly Type
        sTableName = "AssemblyType"
        objWorkcenterQuery.GetCodeListArray m_oPOM, sTableName, m_strListOfAssemblyType, strShortDescription1, m_lListOfAssemblyTypeNumber
    End If
    
    For i = LBound(m_strListOfAssemblyType) To UBound(m_strListOfAssemblyType)
        oAssemTypeColl.Add m_strListOfAssemblyType(i)
    Next i
    
    Set GetAssemTypeColl = oAssemTypeColl
    
    Set objSRDQuery = Nothing
    Set objWorkcenterQuery = Nothing
    Set oAssemTypeColl = Nothing
    
Exit Function
ErrorHandler:
    'ReportUnanticipatedError MODULE, METHOD
End Function

'Gets the collection of assembly stages from catalog
Private Sub GetAssemStageColl()
Const METHOD = "GetAssemStageColl"
On Error GoTo ErrorHandler

    Dim objSRDQuery                 As IJSRDQuery
    Dim objWorkcenterQuery          As IJSRDWorkcenterQuery
    Dim strShortDescription1()      As String
    Dim sTableName                  As String
    Dim i                           As Long

    ' Create the non-persistent SRDServices.SRDQuery to get the Query Object for Workcenter
    Set objSRDQuery = New SRDQuery
    
    ' Fetch the SRDWorkcenterQuery (persisted in the catalog db) to perform SQL queries on the catalog database
    Set objWorkcenterQuery = objSRDQuery.GetWorkcenterQuery()
    
    If Not objWorkcenterQuery Is Nothing Then
        ' add Assemblystage to Assembly Stage
        sTableName = "AssemblyStage"
        objWorkcenterQuery.GetCodeListArray m_oPOM, sTableName, m_strListOfAssemblyStage, strShortDescription1, m_lListOfAssemblyStageNumber
    End If
        
    Set objSRDQuery = Nothing
    Set objWorkcenterQuery = Nothing
    
Exit Sub
ErrorHandler:
Exit Sub
End Sub

Private Function GetAssyStage(oAssemblyObj As Object) As String
Const METHOD = "GetAssyStage"
On Error GoTo ErrorHandler

    Dim oAssemblyBase           As IJAssemblyBase
    Dim lIndex                  As Long
    Dim lAssyStage              As Long
    
    GetAssyStage = vbNullString
    
    If Not TypeOf oAssemblyObj Is IJAssemblyBase Then Exit Function
        
    Set oAssemblyBase = oAssemblyObj
    
    lAssyStage = oAssemblyBase.Stage
        
    For lIndex = 1 To UBound(m_lListOfAssemblyStageNumber())
        If lAssyStage = m_lListOfAssemblyStageNumber(lIndex) Then
            GetAssyStage = m_strListOfAssemblyStage(lIndex)
            Exit For
        End If
    Next lIndex
    
    Set oAssemblyBase = Nothing

Exit Function
ErrorHandler:
Exit Function
End Function
Private Sub BlockAssignmentReport()
Const METHOD = "BlockAssignmentReport"
On Error GoTo ErrorHandler

    Dim strQuery        As String
    Dim lRecCount       As Long
    Dim lIndex          As Long
    Dim oFolder         As Object
    Dim strOId          As String
    Dim oIMoniker       As IMoniker
    Dim oPlnFailedparts As IJPlnFailedParts
    Dim oPlnUnProcessparts As IJPlnUnprocessedParts
    Dim oAssChild       As IJAssemblyChild
    Dim oAssy           As IJAssembly
    Dim oNamedItem      As IJNamedItem
    Dim lObjCount       As Long
    
    If Not m_oFilterResult Is Nothing Then
        
        Call GetUnassignedPartsFromSelectSet
    Else
        Set m_oFilterResult = New Collection

    
        If Not m_oADORecordset Is Nothing Then
            Set m_oADORecordset = Nothing
        End If
    
        Set m_oADORecordset = CreateOpenRecordset    ' Cretae a Temporary Recordset for storing the query result
       
        Dim strconnType As String
        strconnType = m_oPOM.ConnectionType
          
        If StrComp(strconnType, "ORACLE", vbTextCompare) <> 0 Then
            strQuery = "select oidDestination from dbo.XAssemblyHierarchy where left(oidDestination, 8) in ('00113E14','00113E15')"
        Else
            strQuery = "SELECT CAST(oidDestination as varchar2(32)) oidDestination FROM XAssemblyHierarchy WHERE SUBSTR(oidDestination, 0, 8) IN ('00113E14','00113E15' )"
        End If
               
        
        Set m_oADORecordset = GetRecordSetData(strQuery) 'Get the Query Result
                
        lRecCount = m_oADORecordset.RecordCount
    
        For lIndex = 1 To lRecCount
        
            strOId = m_oADORecordset(0)  'Oid
            Set oIMoniker = m_oPOM.MonikerFromDbIdentifier(strOId)
                    
            Set oFolder = m_oPOM.GetObject(oIMoniker)
            
            m_oFilterResult.Add oFolder
            
            Set oFolder = Nothing
            
            m_oADORecordset.MoveNext
            
        Next lIndex
    
    End If
    
    For lObjCount = 1 To m_oFilterResult.Count
        On Error Resume Next
        Set oPlnFailedparts = m_oFilterResult.Item(lObjCount)
        Set oPlnUnProcessparts = m_oFilterResult.Item(lObjCount)
        On Error GoTo ErrorHandler
        
        Dim strBlockName() As String
        m_oFinalADORecordset.AddNew
        
        If Not oPlnFailedparts Is Nothing Then
            m_oFinalADORecordset.Fields("UNASSIGNED").Value = "Failed Parts"
            m_oFinalADORecordset.Fields("TYPE").Value = "Failed"
        ElseIf Not oPlnUnProcessparts Is Nothing Then
            m_oFinalADORecordset.Fields("UNASSIGNED").Value = "UnAssigned Parts"
            m_oFinalADORecordset.Fields("TYPE").Value = "UnAssigned"
            m_oFinalADORecordset.Fields("STATUS").Value = GetBlockAssignmentStatus(oPlnUnProcessparts)
        End If
        
        Set oAssChild = m_oFilterResult.Item(lObjCount)
        Set oAssy = m_oFilterResult.Item(lObjCount)
        
        'Report the no. of children under Unassigned or Failed parts folder
        m_oFinalADORecordset.Fields("CHILDREN").Value = oAssy.CountChildren
        
        Set oAssy = oAssChild.Parent
        Set oNamedItem = oAssy
        
        'Show the immediate parent of the unassigned or Failed parts folder
        If IsRootBlock(oAssy) Then
            m_oFinalADORecordset.Fields("BLOCKLEVEL1").Value = oNamedItem.Name
        End If
        
        m_oFinalADORecordset.Fields("PARENTBLOCK").Value = oNamedItem.Name
        
        Dim lDepth As Long
        lDepth = 0
        
        'get all the blocks in the hierarchy till project root
        Set oAssChild = oAssy
        Set oAssy = oAssChild.Parent
        
        While Not TypeOf oAssy Is IJConfigProjectRoot
           ReDim Preserve strBlockName(0 To lDepth) As String
           Set oNamedItem = oAssy
           
           strBlockName(lDepth) = oNamedItem.Name
           lDepth = lDepth + 1
           Set oAssChild = oAssy
           Set oAssy = oAssChild.Parent
            
        Wend
Breakloop:
        If lDepth > 0 Then
            lDepth = 1
            For lIndex = UBound(strBlockName) To LBound(strBlockName) Step -1
                m_oFinalADORecordset.Fields("BLOCKLEVEL" & CStr(lDepth)).Value = strBlockName(lIndex)
                lDepth = lDepth + 1
                If lDepth > 5 Then Exit For
            Next lIndex
        End If
        
        Set oPlnFailedparts = Nothing
        Set oPlnUnProcessparts = Nothing
        
        Set oAssChild = Nothing
        Set oAssy = Nothing
        Set oNamedItem = Nothing
   Next lObjCount
        
    
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Sub GetUnassignedPartsFromSelectSet()
Const METHOD = "GetUnassignedPartsFromSelectSet"
On Error GoTo ErrorHandler

    Dim lIndex              As Long
    Dim oHierarxhyHelper    As IJHierarchyHlper
    Dim oCollection         As IJDObjectCollection
    Dim oChildColl          As IJDObjectCollection
    Dim i                   As Long
    Dim j                   As Long
    Dim oMonikers           As IJElements
    Dim oObj                As Object
    Dim vObj                As Variant
    Dim oPlnIntHelper       As GSCADPlnIntHelper.IJDPlnIntHelper
    
    Set oPlnIntHelper = New CPlnIntHelper
    
    Set oChildColl = New JObjectCollection
    
    Set oHierarxhyHelper = New HierarchyHlper
    
    Set oCollection = New JObjectCollection
    
    For i = 1 To m_oFilterResult.Count
        If TypeOf m_oFilterResult.Item(i) Is IJBlock Then
            oCollection.Add m_oFilterResult.Item(i)
        End If
    Next i
    
    m_oFilterResult.Clear
    
    'First get all the blocks under the selected block
    For Each oObj In oCollection
        Set oChildColl = oHierarxhyHelper.CollectLeafBlocksByMoniker(oObj)
        If oChildColl.Count = 0 Then
            'get the UnassignedParts folder from the block
            Set oMonikers = oPlnIntHelper.GetStoredProcAssemblyChildren(oObj, "IJPlnUnprocessedParts", True, Nothing, False)
            If Not oMonikers Is Nothing Then
                If oMonikers.Count >= 1 Then
                    m_oFilterResult.Add oMonikers.Item(1)
                End If
            End If
            'get the failed Parts folder from the block
            Set oMonikers = oPlnIntHelper.GetStoredProcAssemblyChildren(oObj, "IJPlnFailedParts", True, Nothing, False)
            If Not oMonikers Is Nothing Then
                If oMonikers.Count >= 1 Then
                    m_oFilterResult.Add oMonikers.Item(1)
                End If
            End If
        Else
            oChildColl.Add m_oPOM.GetObjectMoniker(oObj)
        
            For Each vObj In oChildColl
                    Set oMonikers = oPlnIntHelper.GetStoredProcAssemblyChildren(m_oPOM.GetObject(vObj), "IJPlnUnprocessedParts", True, Nothing, False)
                    If Not oMonikers Is Nothing Then
                        If oMonikers.Count >= 1 Then
                            m_oFilterResult.Add oMonikers.Item(1)
                        End If
                    End If
                    Set oMonikers = oPlnIntHelper.GetStoredProcAssemblyChildren(m_oPOM.GetObject(vObj), "IJPlnFailedParts", True, Nothing, False)
                    If Not oMonikers Is Nothing Then
                        If oMonikers.Count >= 1 Then
                            m_oFilterResult.Add oMonikers.Item(1)
                        End If
                    End If
            Next
        End If
    Next

    Set oCollection = Nothing
    Set oChildColl = Nothing
    Set oObj = Nothing
    Set oMonikers = Nothing


Exit Sub
ErrorHandler:
    Err.Clear
End Sub


Private Sub CommonPartReport()
Const METHOD = "CommonPartReport"
On Error GoTo ErrorHandler

    Dim oAssemblyObjects            As IJElements
    Dim lObjCount                   As Long
    Dim lObjIndex                   As Long
    Dim strQuery                    As String
    Dim oBlock                      As IJBlock
    Dim oAssemblyBase               As IJAssemblyBase
    Dim oPlnIntHelper               As IJDPlnIntHelper
    Dim oChildAssemblies            As IJElements
    
    Set oAssemblyObjects = New JObjectCollection
    Set oPlnIntHelper = New CPlnIntHelper
    
    'get the assembly objects to be processed.
    If Not m_oFilterResult Is Nothing Then
    
        lObjCount = m_oFilterResult.Count
        
        For lObjIndex = 1 To lObjCount
        
            If TypeOf m_oFilterResult.Item(lObjIndex) Is IJAssemblyBase Then
                
                Set oAssemblyBase = m_oFilterResult.Item(lObjIndex)
                oAssemblyObjects.Add oAssemblyBase
                
                Set oChildAssemblies = oPlnIntHelper.GetAssemblyChildren(oAssemblyBase, "IJAssemblyBase", True)
                
                If Not oChildAssemblies Is Nothing Then
                    If oChildAssemblies.Count > 0 Then
                        oAssemblyObjects.AddElements oChildAssemblies
                    End If
                End If
            End If
        
        Next
    Else
        
        Set oBlock = oPlnIntHelper.GetTopLevelBlock(m_oPOM)
        
        If oBlock Is Nothing Then Exit Sub
                        
        oAssemblyObjects.Add oBlock
        
        Set oChildAssemblies = oPlnIntHelper.GetAssemblyChildren(oBlock, "IJAssemblyBase", True)
        
        If Not oChildAssemblies Is Nothing Then
            If oChildAssemblies.Count > 0 Then
                oAssemblyObjects.AddElements oChildAssemblies
            End If
        End If
            
    End If

    ProcessCommonPartReport oAssemblyObjects
    
Exit Sub
ErrorHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Sub

'**************************************************
'Method : ProcessCommonPartReport
'         This method reports the details of CommonParts
'         as shown in the Common Part Explorer.
'
'**************************************************
Private Sub ProcessCommonPartReport(oInputElems As IJElements)
Const METHOD As String = "ProcessCommonPartReport"
On Error GoTo ErrorHandler

    Dim lAssemblyCount      As Long
    Dim lAssemblyIndex      As Long
    Dim lFolderIndex        As Long
    Dim lGroupIndex         As Long
    Dim lSubFolderIndex     As Long
    Dim lPartIndex          As Long
    
    Dim oAssembly           As IJAssemblyBase
    Dim oNamedItem          As IJNamedItem
    Dim oEachPart           As Object
    
    lAssemblyCount = oInputElems.Count
        
    For lAssemblyIndex = 1 To lAssemblyCount
            
        Set oAssembly = oInputElems.Item(lAssemblyIndex)
        
        Dim oCmnPartMgr     As IJCommonPartManager
        Dim oFolders        As IJElements
        
        Set oCmnPartMgr = GetCPM(oAssembly)
                
        'If a Particular Assembly doesnot have CPMgr,
        'continue with the next assemblies.
        If oCmnPartMgr Is Nothing Then
            GoTo ContinueNextAssembly
        End If
        
        m_oFinalADORecordset.AddNew
        
        If Not oAssembly Is Nothing Then
            Set oNamedItem = oAssembly
            m_oFinalADORecordset.Fields("ASSY_NAME_1").Value = oNamedItem.Name
        End If
                
        ' for each manager, get the Folders under it
        oCmnPartMgr.GetCommonPartFolders oFolders, 1
                
        For lFolderIndex = 1 To oFolders.Count
            
            m_oFinalADORecordset.AddNew
            
            Dim oCmnPartFolder As IJCommonPartFolder
            Set oCmnPartFolder = oFolders.Item(lFolderIndex)
            
            Set oNamedItem = oCmnPartFolder
            m_oFinalADORecordset.Fields("CMNPART_FOLDER").Value = oNamedItem.Name
                                   
            'for each folder , get the groups under it
            Dim oGroups As IJElements
            
            oCmnPartFolder.GetCommonPartGroups oGroups
                                    
            For lGroupIndex = 1 To oGroups.Count

                Dim oCmnPartGroup       As IJCommonPartGroup
                Dim oCommonParts        As IJElements
                
                Set oCmnPartGroup = oGroups.Item(lGroupIndex)
                Set oNamedItem = oCmnPartGroup

                m_oFinalADORecordset.AddNew
                m_oFinalADORecordset.Fields("CMNPART_GROUP").Value = oNamedItem.Name
                                
                oCmnPartGroup.GetCommonParts oCommonParts
                                                
                m_oFinalADORecordset.Fields("PART_COUNT").Value = oCommonParts.Count
                
                ' Fill Process Control and Process Purpose
                FillCommonPartGroupDetails oCmnPartGroup
                
                For lPartIndex = 1 To oCommonParts.Count
                    Set oEachPart = oCommonParts.Item(lPartIndex)
                    Set oNamedItem = oEachPart
                    
                    m_oFinalADORecordset.AddNew
                    m_oFinalADORecordset.Fields("COMMON_PART").Value = oNamedItem.Name
                    
                    m_oFinalADORecordset.Fields("PART_STATUS").Value = GetStatus(oEachPart)
                    
                    If TypeOf oEachPart Is ISPSMemberPartPrismatic Then
                        FillMemberPartDetails oEachPart
                    Else
                        'For Profiles
                        If TypeOf oEachPart Is IJProfilePart Then
                            FillProfileDetails oEachPart
                        Else
                            FillPlateDetails oEachPart
                        End If
                        
                        'Material & Grade
                        FillMaterialInfo oEachPart
                    End If
                    
                    'Weight & CG
                    FillWeightCGInfo oEachPart
                    
                    Set oEachPart = Nothing
                Next

            Next lGroupIndex
            
            'Get the next level of folders , if any
            Dim osubFolder As IJCommonPartFolder

            For lSubFolderIndex = 1 To 3 ' 3 types are available
            
                Set osubFolder = oCmnPartFolder.GetSubFolder(lSubFolderIndex)
                                
                If Not osubFolder Is Nothing Then
                    Set oNamedItem = osubFolder
                    
                    'SubFolders like Uncommon/Relocated/Removed Parts are also
                    'listed under Cmnpart_Group category.
                    m_oFinalADORecordset.AddNew
                    m_oFinalADORecordset.Fields("CMNPART_GROUP").Value = oNamedItem.Name

                    Dim oSubFolderParts As IJElements
                                        
                    Set oSubFolderParts = GetSubFolderParts(osubFolder)
                                        
                    m_oFinalADORecordset.Fields("PART_COUNT").Value = oSubFolderParts.Count
                    
                    For lPartIndex = 1 To oSubFolderParts.Count
                        Set oEachPart = oSubFolderParts.Item(lPartIndex)
                        Set oNamedItem = oEachPart
                        
                        m_oFinalADORecordset.AddNew
                        m_oFinalADORecordset.Fields("COMMON_PART").Value = oNamedItem.Name
                        
                        m_oFinalADORecordset.Fields("PART_STATUS").Value = GetStatus(oEachPart)
                        
                        If TypeOf oEachPart Is ISPSMemberPartPrismatic Then
                            FillMemberPartDetails oEachPart
                        Else
                            'For Profiles
                            If TypeOf oEachPart Is IJProfilePart Then
                                FillProfileDetails oEachPart
                            Else
                                FillPlateDetails oEachPart
                            End If
                            
                            'Material & Grade
                            FillMaterialInfo oEachPart
                        End If
                        
                        'Weight & CG
                        FillWeightCGInfo oEachPart
                    
                        Set oEachPart = Nothing
                    
                    Next
                End If
                
                Set osubFolder = Nothing

            Next lSubFolderIndex
            
        Next lFolderIndex
            
        Set oAssembly = Nothing
        Set oNamedItem = Nothing
    
ContinueNextAssembly:
   Next lAssemblyIndex
   
Exit Sub
ErrorHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Sub

Private Function GetCPM(oAsm As IJAssemblyBase) As IJCommonPartManager
Const METHOD As String = "GetCPM"
On Error GoTo ErrorHandler

    ' Navigate the common part Mgr relation and get the required object
    Dim oAssocRel As IJDAssocRelation
    Dim oTargetObjCol As IJDTargetObjectCol
    
    Set oAssocRel = oAsm
    Set oTargetObjCol = oAssocRel.CollectionRelations(IID_IJMfgParent, "CPMgrChild")
    
    If oTargetObjCol.Count = 1 Then
        Set GetCPM = oTargetObjCol.Item(1)
    End If
    
Exit Function
ErrorHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Function

Private Function GetSubFolderParts(osubFolder As IJCommonPartFolder) As IJElements
Const METHOD = "GetSubFolderParts"
On Error GoTo ErrorHandler

    Set GetSubFolderParts = New JObjectCollection

    ' Navigate the common part hierarchy relation and get the required object
    Dim oAssocRel As IJDAssocRelation
    Dim oTargetObjCol As IJDTargetObjectCol
    
    Set oAssocRel = osubFolder
    Set oTargetObjCol = oAssocRel.CollectionRelations(IID_IJCommonPartParent, "CommonMember")
    
    If oTargetObjCol.Count > 0 Then
        Dim i As Long
        
        For i = 1 To oTargetObjCol.Count
            GetSubFolderParts.Add oTargetObjCol.Item(i)
        Next
    End If
        
Exit Function
ErrorHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Function

Private Sub FillPlateDetails(oObject As Object)
Const METHOD = "FillPlateDetails"
On Error GoTo ErrHandler
    
    Dim oPlate   As IJPlate
    Set oPlate = oObject
    
    m_oFinalADORecordset.Fields("MATERIAL_THICKNESS").Value = oPlate.thickness * 1000  '' in millimeters ''
    m_oFinalADORecordset.Fields("PART_TYPE").Value = GetPlateType(oPlate.plateType)
    m_oFinalADORecordset.Fields("PART_DESC").Value = oPlate.Description

    Set oPlate = Nothing
    
Exit Sub
ErrHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Sub

Private Sub FillProfileDetails(oObj As Object)
Const METHOD As String = "FillProfileDetails"
On Error GoTo ErrorHandler
    
    Dim oProfilePart  As IJProfilePart
    Dim oStiffener    As IJStiffener
    
    Set oProfilePart = oObj
    m_oFinalADORecordset.Fields("OBJ_LENGTH").Value = oProfilePart.ProfileLength
    
    Set oStiffener = oObj
    m_oFinalADORecordset.Fields("PART_TYPE").Value = GetProfileType(oStiffener.pType)
    
    m_oFinalADORecordset.Fields("PART_DESC").Value = oStiffener.Description
    
Exit Sub
ErrorHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Sub

Private Sub FillMemberPartDetails(oObject As Object)
Const METHOD = "FillMemberPartDetails"
On Error GoTo ErrHandler
   
    Dim oMemberPart As ISPSMemberPartPrismatic
    Set oMemberPart = oObject
    
    Dim oIJDCodeListMetaData    As IJDCodeListMetaData
    Dim oContext                As IJContext
    Dim oAccessMiddle           As IJDAccessMiddle
    Dim lMemberPartType         As Long
    Dim strMemberPartType       As String
    
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set oIJDCodeListMetaData = oAccessMiddle.GetResourceManagerFromType("Catalog")
    
    ' Use ShortStringValue as LongStringValue is empty
    lMemberPartType = oMemberPart.MemberType.Type
    strMemberPartType = oIJDCodeListMetaData.ShortStringValue("StructuralMemberType", lMemberPartType)
    
    m_oFinalADORecordset.Fields("PART_TYPE").Value = strMemberPartType
    m_oFinalADORecordset.Fields("MATERIAL").Value = oMemberPart.MaterialDefinition.MaterialType
    m_oFinalADORecordset.Fields("MATERIAL_GRADE").Value = oMemberPart.MaterialDefinition.MaterialGrade
    m_oFinalADORecordset.Fields("OBJ_LENGTH").Value = oMemberPart.CutLength
    
Exit Sub
ErrHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Sub

Private Sub FillCommonPartGroupDetails(oCmnPartGroup As IJCommonPartGroup)
Const METHOD = "FillCommonPartGroupDetails"
On Error GoTo ErrHandler
    
    Dim oIJDCodeListMetaData    As IJDCodeListMetaData
    Dim oContext                As IJContext
    Dim oAccessMiddle           As IJDAccessMiddle
    Dim lProcessControl         As Long
    Dim lProcessPurpose         As Long
    Dim strProcessControl       As String
    Dim strProcessPurpose       As String
    
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set oIJDCodeListMetaData = oAccessMiddle.GetResourceManagerFromType("Catalog")
    
    lProcessControl = oCmnPartGroup.ProcessControl
    strProcessControl = oIJDCodeListMetaData.LongStringValue("ProcessControl", lProcessControl)
    
    lProcessPurpose = oCmnPartGroup.ProcessPurpose
    strProcessPurpose = oIJDCodeListMetaData.LongStringValue("ProcessPurpose", lProcessPurpose)
    
    m_oFinalADORecordset.Fields("GROUP_PROCESS_CONTROL").Value = strProcessControl
    m_oFinalADORecordset.Fields("GROUP_PROCESS_PURPOSE").Value = strProcessPurpose
    
Exit Sub
ErrHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Sub

Private Sub FillMaterialInfo(oObj As Object)
Const METHOD As String = "FillMaterialInfo"
On Error GoTo ErrorHandler
   
    Dim oAttributeMetaData          As IJDAttributeMetaData
    Dim oAttrHelper                 As IJDAttributes
    Dim oAttributesCollection       As IJDAttributesCol
    Dim oAttribute                  As IJDAttribute
    Dim oInterfaceInfo              As IJDInterfaceInfo
    
    Set oAttributeMetaData = oObj
    Set oAttrHelper = oObj

    If oAttributeMetaData Is Nothing Then Exit Sub
    If oAttrHelper Is Nothing Then Exit Sub

    Set oInterfaceInfo = oAttributeMetaData.InterfaceInfo(oAttributeMetaData.iID("IJStructureMaterial"))

    If Not oInterfaceInfo Is Nothing Then
        
        On Error Resume Next
        Set oAttributesCollection = oAttrHelper.CollectionOfAttributes(oInterfaceInfo.Type)
        Err.Clear
        
        On Error GoTo ErrorHandler

        If Not oAttributesCollection Is Nothing Then
            For Each oAttribute In oAttributesCollection
                
                If oAttribute.AttributeInfo.Name = "MaterialName" Then
                    m_oFinalADORecordset.Fields("MATERIAL").Value = oAttribute.Value
                End If
                If oAttribute.AttributeInfo.Name = "MaterialGrade" Then
                    m_oFinalADORecordset.Fields("MATERIAL_GRADE").Value = oAttribute.Value
                End If
                
            Next
        End If
    End If
    
    Set oAttributeMetaData = Nothing
    Set oAttrHelper = Nothing
    Set oAttributesCollection = Nothing
    Set oAttribute = Nothing
    Set oInterfaceInfo = Nothing
        
Exit Sub
ErrorHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Sub

Private Sub FillWeightCGInfo(oObj As Object)
Const METHOD As String = "FillWeightCGInfo"
On Error GoTo ErrorHandler

    Dim dWeightCg As structWeightCG
    
    dWeightCg = GetWeightCG(oObj)

    m_oFinalADORecordset.Fields("OBJ_WEIGHT").Value = dWeightCg.Weight
    m_oFinalADORecordset.Fields("OBJ_COGX").Value = dWeightCg.CGX
    m_oFinalADORecordset.Fields("OBJ_COGY").Value = dWeightCg.CGY
    m_oFinalADORecordset.Fields("OBJ_COGZ").Value = dWeightCg.CGZ
    
Exit Sub
ErrorHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Sub

Private Function GetPlateType(lPlateType As Long) As String

    Dim oIJDCodeListMetaData    As IJDCodeListMetaData
    Dim oContext                As IJContext
    Dim oAccessMiddle           As IJDAccessMiddle
    
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set oIJDCodeListMetaData = oAccessMiddle.GetResourceManagerFromType("Catalog")
    
    GetPlateType = oIJDCodeListMetaData.LongStringValue("StructPlateType", lPlateType)

End Function

Private Function GetProfileType(lProfileType As Long) As String


    Dim oIJDCodeListMetaData    As IJDCodeListMetaData
    Dim oContext                As IJContext
    Dim oAccessMiddle           As IJDAccessMiddle

    
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set oIJDCodeListMetaData = oAccessMiddle.GetResourceManagerFromType("Catalog")
    
    GetProfileType = oIJDCodeListMetaData.LongStringValue("StructProfileType", lProfileType)

End Function

Private Function GetDimension(oAssembly As Object, bActual As Boolean, strPropName As String) As Double
Const METHOD = "GetDimension"
On Error GoTo ErrorHandler

    Dim objAssemblyOrientation              As GSCADAssembly.IJAssemblyOrientation
    Dim oRootPoint                          As IJDPosition
    Dim oXAxis                              As IJDVector
    Dim oYAxis                              As IJDVector
    Dim oZAxis                              As IJDVector
    
    Set objAssemblyOrientation = oAssembly
    
    On Error Resume Next
        If bActual Then
            objAssemblyOrientation.GetRangeBoxInAssemblyOrientation False, True, oRootPoint, oXAxis, oYAxis, oZAxis
        Else
            objAssemblyOrientation.GetRangeBoxInAssemblyOrientation True, True, oRootPoint, oXAxis, oYAxis, oZAxis
        End If
    Err.Clear
    
    On Error GoTo ErrorHandler
    
    Select Case strPropName
        Case "Length Actual"
            GetDimension = CVar(oXAxis.Length)
        Case "Length Nominal"
            GetDimension = CVar(oXAxis.Length)
        Case "Width Actual"
            GetDimension = CVar(oYAxis.Length)
        Case "Width Nominal"
            GetDimension = CVar(oYAxis.Length)
        Case "Height Actual"
            GetDimension = CVar(oZAxis.Length)
        Case "Height Nominal"
            GetDimension = CVar(oZAxis.Length)
    End Select

wrapup:
    Set objAssemblyOrientation = Nothing
    Set oRootPoint = Nothing
    Set oXAxis = Nothing
    Set oYAxis = Nothing
    Set oZAxis = Nothing
    
Exit Function
ErrorHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Function

'***************************************************************************'
' Description: This function gives the Assembly Type for a given Assembly as input

Private Function GetAssemblyTypeName(oAssemblyObj As Object) As String
Const METHOD = "GetAssemblyTypeName"
On Error GoTo ErrorHandler

    Dim oAssemblyBase As IJAssemblyBase
    Dim oMetaData As IJDCodeListMetaData
    Dim lAsymType As Long
    
    If TypeOf oAssemblyObj Is IJAssemblyBase Then
        Set oAssemblyBase = oAssemblyObj
        lAsymType = oAssemblyBase.Type
        Set oMetaData = oAssemblyBase
        GetAssemblyTypeName = oMetaData.LongStringValue("AssemblyType", lAsymType)
    End If
    
    Set oAssemblyBase = Nothing
    Set oMetaData = Nothing
        
Exit Function
ErrorHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Function

Private Function GetProductionRouting(oEntity As Object) As Object
Const METHOD = "GetProductionRouting"
On Error GoTo ErrorHandler

 ' Navigate the common part hierarchy relation and the required object
    Dim oAssocRel As IJDAssocRelation
    Dim oTargetObjCol As IJDTargetObjectCol

    Set oAssocRel = oEntity
    Set oTargetObjCol = oAssocRel.CollectionRelations(IID_IJDProductionRouting, "ProductionRouting")
    
    If oTargetObjCol.Count > 0 Then
        Set GetProductionRouting = oTargetObjCol.Item(1)
    End If

    Set oAssocRel = Nothing
    Set oTargetObjCol = Nothing

Exit Function
ErrorHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Function

Private Function GetActionAttrs(ByVal oSp As IJDRoutingAction) As Collection
Const METHOD = "GetActionAttrs"
On Error GoTo ErrorHandler
    Dim oMetaData                       As IJDAttributeMetaData
    'Dim oWorkingset                     As IJDWorkingSet
    'Dim oActiveConnection               As IJDConnection
    Dim varInterfaceID                  As Variant
    'Dim oTrader                         As IJDTrader
    Dim oAttributes                     As IJDAttributes
    Dim oInterfaceInfo                  As IJDInterfaceInfo
    Dim iID                             As Variant
    Dim lIndex                          As Long
    Dim ii                              As Long
    Dim oAttr                           As IJDAttribute
    Dim oAttrCol                        As IJDAttributesCol
    Dim oActionAttrsColl              As Collection

    Set oActionAttrsColl = New Collection
    
    'Set oTrader = New Trader
    'Set oWorkingset = oTrader.Service("WorkingSet", vbNullString)
    'Set oActiveConnection = oWorkingset.ActiveConnection
    'Set oMetaData = GetPOM("Model")
    Set oMetaData = oSp
    
    'Get the attributes on the active object
   
    Set oAttributes = oSp

    'iterate through all interfaces and search for the attribute
    For Each varInterfaceID In oAttributes
        Set oInterfaceInfo = oMetaData.InterfaceInfo(varInterfaceID)
        If oInterfaceInfo.IsHardCoded = False Then
            'add the attribute to the collection
             iID = oInterfaceInfo.Type
             Set oAttrCol = oAttributes.CollectionOfAttributes(iID)
             For ii = 1 To oAttrCol.Count
                Set oAttr = oAttrCol.Item(ii)
                oActionAttrsColl.Add oAttr
                Set oAttr = Nothing
            Next ii
            Set oAttrCol = Nothing
        End If
    Next
    
    Set GetActionAttrs = oActionAttrsColl


    Set oMetaData = Nothing
    Set oAttributes = Nothing
    'Set oTrader = Nothing
    'Set oWorkingset = Nothing
    Set oInterfaceInfo = Nothing
    'Set oActiveConnection = Nothing
    Set oActionAttrsColl = Nothing

Exit Function
ErrorHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
End Function

Private Function GetBlockAssignmentStatus(oUnProcessedPartsFolder As IJPlnUnprocessedParts) As String
Const METHOD = "GetBlockAssignmentStatus"
On Error GoTo ErrorHandler
    GetBlockAssignmentStatus = "Completed"
    If Not oUnProcessedPartsFolder Is Nothing Then
        Dim oAssemblyChild As IJAssemblyChild
        Set oAssemblyChild = oUnProcessedPartsFolder
        If Not oAssemblyChild.Parent Is Nothing Then
            If TypeOf oAssemblyChild.Parent Is IJBlock Then
                Dim oPlnBlock As IJPlnBlock
                Set oPlnBlock = oAssemblyChild.Parent
                 If oPlnBlock.AssignmentDelayed = True Then
                    GetBlockAssignmentStatus = "ReAssignment Required"
                 End If
            End If
        End If
    End If
Exit Function
ErrorHandler:
Log.AddFromErr Err, Err.Description, METHOD, MODULE
Err.Clear
End Function
