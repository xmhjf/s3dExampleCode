Attribute VB_Name = "GeneralUtility"
Option Explicit
'-----------------------------------------------------------------------------------------
'  Copyright (C) 2001 - 2011, Intergraph Corporation.  All rights reserved.
'
'
'  File Info:
'      Folder:   SharedContent\Src\StructDetail\Rules\
'      Project:  SDBoundedUSSRule
'      Class:    GeneralUtility
'
'  Abstract:
'       This class module contains the code that:
'       Provides General utitlites that are used by the SDBoundedUSSRule project
'
'  Notes:
'
'
'  History:
'-----------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------'
'
Private Const MODULE = "SharedContent\Src\StructDetail\Rules\SDBoundedUSSRule\GeneralUtility.bas"
'
Public Const C_BaseSide = "Base"
Public Const C_OffsetSide = "Offset"

Public Const INPUTGROUP_Bounded = "Bounded"
Public Const INPUTGROUP_Bounding = "Bounding"
Public Const INPUTGROUP_FlangeCutPoints = "FlangeCutPoints"
'

'********************************************************************
' ' Routine: zMsgBox
'
' Description:
'********************************************************************
Public Sub zMsgBox(sText As String, _
                   Optional sDumArg1 As String, Optional sTitle As String)
On Error Resume Next

Dim iFileNumber
Dim sFileName As String

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

'********************************************************************
' ' Routine: LogError
'
' Description:  default Error Reporter
'********************************************************************
Public Function LogError(oErrObject As ErrObject, _
                         Optional strSourceFile As String = "", _
                         Optional strMethod As String = "", _
                         Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
     
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
End Function

'********************************************************************
' Get_BoundedSymbolGrahicInputs
'   Retrieve the Bounded Symbols Graphic Inputs
'
'
'       In:   IDisp i/f of Symbol Definition
'       Out:
'             IDisp i/f of Bounded Object
'             IDisp i/f of Bounding Object
'
'********************************************************************
Public Sub Get_BoundedSymbolGrahicInputs(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, _
                                        oBoundedObject As Object, _
                                        oBoundingObject As Object, _
                                        oFlangeCutPoints As Object)
Const MT = "Get_BoundedSymbolGrahicInputs"
On Error GoTo ErrorHandler

    'Get the grahic inputs to the Penetration symbol
    Dim DefPlayerEx As IMSSymbolEntities.IJDDefinitionPlayerEx
    Set DefPlayerEx = pSymbolDefinition
    
    Dim oSymbolOcc As IJDSymbol
    Set oSymbolOcc = DefPlayerEx.PlayingSymbol
    
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = pSymbolDefinition
    
    Dim found As Long
    Dim pEnumArg As IJDArgument
    Dim pEnumJDArgument As IEnumJDArgument
    
    Dim oGraphicInput As IMSSymbolEntities.IJDInput
    
    'Get the enum of arguments set by reference on the symbol if any
    If TypeOf oSymbolOcc Is IJDFlavor Then
         ' got a Flavor get the references via the Inputsarg interface
        Dim oInputsArg As IJDInputsArg
        Set oInputsArg = oSymbolOcc
        Set pEnumJDArgument = oInputsArg.GetInputs(igINPUT_ARGUMENTS_SET)
    Else
        ' Get the enum of arguments set by reference on the symbol if any
        Set pEnumJDArgument = oSymbolOcc.IJDReferencesArg.GetReferences()
    End If

    If Not pEnumJDArgument Is Nothing Then
        pEnumJDArgument.Reset
        Do
            pEnumJDArgument.Next 1, pEnumArg, found
            If found = 0 Then Exit Do
            
            Set oGraphicInput = oInputs.Item(pEnumArg.Index)
            If oGraphicInput.Properties <> igINPUT_IS_A_PARAMETER Then
            
                If oGraphicInput.Name = INPUTGROUP_Bounded Then
                    Set oBoundedObject = pEnumArg.Entity
                ElseIf oGraphicInput.Name = INPUTGROUP_Bounding Then
                    Set oBoundingObject = pEnumArg.Entity
               ElseIf oGraphicInput.Name = INPUTGROUP_FlangeCutPoints Then
                  Set oFlangeCutPoints = pEnumArg.Entity
                End If
                
            End If
            
            Set oGraphicInput = Nothing
        Loop
    End If
    
   Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number

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
'           "TOP"       given Port is Profile JXSEC_TOP port
'           "BOTTOM"    given Port is Profile JXSEC_BOTTOM port
'           "WEB_LEFT"  given Port is Profile JXSEC_WEB_LEFT port
'           "WEB_RIGHT" given Port is Profile JXSEC_WEB_RIGHT port
'
'***********************************************************************
Public Function Get_PortFaceType(oPortObject As Object) As String
Const MT = "BoundedDefinitions.GetPortBaseOffsetSide"
On Error GoTo ErrorHandler
    
    On Error Resume Next
    
    Dim lBaseCheck As Long
    Dim lNplusCheck As Long
    Dim lNminusCheck As Long
    Dim lOffsetCheck As Long
    
    Dim lContextID As Long
    Dim lOperatorID As Long
    Dim lOperatationID As Long
    
    Dim ePortType As JS_TOPOLOGY_PROXY_TYPE
    
    Dim sPortSide As String
    Dim eContextID As eUSER_CTX_FLAGS
    Dim oStructPort As IJStructPort
    
    Dim eAxisPortIndex As SPSMemberAxisPortIndex
    Dim oSplitAxisPort As ISPSSplitAxisPort
    
    Dim oMemberFactory As SPSMembers.SPSMemberFactory
    Dim oMemberConnectionServices As SPSMembers.ISPSMemberConnectionServices
    
    ' SP3D Member Object Ports do not implement IJStructPort interface
    sPortSide = ""
    If TypeOf oPortObject Is IJStructPort Then
        Set oStructPort = oPortObject
        eContextID = oStructPort.ContextID
        lBaseCheck = eContextID And CTX_BASE
        lNplusCheck = eContextID And CTX_NPLUS
        lNminusCheck = eContextID And CTX_NMINUS
        lOffsetCheck = eContextID And CTX_OFFSET
        
        If lBaseCheck <> 0 Then
            sPortSide = C_BaseSide
            
        ElseIf lOffsetCheck <> 0 Then
            sPortSide = C_OffsetSide
                
        ElseIf lNplusCheck <> 0 Then
            sPortSide = C_BaseSide
                
        ElseIf lNminusCheck <> 0 Then
            sPortSide = C_OffsetSide
        
        Else
            ' Connecting Port is not Base or Offset port
            ' check if Profile Top/Bottom/Left/Right port
            lOperatorID = oStructPort.OperatorID
            If lOperatorID = JXSEC_TOP Then
                sPortSide = "TOP"
            ElseIf lOperatorID = JXSEC_BOTTOM Then
                sPortSide = "BOTTOM"
            ElseIf lOperatorID = JXSEC_WEB_LEFT Then
                sPortSide = "WEB_LEFT"
            ElseIf lOperatorID = JXSEC_WEB_RIGHT Then
                sPortSide = "WEB_RIGHT"
            Else
                sPortSide = ""
            End If
            
        End If
    
    ElseIf TypeOf oPortObject Is ISPSSplitAxisPort Then
        ' SP3D Member Object Axis Port
        Set oSplitAxisPort = oPortObject
        eAxisPortIndex = oSplitAxisPort.PortIndex
    
        If eAxisPortIndex = SPSMemberAxisStart Then
            sPortSide = C_BaseSide
        ElseIf eAxisPortIndex = SPSMemberAxisEnd Then
            sPortSide = C_OffsetSide
        ElseIf eAxisPortIndex = SPSMemberAxisAlong Then
            sPortSide = ""
        Else
            sPortSide = ""
        End If

    ElseIf TypeOf oPortObject Is IJPort Then
        ' SP3D Member Object Solid Port
        Set oMemberFactory = New SPSMembers.SPSMemberFactory
        Set oMemberConnectionServices = oMemberFactory.CreateConnectionServices
        
        oMemberConnectionServices.GetStructPortInfo oPortObject, ePortType, _
                                                    lContextID, lOperatationID, lOperatorID
        If lOperatorID = JXSEC_TOP Then
            sPortSide = "TOP"
        ElseIf lOperatorID = JXSEC_BOTTOM Then
            sPortSide = "BOTTOM"
        ElseIf lOperatorID = JXSEC_WEB_LEFT Then
            sPortSide = "WEB_LEFT"
        ElseIf lOperatorID = JXSEC_WEB_RIGHT Then
            sPortSide = "WEB_RIGHT"
        Else
            sPortSide = ""
        End If
        
    End If
    
    Get_PortFaceType = sPortSide
    
    Exit Function
    
ErrorHandler:
   Err.Raise LogError(Err, MODULE, MT, "").Number
End Function

'***********************************************************************
' Method:
'    SD_GetBasePort
'
' Abstract: Get Base (Global) Port from Ship Struct Detailing object
'
' Description:
'
' Inputs:
'
' Outputs:
'
'***********************************************************************
Public Function SD_GetBasePort(oConnectable As Object, _
                              lPortCtx As Long) As Object
Const MT = "SD_GetBasePort"
On Error GoTo ErrorHandler
    
    Dim ePortType As JS_TOPOLOGY_FILTER_TYPE
    
    Dim oListOfPorts As IEnumUnknown
    Dim oCollectionPorts As Collection
    
    Dim ConvertUtils As CONVERTUTILITIESLib.CCollectionConversions
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    
    ' Select the type of port to create
    If lPortCtx = CTX_BASE Then
        ePortType = JS_TOPOLOGY_FILTER_SOLID_BASE_LFACE
    ElseIf lPortCtx = CTX_OFFSET Then
        ePortType = JS_TOPOLOGY_FILTER_SOLID_OFFSET_LFACE
    Else
        ePortType = JS_TOPOLOGY_FILTER_LCONNECT_PRT_LATERAL_LFACES
    End If
    
    ' Get the list of ports
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
    oTopologyLocate.GetNamedPorts oConnectable, ePortType, oListOfPorts
    If oListOfPorts Is Nothing Then
        Set SD_GetBasePort = Nothing
        Exit Function
    End If
    
    'Convert the IEnumUnknown to a VB collection that we can use in VB
    Set ConvertUtils = New CONVERTUTILITIESLib.CCollectionConversions
    ConvertUtils.CreateVBCollectionFromIEnumUnknown oListOfPorts, oCollectionPorts

    'Get first port on List
    If oCollectionPorts Is Nothing Then
        Set SD_GetBasePort = Nothing
    ElseIf oCollectionPorts.Count < 1 Then
        Set SD_GetBasePort = Nothing
    Else
        Set SD_GetBasePort = oCollectionPorts.Item(1)
    End If

Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT, "").Number
End Function


