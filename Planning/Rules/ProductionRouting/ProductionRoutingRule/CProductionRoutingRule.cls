VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'  Copyright (C) 2008, Intergraph Corporation. All rights reserved.
'
'  Project     : SharedContent\Src\Planning\Rules\ProductionRouting\
'  File        : ProductionRoutingRule.cls
'
'  Description : ProductionRouting Rule
'
'  History     :
'
'**************************************************************************************

Option Explicit

Private Const MODULE = "PRRule_All"
Private m_lNumActions As Long

Implements IJProductionRoutingRule

Private Type RoutingCode
    Code As String
End Type

Private Type RoutingMachine
    Machine(1 To 50) As RoutingCode
    Name As String
    Type As String
End Type

Private Type RoutingAction
    Action(1 To 50) As RoutingMachine
    Name As String
End Type



Private m_oIJDPOM               As IJDPOM
Private m_oPartMoniker          As IUnknown

Private m_oRoutingActionData(1 To 50) As RoutingAction

Private Sub Class_Initialize()
    Dim strXmlFilename As String
        
    strXmlFilename = GetXMLPath
    
    Readxml (strXmlFilename)
End Sub

Private Sub IJProductionRoutingRule_ComputeActionData(ByVal pPart As Object, pProductionRouting As Object, pAction As Object)
    Call ComputeActionData(pProductionRouting, pPart, pAction)
End Sub

Private Sub IJProductionRoutingRule_ComputeProductionData(ByVal pPart As Object, pProductionRouting As Object)
On Error GoTo ErrorHandler
Const METHOD = "IJProductionRoutingRule_ComputeProductionData"


    Dim oIJDObject             As IJDObject
    
    Set oIJDObject = pPart
    
    Set m_oIJDPOM = oIJDObject.ResourceManager
    Set m_oPartMoniker = m_oIJDPOM.GetObjectMoniker(pPart)
    
    Set oIJDObject = Nothing
   
    Call ComputeproductionData(pProductionRouting, pPart)
    
    
    
    
Exit Sub
ErrorHandler:
     MsgBox Err.Description
End Sub

Private Sub Readxml(strXMLFile As String)
    Const METHOD = "Readxml"
    On Error GoTo ErrorHandler

    
    If strXMLFile <> vbNullString Then
        Dim oNodeList As IXMLDOMNodeList
        Dim oXMLDOC As New MSXML2.DOMDocument60
        Dim oXMLNode As IXMLDOMNode
        Dim oXMLNodeList As IXMLDOMNodeList
        Dim oChildXMLNode As IXMLDOMNode
        Dim oChildXMLNodes As IXMLDOMNodeList
        Dim oSubChildXMLNode As IXMLDOMNode
        Dim oSubChildXMLNodes As IXMLDOMNodeList
        Dim oSCAttribNode As IXMLDOMAttribute
        Dim oSAAttribNode As IXMLDOMAttribute
        Dim strXMLCSName As String
        Dim i As Long
        Dim j As Long
        Dim k As Long

        i = 1
        'Load the file into XMLDOM
        oXMLDOC.Load (strXMLFile)
        Set oXMLNodeList = oXMLDOC.getElementsByTagName("RoutingAction")

        For Each oXMLNode In oXMLNodeList
            For Each oSAAttribNode In oXMLNode.Attributes
                If UCase(oSAAttribNode.nodeName) = UCase("Name") Then
                    m_oRoutingActionData(i).Name = oSAAttribNode.nodeValue
                End If
            Next oSAAttribNode
            j = 1
            For Each oChildXMLNode In oXMLNode.childNodes
                If UCase(oChildXMLNode.nodeName) = UCase("RoutingMachine") Then
                    
                    For Each oSAAttribNode In oChildXMLNode.Attributes
                        If UCase(oSAAttribNode.nodeName) = UCase("Name") Then
                            m_oRoutingActionData(i).Action(j).Name = oSAAttribNode.nodeValue
                        End If
                        If UCase(oSAAttribNode.nodeName) = UCase("Type") Then
                            m_oRoutingActionData(i).Action(j).Type = oSAAttribNode.nodeValue
                        End If
                        
                    Next oSAAttribNode
                    k = 1
                    For Each oSubChildXMLNode In oChildXMLNode.childNodes
                        If UCase(oSubChildXMLNode.nodeName) = UCase("RoutingCode") Then
                            For Each oSCAttribNode In oSubChildXMLNode.Attributes
                                
                                If UCase(oSCAttribNode.nodeName) = UCase("Name") Then
                                    m_oRoutingActionData(i).Action(j).Machine(k).Code = oSCAttribNode.nodeValue
                                End If
                            Next oSCAttribNode
                        End If
                        k = k + 1
                    Next oSubChildXMLNode
                End If
                j = j + 1
            Next oChildXMLNode
            i = i + 1
        Next oXMLNode
    
    
        Set oXMLNodeList = oXMLDOC.getElementsByTagName("RoutingActions")
    
        For Each oXMLNode In oXMLNodeList
            For Each oSAAttribNode In oXMLNode.Attributes
                If UCase(oSAAttribNode.nodeName) = UCase("NumberOfActions") Then
                    m_lNumActions = oSAAttribNode.nodeValue
                    Exit For
                End If
            Next oSAAttribNode
        Next oXMLNode
    End If

              
    Exit Sub
    
ErrorHandler:
     MsgBox Err.Description
End Sub



Private Function IJProductionRoutingRule_GetActions(ByVal pProductionRouting As Object) As Object
Const METHOD = "IJProductionRoutingRule_GetActions"
On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim k As Long
    Dim j As Long
    Dim oActions As Collection
    
    Set oActions = New Collection
    
    For i = 1 To 20
    
        If m_oRoutingActionData(i).Name <> vbNullString Then
            oActions.Add m_oRoutingActionData(i).Name
        End If

    Next i
    
    Set IJProductionRoutingRule_GetActions = oActions
    
Exit Function
    
ErrorHandler:
     MsgBox Err.Description
End Function

Private Function IJProductionRoutingRule_GetCodes(ByVal strAction As String, ByVal strMachine As String) As Object
Const METHOD = "IJProductionRoutingRule_GetCodes"
On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim k As Long
    Dim j As Long
    Dim oCodes As Collection
    
    Set oCodes = New Collection
    
'    Dim oPart As Object
'    Dim strType As String
'
'    Set oPart = m_oIJDPOM.GetObject(m_oPartMoniker)
'
'    If TypeOf oPart Is IJPlatePart Then
'        strType = "Plate"
'    ElseIf TypeOf oPart Is IJProfilePart Then
'        strType = "Profile"
'    End If
    
    For i = 1 To 20
    
        For j = 1 To 20
            For k = 1 To 20
                If UCase(m_oRoutingActionData(i).Name) = UCase(strAction) Then
                    If UCase(m_oRoutingActionData(i).Action(j).Name) = UCase(strMachine) Then
'                        If m_oRoutingActionData(i).Action(j).Type = strType Then
                            If m_oRoutingActionData(i).Action(j).Machine(k).Code <> vbNullString Then
                                oCodes.Add m_oRoutingActionData(i).Action(j).Machine(k).Code
                            End If
'                        End If
                    End If
                End If
            Next k
        Next j
    
    Next i
    
    Set IJProductionRoutingRule_GetCodes = oCodes
    
Exit Function
    
ErrorHandler:
     MsgBox Err.Description
End Function

Private Function IJProductionRoutingRule_GetMachines(ByVal strAction As String, ByVal oPart As Object) As Object
Const METHOD = "IJProductionRoutingRule_GetMachines"
On Error GoTo ErrorHandler

    Dim i As Long
    Dim j As Long
    
    Dim oMachines As Collection

    Set oMachines = New Collection
    
    Dim strType As String
    
       
    If TypeOf oPart Is IJPlatePart Then
        strType = "Plate"
    ElseIf TypeOf oPart Is IJProfilePart Then
        strType = "Profile"
    End If
    
    Dim bPlanarPart                 As Boolean
    Dim oSDPartSupport          As IJPartSupport
    Dim oSDPlatePartSupport     As IJPlatePartSupport
    
    If TypeOf oPart Is IJPlatePart Then
       
        Set oSDPlatePartSupport = New PlatePartSupport
        Set oSDPartSupport = oSDPlatePartSupport
        Set oSDPartSupport.Part = oPart
        
        Dim bIsPlanar As Boolean
        bIsPlanar = False
        If oSDPlatePartSupport.IsShellPlate = False Then
            oSDPlatePartSupport.IsPlanar 0.001, bIsPlanar
            
            If bIsPlanar = True Then ' do not add routing actions to planar plates
                bPlanarPart = True
            End If
        End If
    
    End If
    
    If TypeOf oPart Is IJProfilePart Then
        
        Dim oTopologyLocate As TopologyLocate
        Dim oLandingCrv     As IJWireBody
        Dim oUtils As SGOWireBodyUtilities

        Set oTopologyLocate = New TopologyLocate
        Set oUtils = New SGOWireBodyUtilities

        Set oLandingCrv = oTopologyLocate.GetProfileParentWireBody(oPart)

        'If non-linear , add to excluded parts
        If oUtils.IsLinear(oLandingCrv) <> 0 Then
            bPlanarPart = True
        End If

        Set oLandingCrv = Nothing

    End If
    
       
    If (bPlanarPart = True) And (UCase(strAction) = UCase("Bending")) Then
        oMachines.Add "None"
        Set IJProductionRoutingRule_GetMachines = oMachines
        Exit Function
    End If
    
    
    
    For i = 1 To 20

        For j = 1 To 20
            
            If UCase(m_oRoutingActionData(i).Name) = UCase(strAction) Then
                If (m_oRoutingActionData(i).Action(j).Name <> vbNullString) Then
                    If UCase(strAction) = UCase("Bending") Then
                        If (UCase(m_oRoutingActionData(i).Action(j).Type) = UCase(strType)) Then
                            oMachines.Add m_oRoutingActionData(i).Action(j).Name
                        End If
                    Else
                        oMachines.Add m_oRoutingActionData(i).Action(j).Name
                    End If
                End If
            End If

        Next j

    Next i

    Set IJProductionRoutingRule_GetMachines = oMachines
    Set oPart = Nothing
Exit Function

ErrorHandler:
     MsgBox Err.Description
End Function

Private Function IJProductionRoutingRule_GetStageCodes(ByVal pPart As Object) As Object
On Error GoTo ErrorHandler
Const METHOD = "IJProductionRoutingRule_GetStageCodes"
    
    Dim strXmlFilename As String
   
     
    strXmlFilename = GetXMLPath()
    
    Set IJProductionRoutingRule_GetStageCodes = GetStageCodeColl(strXmlFilename, pPart)
    
    Exit Function
ErrorHandler:
     MsgBox Err.Description
End Function

Private Function GetXMLPath() As String
On Error GoTo ErrorHandler
Const METHOD = "GetXMLPath"
    
    Dim strXMLFile As String
    Dim oContext            As IJContext
   
    Set oContext = GetJContext()
     
    Dim oSRDQuery As IJSRDQuery
    Dim oRuleQuery As IJSRDRuleQuery
    Dim oRule As IJSRDRule
    Dim oRuleObj As Object
    Dim strRuleNames()  As String
    Dim ii As Integer
    Dim strArguments As String
    
    Set oSRDQuery = New SRDQuery
    Set oRuleQuery = oSRDQuery.GetRulesQuery
    
    'Getting the Production Routing xml path from the catalog database
    If Not oRuleQuery Is Nothing Then
        On Error Resume Next
        strRuleNames = oRuleQuery.GetRuleNamesByType(200, 209, True) ' production routing srd rule
        Err.Clear
        On Error GoTo ErrorHandler
        
        If UBound(strRuleNames) >= LBound(strRuleNames) Then
            For ii = LBound(strRuleNames) To UBound(strRuleNames)
                Set oRuleObj = oRuleQuery.GetRule(strRuleNames(ii))
                Exit For
            Next ii
        End If
        
        If Not oRuleObj Is Nothing Then
            Set oRule = oRuleObj
            strArguments = oRule.RuleArguments
            'Getting the Production Routing xml path
        End If
    End If
    
    If strArguments = "" Then
       If Not oRuleQuery Is Nothing Then
            On Error Resume Next
            strRuleNames = oRuleQuery.GetRuleNamesByType(200, 212, True)
            
            If Not UBound(strRuleNames) >= 0 Then
                strRuleNames = oRuleQuery.GetRuleNamesByType(400, 420, True)
            End If
            
            Err.Clear
            On Error GoTo ErrorHandler
            
            If UBound(strRuleNames) >= LBound(strRuleNames) Then
            For ii = LBound(strRuleNames) To UBound(strRuleNames)
                If strRuleNames(ii) = "ProductionRoutingXML Path" Then
                   Set oRule = oRuleQuery.GetRule(strRuleNames(ii))
                   Exit For
                End If
            Next ii
        End If
           
            If Not oRule Is Nothing Then
               strArguments = oRule.ProgId
            End If
       End If
    End If

    'If there is no Production Routing xml path in the database then use the Production Routing xml path from default location
    If strArguments = "" And strArguments = vbNullString Then
        strXMLFile = oContext.GetVariable("OLE_SERVER") & "\Production\ProductionRouting\ProductionRoutingData.xml"
    Else
        strXMLFile = oContext.GetVariable("OLE_SERVER") & strArguments
    End If
    
    GetXMLPath = strXMLFile
     
    Set oRule = Nothing
    Set oRuleObj = Nothing
    Set oSRDQuery = Nothing
    Set oRuleQuery = Nothing
    
    Exit Function
ErrorHandler:
     MsgBox Err.Description
End Function
