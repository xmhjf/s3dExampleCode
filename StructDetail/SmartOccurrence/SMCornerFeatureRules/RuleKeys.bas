Attribute VB_Name = "RuleKeys"
'*******************************************************************
'
'Copyright (C) 2013-14 Intergraph Corporation. All rights reserved.
'
'File : RuleKeys.cls
'
'Author : Alligators
'
'Description :
'
'
'History :
'    6/Nov/2013 - vb/svsmylav
'          DI-CP-240506 'UpdateACsForEdgePortsOfCF' method is called conditinally, not for seam movement case.
'
'    21/Nov/2013 - svsmylav
'          TR-234194 Corrected 'UpdateACsForEdgePortsOfCF' to remove seam-movement case check. Also,
'                    Rangebox of the CF is obtained by using 'GetFeatureContourRngBox' method.
'    14/Feb/2014 - svsmylav
'         TR-245505(DM-248838 for v2015): Modified code to use new method 'ForceUpdateOnMemberObjects'
'                    in UpdateACsForEdgePortsOfCF.
'
'*****************************************************************************

Option Explicit

Public Const INPUT_PORT1FACE = "LogicalFace"
Public Const INPUT_PORT2EDGE = "Support1"
Public Const INPUT_PORT3EDGE = "Support2"

'Global string constants for questions and answers
Public Const gsDrainage = "Drainage on Part"
Public Const gsCornerFlip = "Corner Feature Orientation"
Public Const gsPlacement = "Corner Placement"
Public Const gsCrackArrest = "Arrest Stress Cracking"
Public Const gsApplyTreatment = "ApplyTreatment"
Public Const gsSeamAdjustment = "EnforceSeamAdjustment"
Public Const SEAM_SEARCHDISTANCE As Double = 0.015

Public Const IJTestAttributes As String = "IJTestAttributes"
Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\RuleKeys.bas"
Public Const IID_IJProxyMember = "{C1FD8CAF-9767-11D1-9425-0060973D4777}"



'
' This method writes IJWeldingSymbol attributes
'
Public Sub SetAttributeOnInterface( _
                     ByVal oSmartOcc As IJSmartOccurrence, ByVal oInterfaceName As String, _
                     ByVal oAttrNAme As String, ByVal mAttrValue As Variant)
                                               
   On Error GoTo ErrorHandler

   Dim oAttributes             As IMSAttributes.IJDAttributes
   Dim vInterfaceType          As Variant
   Dim nAttributeSet           As Integer
    
   Dim nAttributeCount As Double
   nAttributeCount = 1
    
   nAttributeSet = 0
   Set oAttributes = oSmartOcc
   For Each vInterfaceType In oAttributes
      Dim oAttributesCol          As IJDAttributesCol

      Set oAttributesCol = oAttributes.CollectionOfAttributes(vInterfaceType)
      If oAttributesCol.Count > 0 Then
         Dim oAttribute       As IJDAttribute
         Dim oAttributeInfo   As IJDAttributeInfo
         Dim oInterface       As IJDInterfaceInfo

         For Each oAttribute In oAttributesCol
            Dim nAttributeIndex As Integer
            
            Set oAttributeInfo = oAttribute.AttributeInfo
            If LCase(oAttributeInfo.interfaceName) = LCase(oInterfaceName) Then
               For nAttributeIndex = 1 To nAttributeCount
                  If LCase(oAttributeInfo.Name) = LCase(oAttrNAme) Then
                     oAttribute.value = mAttrValue
                     Exit For
                  End If
               Next
            End If
            
         Next
      End If
      
   Next
   
   Exit Sub
   
ErrorHandler:
   Err.Raise Err.Number, MODULE & "WeldSymbols::GetAttributesOnIJWeldingSymbol"
End Sub


'
' This method writes IJWeldingSymbol attributes
'
Public Sub GetAttributeOnInterface( _
                     ByVal oSmartOcc As IJSmartOccurrence, ByVal oInterfaceName As String, _
                     ByVal oAttrNAme As String, ByRef mAttrValue As Variant)
                                               
   On Error GoTo ErrorHandler

   Dim oAttributes             As IMSAttributes.IJDAttributes
   Dim vInterfaceType          As Variant
   Dim nAttributeSet           As Integer
    
   Dim nAttributeCount As Double
   nAttributeCount = 1
    
   nAttributeSet = 0
   Set oAttributes = oSmartOcc
   For Each vInterfaceType In oAttributes
      Dim oAttributesCol          As IJDAttributesCol

      Set oAttributesCol = oAttributes.CollectionOfAttributes(vInterfaceType)
      If oAttributesCol.Count > 0 Then
         Dim oAttribute       As IJDAttribute
         Dim oAttributeInfo   As IJDAttributeInfo
         Dim oInterface       As IJDInterfaceInfo

         For Each oAttribute In oAttributesCol
            Dim nAttributeIndex As Integer
            
            Set oAttributeInfo = oAttribute.AttributeInfo
            If LCase(oAttributeInfo.interfaceName) = LCase(oInterfaceName) Then
               For nAttributeIndex = 1 To nAttributeCount
                  If LCase(oAttributeInfo.Name) = LCase(oAttrNAme) Then
                      mAttrValue = oAttribute.value
                     Exit For
                  End If
               Next
            End If
            
         Next
      End If
      
   Next
   
   Exit Sub
   
ErrorHandler:
   Err.Raise Err.Number, MODULE & "WeldSymbols::GetAttributesOnIJWeldingSymbol"
End Sub

Public Sub GetCFlengths(sCornerFeatureName As String, oCornerFeature As StructDetailObjectsex.CornerFeature, strCrackArrest As String, strDrainage As String, dULength As Double, dVLength As Double)
On Error GoTo LengthSelector

     Dim oPart As Object
     Set oPart = oCornerFeature.GetPartObject
     
     Dim oProfile As New StructDetailObjects.ProfilePart
     Dim oBeam As New StructDetailObjects.BeamPart
     Dim oMember As New StructDetailObjects.MemberPart
     
     Dim bPlate As Boolean
     Dim mHeight As Double
     
     Dim oGapLengthGap As Double
     Dim oGapHeightGap As Double
     Dim oGapTolerance As Double
    
     
     If TypeOf oPart Is IJPlate Then
        bPlate = True
     ElseIf TypeOf oPart Is IJProfile Or TypeOf oPart Is ISPSMemberPartPrismatic Or TypeOf oPart Is IJBeam Then
        If TypeOf oPart Is IJProfile Then
            Set oProfile.object = oPart
            mHeight = oProfile.Height
        ElseIf TypeOf oPart Is IJBeam Then
            Set oBeam.object = oPart
            mHeight = oBeam.Height
        ElseIf TypeOf oPart Is ISPSMemberPartPrismatic Then
            Set oMember.object = oPart
            mHeight = oMember.Height
        End If
    End If
    
    Dim oSmOcc As IJSmartOccurrence
    Set oSmOcc = oCornerFeature.object
    
    Dim oSmartItem As IJSmartItem
    Dim oSymbol As IMSSymbolEntities.IJDSymbol
    
    If Not oSmOcc.Item = "" Then
        Set oSmartItem = oSmOcc.SmartItemObject
        Set oSymbol = oCornerFeature.object
    End If
    
    Dim mUlength As Double: Dim mVlength As Double: Dim mRadius As Double
    mUlength = 0: mVlength = 0: mRadius = 0
    
    If Not sCornerFeatureName = "" Then
        
        mUlength = GetParameterValue(oSymbol, "Ulength")
        mVlength = GetParameterValue(oSymbol, "Vlength")
        mRadius = GetParameterValue(oSymbol, "Radius")
    
        If InStr(1, sCornerFeatureName, "longscallop", vbTextCompare) > 0 Then
             dULength = mUlength
             dVLength = mRadius
        ElseIf InStr(1, sCornerFeatureName, "crackarrest", vbTextCompare) > 0 Then
            dULength = mUlength
            dVLength = mUlength
        ElseIf InStr(1, (sCornerFeatureName), "snipe", vbTextCompare) > 0 Then
            dULength = mUlength
            dVLength = mVlength
        ElseIf InStr(1, sCornerFeatureName, "scallop", vbTextCompare) > 0 Then
            dULength = mRadius
            dVLength = mRadius
        End If
        
    Else 'case where there is no Corner feature alreay existing at the location
        
        If TypeOf oPart Is IJPlate Then
            If IsCollar(oPart) Then
                dULength = 0.015
                dVLength = 0.015
            ElseIf IsBracket(oPart) Then
                dULength = 0.05
                dVLength = 0.05
            Else
                Dim oPlate As New StructDetailObjects.PlatePart
                Set oPlate.object = oPart
                Select Case oPlate.Tightness
                    Case NonTight
                        Select Case strCrackArrest
                            Case "Yes" 'long scallop item
                                oGapTolerance = 0.001 'will not find gaps < 1 mm
                                If oCornerFeature.MeasureCornerGap(oGapTolerance, _
                                                                           oGapLengthGap, _
                                                                           oGapHeightGap) Then
                                    If oGapLengthGap > 0# Then
                                        dULength = oGapLengthGap + 0.015
                                    ElseIf oGapHeightGap > 0# Then
                                        dULength = oGapHeightGap + 0.015
                                    End If
                                    dVLength = 0.05
                                Else    ' long scallop item
                                    dULength = 0.1
                                    dVLength = 0.05
                                End If
                            Case "No"   'scallop item
                                dVLength = 0.05
                                dULength = 0.05
                            End Select
    
                    Case Else   'scallop item
                        dVLength = 0.05
                        dULength = 0.05
                    End Select
            End If
    
        ElseIf TypeOf oPart Is IJProfile Or TypeOf oPart Is ISPSMemberPartPrismatic Or TypeOf oPart Is IJBeam Then
        
            Select Case strDrainage
                Case "Yes"
                    If oCornerFeature.MeasureCornerGap(0.001, _
                                                oGapLengthGap, oGapHeightGap) Then
                            If oGapLengthGap > 0# Then
                                dULength = oGapLengthGap + 0.015
                            ElseIf oGapHeightGap > 0# Then
                                dULength = oGapHeightGap + 0.015
                            End If
                            dVLength = 0.05
                    Else
                        Select Case mHeight
                            Case 0 To 0.2
                                dULength = 0.035
                                dVLength = 0.035
                            Case 0.2 To 0.4
                                dULength = 0.05
                                dVLength = 0.05
                            Case Is >= 0.4
                                dULength = 0.075
                                dVLength = 0.075
                        End Select
                    End If
                Case "no"
                    Select Case mHeight
                        Case 0 To 0.2
                            dULength = 0.035
                            dVLength = 0.035
                        Case 0.2 To 0.4
                            dULength = 0.05
                            dVLength = 0.05
                        Case Is >= 0.4
                            dULength = 0.075
                            dVLength = 0.075
                    End Select
            End Select
        End If
    End If
    
    Exit Sub
    
LengthSelector:
    dULength = 0.05
    dVLength = 0.05
End Sub


Public Function GetACPortByBruteForce(oPort As IJPort) As IJPort
On Error GoTo ErrorHandler

    
    Set GetACPortByBruteForce = Nothing
    
    Dim oStructPort As IJStructPort
    Set oStructPort = oPort
    
    Dim nConnections As Long
    Dim connData() As ConnectionData
    
    Dim oHelper As New StructDetailObjects.Helper
    
    oHelper.Object_AppConnections oPort.Connectable, AppConnectionType_Assembly, nConnections, connData
    
    'nConnections = GetAllConnectables(oPort.Connectable, AppConnectionType_Assembly, connData)

    Dim i As Long
    For i = 1 To nConnections
    
        Dim oConnStructPort As IJStructPort
        Set oConnStructPort = connData(i).ConnectingPort
    
        If TypeOf oPort.Connectable Is IJProfile Then
            Dim portPrimaryContext As Long
            Dim connPortPrimaryContext As Long
        
            portPrimaryContext = oStructPort.ContextID And (CTX_BASE Or CTX_OFFSET Or CTX_LATERAL)
            connPortPrimaryContext = oConnStructPort.ContextID And (CTX_BASE Or CTX_OFFSET Or CTX_LATERAL)
        
            If (portPrimaryContext = connPortPrimaryContext) Then
                Set GetACPortByBruteForce = oConnStructPort
                Exit Function
            End If
        ElseIf oConnStructPort.ContextID = oStructPort.ContextID And _
               oConnStructPort.operationID = oStructPort.operationID And _
               oConnStructPort.operatorID = oStructPort.operatorID And _
               oConnStructPort.SectionID = oStructPort.SectionID Then
               
               Set GetACPortByBruteForce = oConnStructPort
               Exit Function
        End If
    Next i
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetACPortByForce").Number

End Function


Public Function GetPCPortByBruteForce(oPort As IJPort) As IJPort
On Error GoTo ErrorHandler
    Set GetPCPortByBruteForce = Nothing
    
    Dim oStructPort As IJStructPort
    Set oStructPort = oPort
    
    Dim nConnections As Long
    Dim connData() As ConnectionData
    
    Dim oHelper As New StructDetailObjects.Helper
    
    oHelper.Object_AppConnections oPort.Connectable, AppConnectionType_Physical, nConnections, connData
    
    'nConnections = GetAllConnectables(oPort.Connectable, AppConnectionType_Assembly, connData)

    Dim i As Long
    For i = 1 To nConnections
    
        Dim oConnStructPort As IJStructPort
        Set oConnStructPort = connData(i).ConnectingPort
    
        If TypeOf oPort.Connectable Is IJProfile Then
            Dim portPrimaryContext As Long
            Dim connPortPrimaryContext As Long
        
            portPrimaryContext = oStructPort.ContextID And (CTX_BASE Or CTX_OFFSET Or CTX_LATERAL)
            connPortPrimaryContext = oConnStructPort.ContextID And (CTX_BASE Or CTX_OFFSET Or CTX_LATERAL)
        
            If (portPrimaryContext = connPortPrimaryContext) Then
                Set GetPCPortByBruteForce = oConnStructPort
                Exit Function
            End If
        ElseIf oConnStructPort.ContextID = oStructPort.ContextID And _
               oConnStructPort.operationID = oStructPort.operationID And _
               oConnStructPort.operatorID = oStructPort.operatorID And _
               oConnStructPort.SectionID = oStructPort.SectionID Then
               
               Set GetPCPortByBruteForce = oConnStructPort
               Exit Function
        End If
    Next i
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetPCPortByBruteForce").Number

End Function

Public Sub UpdateACsForEdgePortsOfCF(oFeature As IJStructFeature)
On Error GoTo ErrorHandler
      Dim oStructFeatUtils As IJSDFeatureAttributes
      Dim oFacePort As Object
      Dim oEdgePort1 As Object
      Dim oEdgePort2 As Object
      
      Set oStructFeatUtils = New SDFeatureUtils
        
      oStructFeatUtils.get_CornerCutInputsEx oFeature, _
                                           oFacePort, _
                                           oEdgePort1, _
                                           oEdgePort2
    'get the PCs that both edge ports are involved and recompute them
    If Not TypeOf oFacePort Is IJPort Or _
       Not TypeOf oEdgePort1 Is IJPort Or _
       Not TypeOf oEdgePort2 Is IJPort Then
       
       Exit Sub
    End If
    
    Dim ppConnections1 As IJElements
    Dim ppConnections2 As IJElements
    
    Dim oConnection As GSCADCreateModifyUtilities.IJStructConnection
    Dim oConnObj As Object
    
    Dim oConnPort1 As IJPort
    Dim oConnPort2 As IJPort
    
    Set oConnPort1 = GetACPortByBruteForce(oEdgePort1)

    If oConnPort1 Is Nothing Then
       Set oConnPort1 = oEdgePort1
    End If
    
    Set oConnPort2 = GetACPortByBruteForce(oEdgePort2)
    
    If oConnPort2 Is Nothing Then
       Set oConnPort2 = oEdgePort2
    End If
    
    oConnPort1.enumConnections ppConnections1, ConnectionAssembly, ConnectionStandard
    
    oConnPort2.enumConnections ppConnections2, ConnectionAssembly, ConnectionStandard
    
    'Get Rangebox of corner feature's projected contour
    Dim gCornerFeatureRngBox As GBox
    gCornerFeatureRngBox = GetFeatureContourRngBox(oFeature)
    
    Dim nConnections As Long
    If Not ppConnections1 Is Nothing Then
        For nConnections = 1 To ppConnections1.Count
            Set oConnObj = ppConnections1.Item(nConnections)
            If TypeOf oConnObj Is IJAssemblyConnection Then
                If NeedForceUpdateofAC(oConnObj, gCornerFeatureRngBox) Then
                    ForceUpdateOnMemberObjects oConnObj
                    Exit For
                End If
            End If
        Next
    End If
    
    If Not ppConnections2 Is Nothing Then
        For nConnections = 1 To ppConnections2.Count
            Set oConnObj = ppConnections2.Item(nConnections)
            If TypeOf oConnObj Is IJAssemblyConnection Then
                If NeedForceUpdateofAC(oConnObj, gCornerFeatureRngBox) Then
                    ForceUpdateOnMemberObjects oConnObj
                    Exit For
                End If
            End If
        Next
    End If
    
CleanUp:
    Set oStructFeatUtils = Nothing
    Set oFacePort = Nothing
    Set oEdgePort1 = Nothing
    Set oEdgePort2 = Nothing
    Set ppConnections1 = Nothing
    Set ppConnections2 = Nothing
    Set oConnection = Nothing
    Set oConnObj = Nothing
    Set oConnPort1 = Nothing
    Set oConnPort2 = Nothing
    
    Exit Sub
ErrorHandler:
   Err.Raise Err.Number, MODULE & "UpdateACsForEdgePortsOfCF"
    
End Sub

'********************************************************************
' Routine: CheckPartClassExist
' Description:  Check if the given PartClass exist in Catalog
'
'********************************************************************
Public Function CheckPartClassExist(sPartClassName As String) As Boolean
On Error GoTo ErrorHandler
    Const sMETHOD = "CheckPartClassExist"
    On Error GoTo ErrorHandler
    
    Dim oCatalogQuery As IJSRDQuery
    Dim oSmartQuery As IJSmartQuery
    Dim oPartClass As IJSmartClass
    
    Set oCatalogQuery = New SRDQuery
    Set oSmartQuery = oCatalogQuery
    
    ' Query for SmartClass... Check if exist
    On Error Resume Next
    Set oPartClass = oSmartQuery.GetClassByName(sPartClassName)
    On Error GoTo ErrorHandler
    
    If Not oPartClass Is Nothing Then
        CheckPartClassExist = True
    End If
    
    Set oPartClass = Nothing
    Set oSmartQuery = Nothing
    Set oCatalogQuery = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
    
End Function


'***********************************************************************
' METHOD:   Get Parameter Value
'
' DESCRIPTION:  Given the parameter name on a symbol return the value
'               of the parameter
'
'***********************************************************************

Public Function GetParameterValue(oSymbol As IMSSymbolEntities.IJDSymbol, strParameterName As String) As Variant
On Error GoTo ErrorHandler
  Const sMETHOD = "GetParameterValue"
    
    ' Get the symbol definition
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oSymbolDefinition = oSymbol.IJDSymbolDefinition(0)
            
    ' Get the symbol inputs
    Dim oSymbolInputs As IMSSymbolEntities.IJDInputs
    Set oSymbolInputs = oSymbolDefinition.IJDInputs
    
    ' Get the inputs count
    Dim symInputCount As Long
    symInputCount = oSymbolInputs.InputCount
        
    Dim nInputIndex As Long
    Dim oAnInput As IJDInput
     
    Dim bFoundParameter As Boolean
    bFoundParameter = False
    
    ' Loop through the inputs to find the paramter matching the name
    For nInputIndex = 1 To symInputCount
            
        If bFoundParameter = True Then
            Exit For
        End If
        
        Set oAnInput = oSymbolInputs.GetInputByIndex(nInputIndex)
            
        If oAnInput.IsPropertySet(igINPUT_IS_A_PARAMETER) = True And _
           oAnInput.IsPropertySet(igDESCRIPTION_INVALIDED) = False And _
           oAnInput.Name = strParameterName Then
            
            Dim argIndex As Long
            Dim oOccParameters As IMSSymbolEntities.IJDArguments
            Dim oOccArg As IMSSymbolEntities.IJDArgument
            
            Set oOccParameters = oSymbol.IJDValuesArg.GetValues(igINPUT_ARGUMENTS_MERGE)
                    
            For Each oOccArg In oOccParameters
                        
                If oOccArg.Index = oAnInput.Index Then
                    
                    Dim oParVal As IMSSymbolEntities.IJDParameterContent
                    Set oParVal = oOccArg.Entity
                                  
                    If oParVal.Type = igValue Then
                        GetParameterValue = oParVal.UomValue
                    Else
                        GetParameterValue = oParVal.String
                    End If
                    
                    bFoundParameter = True
                            
                    Exit For
                End If
                
                Set oOccArg = Nothing
            Next
        End If
        
        Set oAnInput = Nothing
    Next

    Exit Function
    
ErrorHandler:
 Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

'
' This method will check wheheter RootCornerCOllar functionality exists in catalog
'
Public Function IsCornerCollarFunctionalityBulkloaded() As Boolean
Const METHOD = "IsCornerCollarFunctionalityBulkloaded"

    On Error GoTo ErrorHandler
     Dim sMsg As String
    IsCornerCollarFunctionalityBulkloaded = False
    'Before proceeding further check if "Corner COllar Functionality is bulkloaded"
    'if not then exit right away before proceeding further.
    If CheckPartClassExist("RootCornerCollar") Then
        '!!! if the RootCornerCollar smart class is not bulkloaded
        'Then this functionality(Corner Collars) cannot be used.
        'This check is needed for Bulkload optional as users might or might not bulkalod
        'some new smart class as per their requirement
        IsCornerCollarFunctionalityBulkloaded = True
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

'*************************************************************************
'Function
'   NeedForceUpdateofAC
'
'Abstract
'   Given the AC and corner feature rangebox, this function determines if Force update of AC is needed.
'
'Inputs
'   oAssyConn As Object
'   ogCFRngBox As GBox
'
'Return
'   Value is 'True' if the AC needs force-update, otherwise 'False'.

'Exceptions
'
'***************************************************************************
Public Function NeedForceUpdateofAC(oAssyConn As Object, gCFRngBox As GBox) As Boolean
  On Error GoTo ErrorHandler

    Dim sMETHOD As String
    sMETHOD = "NeedForceUpdateofAC"
    
    NeedForceUpdateofAC = False 'Initialize

    'Compare Rangebox of AC and Rangebox of the corner feature
    Dim oRange As IJRangeAlias
    Dim gACRngBox As GBox
    Set oRange = oAssyConn
    gACRngBox = oRange.GetRange()
    
    'Rejection criteria: if any of following six conditions is true, AC is out of the CF range, so exit
    If gACRngBox.m_low.x > gCFRngBox.m_high.x Then GoTo CleanUp
    If gACRngBox.m_high.x < gCFRngBox.m_low.x Then GoTo CleanUp
    
    If gACRngBox.m_low.y > gCFRngBox.m_high.y Then GoTo CleanUp
    If gACRngBox.m_high.y < gCFRngBox.m_low.y Then GoTo CleanUp
    
    If gACRngBox.m_low.z > gCFRngBox.m_high.z Then GoTo CleanUp
    If gACRngBox.m_high.z < gCFRngBox.m_low.z Then GoTo CleanUp
        
    'Below code identifies if AC-low or AC-high is within the CF rangebox
    Dim oACLowRngVec As IJDVector
    Set oACLowRngVec = New dVector
    
    Dim oACHighRngVec As IJDVector
    Set oACHighRngVec = New dVector
    oACLowRngVec.Set gACRngBox.m_low.x - gCFRngBox.m_low.x, gACRngBox.m_low.y - gCFRngBox.m_low.y, gACRngBox.m_low.z - gCFRngBox.m_low.z
    oACHighRngVec.Set gCFRngBox.m_high.x - gACRngBox.m_high.x, gCFRngBox.m_high.y - gACRngBox.m_high.y, gCFRngBox.m_high.z - gACRngBox.m_high.z

    Dim bACLowRngGtThanCFRng As Boolean
    Dim bACHighRngLessThanCFRng As Boolean
    
    bACLowRngGtThanCFRng = False
    If Sgn(oACLowRngVec.x) = 1 And Sgn(oACLowRngVec.y) = 1 And Sgn(oACLowRngVec.z) = 1 Then
        bACLowRngGtThanCFRng = True
    End If

    bACHighRngLessThanCFRng = False
    If Sgn(oACHighRngVec.x) = 1 And Sgn(oACHighRngVec.y) = 1 And Sgn(oACHighRngVec.z) = 1 Then
            bACHighRngLessThanCFRng = True
    End If
    
    If (bACLowRngGtThanCFRng Or bACHighRngLessThanCFRng) Then
        'Atleast one end of AC rangebox is within the CF rangebox (other cases are eliminated from
        'further processing)
        
        Dim oACParent As IJDesignParent
        Dim oChildrenOfAC As IJDObjectCollection
    
        Set oACParent = oAssyConn
        oACParent.GetChildren oChildrenOfAC
        If Not oChildrenOfAC Is Nothing Then
            If bACLowRngGtThanCFRng And bACHighRngLessThanCFRng Then
                'Case 1: AC rangebox is completely within the CF rangebox and AC has child (a PC)
                ' => we need to Force-update AC: since CF removed the material, PC creation is not possible.
                If oChildrenOfAC.Count = 1 Then
                    Dim oTestObj As Object
                    For Each oTestObj In oChildrenOfAC
                        If TypeOf oTestObj Is IJAppConnection Then
                            Dim oAppConn As IJAppConnectionType
                            Set oAppConn = oTestObj
                            If oAppConn.Type = ConnectionPhysical Then NeedForceUpdateofAC = True
                        End If
                    Next
                End If
            ElseIf oChildrenOfAC.Count = 0 Then
                'Case 2: One end of AC rangebox is within the CF rangebox, but the other end extends beyond
                ' the CF rangebox, and AC has NO PC. => we need to Force-update AC to get PC.
                NeedForceUpdateofAC = True
            End If
        End If
    End If
    
CleanUp:
    Set oRange = Nothing
    
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

