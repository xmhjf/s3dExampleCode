VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StdAssemblyRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'  Copyright (C) 2013, Intergraph Corporation. All rights reserved.
'
'  Project     : SharedContent\Src\Planning\Rules\CommonEntity\ComparisonRule\CommonEntityComparison.vbp
'  File        : StdAssemblyRule.cls
'
'  Description : Common Part Assembly comparison rule
'
'  History     :
'   28th Feb 2013      Arun Prasad     Initial creation
'**************************************************************************************

Option Explicit

Implements IJCommonEntityComparison

Private Const IID_IJMfgParent As String = "{07AE45B9-D88F-4DF2-865D-FEBE67D82E32}"
Private Const IID_IJAssemblyOrientation As String = "{D57B56CC-9120-11D3-806B-0090276F429E}"


Private Sub Class_Initialize()

    sSOURCEFILE = "StdAssemblyRule.cls"
    
    If m_oErrors Is Nothing Then
        Set m_oErrors = CreateObject(ERRORPROGID)
    End If
    
    'How should tolerances be read from the XML?
    'XML should be sent as an input to this class

    m_dDistanceTolerance = 0.0001
    m_dAreaTolerance = 0.001
    m_dVolumeTolerance = 0.001
    m_dWeightTolerance = 0.01
    m_dAngleTolerance = 0.001
    
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

Private Function IJCommonEntityComparison_AreCommon(ByVal pCandidateAssy As Object, ByVal pTargetAssy As Object, strDisplayChecksColl() As String, strKeywordChecksColl() As String, ByVal pManager As Object, strResultMessage As String, saCandCheckValues() As Variant, saTgtCheckValues() As Variant, saStatus() As Variant) As Boolean
Const METHOD = "IJCommonEntityComparison_AreCommon"
On Error GoTo ErrorHandler

    Dim j                           As Long
    Dim oCandidate                  As IJAssembly
    Dim oTarget                     As IJAssembly
    Dim bCommon                     As Boolean
    Dim bIsGeomSame                 As Boolean
    
    
    Dim vCndtPropValue              As Variant
    Dim vTgtPropValue               As Variant
    Dim bAPISuccess                 As Boolean
        
    Dim dCandLength                     As Double
    Dim dCandWidth                      As Double
    Dim dCandHeight                     As Double
    Dim dTargLength                     As Double
    Dim dTargWidth                      As Double
    Dim dTargHeight                     As Double
    
    'Create a Common Part Common Helper

    Dim oCPSCompareHelperEx            As IJPlnCompareHelperEx
    Dim oCPSCompareHelper             As IJPlnCompareHelper
    Dim oCPSAssemblyHelper            As IJPlnAssemblyHelper
    
    Set oCPSCompareHelper = New CPlnAssemblyHelper
    Set oCPSCompareHelperEx = oCPSCompareHelper
    Set oCPSAssemblyHelper = oCPSCompareHelper
    
    'By default the comparison returns true
    bCommon = True
    bIsGeomSame = False
    strResultMessage = vbNullString
    
     'Setting the tolerance values to the helper
    oCPSCompareHelperEx.Tolerance(DistanceTolerance) = m_dDistanceTolerance
    oCPSCompareHelperEx.Tolerance(AreaTolerance) = m_dAreaTolerance
    oCPSCompareHelperEx.Tolerance(VolumeTolerance) = m_dVolumeTolerance
    
    Set oCandidate = pCandidateAssy
    Set oTarget = pTargetAssy
    
    'Set candidate and Target
    oCPSCompareHelper.Candidate = oCandidate
    oCPSCompareHelper.Target = oTarget
    
    ReDim saCandCheckValues(UBound(strKeywordChecksColl)) As Variant
    ReDim saTgtCheckValues(UBound(strKeywordChecksColl)) As Variant
    ReDim saStatus(UBound(strKeywordChecksColl)) As Variant
    
    Dim oPlnIntHelper As IJDPlnIntHelper
    Set oPlnIntHelper = New CPlnIntHelper
    
    Dim oChildren As IJElements
    Dim tempCandValue As Variant
    Dim tempTargValue As Variant
    
    For j = LBound(strKeywordChecksColl) To UBound(strKeywordChecksColl)
        
        bAPISuccess = False
        bCommon = True
        
        Select Case strKeywordChecksColl(j)
                                    
            Case "AssemblyType"
                bAPISuccess = oCPSCompareHelperEx.GetDirectPropertyValues("Type", vCndtPropValue, vTgtPropValue, "IJAssemblyBase")
                
                If bAPISuccess Then
                    If vCndtPropValue <> vTgtPropValue Then
                        bCommon = False
                    End If
                End If
              
                
            Case "AssemblyStage"
                bAPISuccess = oCPSCompareHelperEx.GetDirectPropertyValues("Stage", vCndtPropValue, vTgtPropValue, "IJAssemblyBase")
                
                If bAPISuccess Then
                    If vCndtPropValue <> vTgtPropValue Then
                        bCommon = False
                    End If
                End If
                
                
            Case "WorkCenter"
                bAPISuccess = oCPSCompareHelperEx.GetDirectPropertyValues("WorkCenterName", vCndtPropValue, vTgtPropValue, "IJPlnAssociateCatalog")
                
                If bAPISuccess Then
                    If vCndtPropValue <> vTgtPropValue Then
                        bCommon = False
                    End If
                End If
                
            Case "AssemblyModuleType"
                 bAPISuccess = True
                 Dim oCandModule As IJElements
                 Dim oTargModule As IJElements
                 
                 Set oCandModule = GetRelatedObjects(oCandidate, IID_IJMfgParent, "Module")
                 Set oTargModule = GetRelatedObjects(oTarget, IID_IJMfgParent, "Module")
                 
                If oCandModule.Count > 0 Then
                    vCndtPropValue = TypeName(oCandModule.Item(1))
                Else
                    vCndtPropValue = ""
                End If
                
                If oTargModule.Count > 0 Then
                    vTgtPropValue = TypeName(oTargModule.Item(1))
                Else
                    vTgtPropValue = ""
                End If
                
                 If oCandModule.Count > 0 And oTargModule.Count > 0 Then
                    
                    If StrComp(TypeName(oCandModule.Item(1)), TypeName(oTargModule.Item(1)), vbTextCompare) <> 0 Then
                        bCommon = False
                    End If
                 End If
                 
                 oCandModule.Clear
                 oTargModule.Clear
           
            Case "BuildMethod"
                bAPISuccess = oCPSCompareHelperEx.GetDirectPropertyValues("BuildMethod", vCndtPropValue, vTgtPropValue)
                
                If bAPISuccess Then
                    If vCndtPropValue <> vTgtPropValue Then
                        bCommon = False
                    End If
                End If
                
                
            Case "SlotConnectivity"
                bAPISuccess = oCPSCompareHelperEx.GetDirectPropertyValues("SlotConnectivity", vCndtPropValue, vTgtPropValue)
                
                If bAPISuccess Then
                    If vCndtPropValue <> vTgtPropValue Then
                        bCommon = False
                    End If
                End If
                
                
                
            Case "AssemblyChildrenCount"
                
                Set oChildren = oPlnIntHelper.GetAssemblyChildren(oCandidate, "", False, True)
                
                If Not oChildren Is Nothing Then
                    vCndtPropValue = oChildren.Count
                    Set oChildren = Nothing
                End If
                
                Set oChildren = oPlnIntHelper.GetAssemblyChildren(oTarget, "", False, True)
                bAPISuccess = True
                If Not oChildren Is Nothing Then
                    vTgtPropValue = oChildren.Count
                    Set oChildren = Nothing
                End If
                
                
                If vCndtPropValue <> vTgtPropValue Then
                    bCommon = False
                End If
 
            Case "SubAssemblyCount"
                
                Set oChildren = oPlnIntHelper.GetAssemblyChildren(oCandidate, "IJAssembly", False, True)
                
                If Not oChildren Is Nothing Then
                    vCndtPropValue = oChildren.Count
                    Set oChildren = Nothing
                End If
                
                Set oChildren = oPlnIntHelper.GetAssemblyChildren(oTarget, "IJAssembly", False, True)
                bAPISuccess = True
                If Not oChildren Is Nothing Then
                    vTgtPropValue = oChildren.Count
                    Set oChildren = Nothing
                End If
                
                If vCndtPropValue <> vTgtPropValue Then
                    bCommon = False
                End If
            
            Case "PartsGeometry"  ' Compares all plates, profiles, etc.,
            
                bAPISuccess = oCPSCompareHelperEx.IsGeometrySame(m_dDistanceTolerance)
            
                bCommon = bAPISuccess
                bIsGeomSame = bAPISuccess
                
                    
            Case "Range"
                
                '4 = Minimum Bounding Box
                oPlnIntHelper.GetDimension 4, oCandidate, dCandLength, dCandWidth, dCandHeight
                
                oPlnIntHelper.GetDimension 4, oTarget, dTargLength, dTargWidth, dTargHeight
                                
                bAPISuccess = True
                
                vCndtPropValue = CStr(FormatNumber(dCandLength, 3)) + "; " + CStr(FormatNumber(dCandWidth, 3)) + "; " + CStr(FormatNumber(dCandHeight, 3))
                vTgtPropValue = CStr(FormatNumber(dTargLength, 3)) + "; " + CStr(FormatNumber(dTargWidth, 3)) + "; " + CStr(FormatNumber(dTargHeight, 3))
                
                If Abs(dCandLength - dTargLength) > m_dDistanceTolerance Or _
                   Abs(dCandWidth - dTargWidth) > m_dDistanceTolerance Or _
                   Abs(dCandHeight - dTargHeight) > m_dDistanceTolerance Then
                   
                   bCommon = False
                   
                End If
            
            Case "OrientationVector"
                
                Dim oCndVector As IJDVector
                Dim oTgtVector As IJDVector
                
                bAPISuccess = GetZAxis(oCandidate, oTarget, oCndVector, oTgtVector)
                                
                vCndtPropValue = CStr(oCndVector.x) + "; " + CStr(oCndVector.Y) + "; " + CStr(oCndVector.Z)
                vTgtPropValue = CStr(oTgtVector.x) + "; " + CStr(oTgtVector.Y) + "; " + CStr(oTgtVector.Z)
                
                If vCndtPropValue <> vTgtPropValue Then
                    bCommon = False
                End If
            
            Case "MfgFrameSystem"
                 
                 Dim oCandFrameCS As IJElements
                Dim oTargFrameCS As IJElements
                
                Set oCandFrameCS = GetRelatedObjects(oCandidate, IID_IJAssemblyOrientation, "MfgFrameSystem_ORIG")
                Set oTargFrameCS = GetRelatedObjects(oTarget, IID_IJAssemblyOrientation, "MfgFrameSystem_ORIG")
                
                bAPISuccess = True
                If Not oCandFrameCS Is Nothing And Not oTargFrameCS Is Nothing Then
                  
                    
                        'For reporting in test command.
                        Dim oNI As IJNamedItem
                        
                        If oCandFrameCS.Count > 0 Then
                           Set oNI = oCandFrameCS.Item(1)
                           vCndtPropValue = oNI.Name
                        Else
                           vCndtPropValue = ""
                        End If
                        
                        If oTargFrameCS.Count > 0 Then
                           Set oNI = oTargFrameCS.Item(1)
                           vTgtPropValue = oNI.Name
                        Else
                           vTgtPropValue = ""
                        End If
                        
                    If oCandFrameCS.Count > 0 And oTargFrameCS.Count > 0 Then
                        If Not oCandFrameCS.Item(1) Is oTargFrameCS.Item(1) Then
                            bCommon = False
                        End If
                    End If
                End If
                oCandFrameCS.Clear
                oTargFrameCS.Clear
            
            Case "DryWeight"
                bAPISuccess = GetDryWeight(oCandidate, oTarget, vCndtPropValue, vTgtPropValue)
                               
                If bAPISuccess Then
                        If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dWeightTolerance Then
                            bCommon = False
                        End If
                End If
                
            Case "WetWeight"
                bAPISuccess = GetWetWeight(oCandidate, oTarget, vCndtPropValue, vTgtPropValue)
                
                If bAPISuccess Then
                        If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dWeightTolerance Then
                            bCommon = False
                        End If
                End If
                
                      
'            Case "PR_StageCode"
'                bAPISuccess = oCPSCompareHelperEx.GetDirectPropertyValues("PR_StageCode", vCndtPropValue, vTgtPropValue)
'
'                If bAPISuccess Then
'                    If vCndtPropValue <> vTgtPropValue Then
'                        bCommon = False
'                    End If
'                End If
            
'            Case "PR_WorkCenter"
'                bAPISuccess = oCPSCompareHelperEx.GetDirectPropertyValues("WorkCenter", vCndtPropValue, vTgtPropValue)
'
'                If bAPISuccess Then
'                    If vCndtPropValue <> vTgtPropValue Then
'                        bCommon = False
'                    End If
'                End If
'
'            Case "RA_Count"
'                bAPISuccess = oCPSCompareHelperEx.GetDirectPropertyValues("RA_Count", vCndtPropValue, vTgtPropValue)
'
'                If bAPISuccess Then
'                    If vCndtPropValue <> vTgtPropValue Then
'                        bCommon = False
'                    End If
'                End If
'
'            Case "RoutingAction"
'                bAPISuccess = oCPSCompareHelperEx.GetDirectPropertyValues("RoutingAction", vCndtPropValue, vTgtPropValue)
'
'                If bAPISuccess Then
'                    If vCndtPropValue <> vTgtPropValue Then
'                        bCommon = False
'                    End If
'                End If
'
            Case "GroupMapping"
                bCommon = oCPSAssemblyHelper.AreGroupsCommon

                bAPISuccess = bCommon
              
        End Select
        
        If pManager Is Nothing Then
        
        
            Dim arr() As String
            arr = Split(GetPropertyValues(oCandidate, vCndtPropValue, vTgtPropValue, bAPISuccess, bCommon, strKeywordChecksColl(j), oCPSCompareHelperEx, bIsGeomSame), "@")
            
            saCandCheckValues(j) = arr(0)
            saTgtCheckValues(j) = arr(1)
            saStatus(j) = CStr(arr(2))
        Else
            If bAPISuccess = False Then
                strResultMessage = "Middle Tier API failed while comparing " + strDisplayChecksColl(j)
                bCommon = False
                Exit For
            ElseIf bCommon = False Then
                strResultMessage = strDisplayChecksColl(j) + " not same"
                Exit For
            End If
        End If
        
        vCndtPropValue = Null
        vTgtPropValue = Null
        bCommon = True
    Next j
    
    IJCommonEntityComparison_AreCommon = bCommon
    
    Set oCPSCompareHelperEx = Nothing
    Set oCPSCompareHelper = Nothing
        
Exit Function
ErrorHandler:
    IJCommonEntityComparison_AreCommon = False
    strResultMessage = "Unexpected error while comparing " + strDisplayChecksColl(j)
    MsgBox strResultMessage
End Function

Private Sub IJCommonEntityComparison_SetPropertyValues(saPropertyValues() As String)
Const METHOD = "IJCommonEntityComparison_SetPropertyValues"
On Error GoTo ErrorHandler

    Dim i As Integer
        
    For i = LBound(saPropertyValues) To UBound(saPropertyValues) '2D Array of {tolerance Name, Value }
    
        Select Case saPropertyValues(i, 0)
            Case "DistanceTolerance"
                m_dDistanceTolerance = saPropertyValues(i, 1)
        
            Case "AreaTolerance"
                m_dAreaTolerance = saPropertyValues(i, 1)
        
            Case "VolumeTolerance"
                m_dVolumeTolerance = saPropertyValues(i, 1)
          
            Case "AngleTolerance"
                m_dAngleTolerance = saPropertyValues(i, 1)
        End Select
    Next

Exit Sub
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Sub
