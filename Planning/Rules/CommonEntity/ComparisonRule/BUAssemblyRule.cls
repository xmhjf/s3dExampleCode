VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BUAssemblyRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'  Copyright (C) 2013, Intergraph Corporation. All rights reserved.
'
'  Project     : SharedContent\Src\Planning\Rules\CommonEntity\ComparisonRule\CommonEntityComparison.vbp
'  File        : BUAssemblyRule.cls
'
'  Description : Built up Assembly comparison rule
'
'  History     :
'   11th Mar 2013      Chaitanya     Initial creation
'**************************************************************************************

Option Explicit

Implements IJCommonEntityComparison

Private Const IID_IJMfgParent As String = "{07AE45B9-D88F-4DF2-865D-FEBE67D82E32}"
Private Const IID_IJAssemblyOrientation As String = "{D57B56CC-9120-11D3-806B-0090276F429E}"
Private Const IID_IJAssemblyChild As String = "{B447C9B4-FB74-11D1-8A49-00A0C9065DF6}"
Private Const IID_IJPlnProductionRouting As String = "{E9B1D9F6-C687-40BA-85E2-311534F4782A}"


Private Sub Class_Initialize()

    sSOURCEFILE = "BUAssemblyRule.cls"
    
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
    
    'This is needed for PartsGeometry check
    oCPSCompareHelper.Checks (Definition)
    
    ReDim saCandCheckValues(UBound(strKeywordChecksColl)) As Variant
    ReDim saTgtCheckValues(UBound(strKeywordChecksColl)) As Variant
    ReDim saStatus(UBound(strKeywordChecksColl)) As Variant
    
    Dim oPlnIntHelper As IJDPlnIntHelper
    Set oPlnIntHelper = New CPlnIntHelper
    
    Dim oChildren As IJElements
    Dim tempCandValue As Variant
    Dim tempTargValue As Variant
    
    Dim oCandModule As IJElements
    Dim oTargModule As IJElements
    
    Dim oCandDesignMember As ISPSDesignedMember
    Dim oTargDesignMember As ISPSDesignedMember
    
    Dim oCandCSDefiniton As ISPSCrossSection
    Dim oTargCSDefiniton As ISPSCrossSection
    
    Dim oCandMemberType As ISPSMemberType
    Dim oTargMemberType As ISPSMemberType
    
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
                
            Case "ModuleType"
                 bAPISuccess = True
                                 
                 Set oCandModule = GetRelatedObjects(oCandidate, IID_IJMfgParent, "Module")
                 Set oTargModule = GetRelatedObjects(oTarget, IID_IJMfgParent, "Module")
                 
                 If oCandModule.Count > 0 Then 'And oTargModule.Count > 0 Then
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
                        bCommon = True
                    End If
                 End If
                 
           
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
                
            Case "PartsCount"
                
                Set oChildren = oPlnIntHelper.GetAssemblyChildren(oCandidate, "", False, True)
                
                vCndtPropValue = oChildren.Count
                Set oChildren = Nothing
                
                Set oChildren = oPlnIntHelper.GetAssemblyChildren(oTarget, "", False, True)
                
                bAPISuccess = True
                vTgtPropValue = oChildren.Count
                Set oChildren = Nothing
                
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


                 
            'Check cross-section Definition
            Case "SectionName"
                If oCandCSDefiniton Is Nothing Or oTargCSDefiniton Is Nothing Then
                
                    If oCandModule Is Nothing Or oTargModule Is Nothing Then
                        Set oCandModule = GetRelatedObjects(oCandidate, IID_IJMfgParent, "Module")
                        Set oTargModule = GetRelatedObjects(oTarget, IID_IJMfgParent, "Module")
                    End If
                    
                    Set oCandDesignMember = oCandModule.Item(1)
                    Set oTargDesignMember = oTargModule.Item(1)
                    
                    Set oCandCSDefiniton = oCandDesignMember
                    Set oTargCSDefiniton = oTargDesignMember
                End If
                
                If Not oCandCSDefiniton Is Nothing And Not oTargCSDefiniton Is Nothing Then
                
                    'variables for populating test command.
                
                    vCndtPropValue = oCandCSDefiniton.SectionName
                    vTgtPropValue = oTargCSDefiniton.SectionName
                    
                    bAPISuccess = True
                    If StrComp(oCandCSDefiniton.SectionName, oTargCSDefiniton.SectionName, vbTextCompare) = 0 Then
                        bCommon = True
                    End If
                End If
                     
            Case "SectionStandard"
                
                If oCandCSDefiniton Is Nothing Or oTargCSDefiniton Is Nothing Then
                    
                    If oCandModule Is Nothing Or oTargModule Is Nothing Then
                        Set oCandModule = GetRelatedObjects(oCandidate, IID_IJMfgParent, "Module")
                        Set oTargModule = GetRelatedObjects(oTarget, IID_IJMfgParent, "Module")
                    End If
                    
                    Set oCandDesignMember = oCandModule.Item(1)
                    Set oTargDesignMember = oTargModule.Item(1)
                    
                    Set oCandCSDefiniton = oCandDesignMember
                    Set oTargCSDefiniton = oTargDesignMember
                End If
                
                 If Not oCandCSDefiniton Is Nothing And Not oTargCSDefiniton Is Nothing Then
                    
                    'variables for populating test command.
                    vCndtPropValue = oCandCSDefiniton.SectionStandard
                    vTgtPropValue = oTargCSDefiniton.SectionStandard
                    
                    bAPISuccess = True
                    If StrComp(oCandCSDefiniton.SectionStandard, oTargCSDefiniton.SectionStandard, vbTextCompare) = 0 Then
                        bCommon = True
                    End If
                End If

            Case "TypeCategory"
            
                If oCandMemberType Is Nothing Or oTargMemberType Is Nothing Then
                    If oCandModule Is Nothing Or oTargModule Is Nothing Then
                            
                        Set oCandModule = GetRelatedObjects(oCandidate, IID_IJMfgParent, "Module")
                        Set oTargModule = GetRelatedObjects(oTarget, IID_IJMfgParent, "Module")
                    End If
                    
                    
                    Set oCandMemberType = oCandModule.Item(1)
                    Set oTargMemberType = oTargModule.Item(1)
                 End If
                 
                If Not oCandMemberType Is Nothing And Not oTargMemberType Is Nothing Then
                
                    'variables for populating test command.
                    vCndtPropValue = oCandMemberType.TypeCategory
                    vTgtPropValue = oTargMemberType.TypeCategory
                    
                    bAPISuccess = True
                    If StrComp(oCandMemberType.TypeCategory, oTargMemberType.TypeCategory, vbTextCompare) = 0 Then
                        bCommon = True
                    End If
                End If
                
                                
            Case "Type"
                If oCandMemberType Is Nothing Or oTargMemberType Is Nothing Then
                    If oCandModule Is Nothing Or oTargModule Is Nothing Then
                            
                        Set oCandModule = GetRelatedObjects(oCandidate, IID_IJMfgParent, "Module")
                        Set oTargModule = GetRelatedObjects(oTarget, IID_IJMfgParent, "Module")
                    End If
                    
                    Set oCandMemberType = oCandModule.Item(1)
                    Set oTargMemberType = oTargModule.Item(1)
                 End If
                 
                If Not oCandMemberType Is Nothing And Not oTargMemberType Is Nothing Then
                
                    'variables for populating test command.
                    vCndtPropValue = oCandMemberType.Type
                    vTgtPropValue = oTargMemberType.Type
                    
                    bAPISuccess = True
                    If StrComp(oCandMemberType.Type, oTargMemberType.Type, vbTextCompare) = 0 Then
                        bCommon = True
                    End If
                End If
                               
                
            Case "Priority"
                If oCandMemberType Is Nothing Or oTargMemberType Is Nothing Then
                    If oCandModule Is Nothing Or oTargModule Is Nothing Then
                            
                        Set oCandModule = GetRelatedObjects(oCandidate, IID_IJMfgParent, "Module")
                        Set oTargModule = GetRelatedObjects(oTarget, IID_IJMfgParent, "Module")
                    End If
                    
                    
                    Set oCandMemberType = oCandModule.Item(1)
                    Set oTargMemberType = oTargModule.Item(1)
                 End If
                 
                If Not oCandMemberType Is Nothing And Not oTargMemberType Is Nothing Then
                
                    'variables for populating test command.
                    vCndtPropValue = oCandMemberType.Priority
                    vTgtPropValue = oTargMemberType.Priority
                    
                    bAPISuccess = True
                    If StrComp(oCandMemberType.Priority, oTargMemberType.Priority, vbTextCompare) = 0 Then
                        bCommon = True
                    End If
                End If
                
                                
            Case "Range"
                                
                '4 = Minimum Bounding Box
                oPlnIntHelper.GetDimension 4, oCandidate, dCandLength, dCandWidth, dCandHeight
                
                oPlnIntHelper.GetDimension 4, oTarget, dTargLength, dTargWidth, dTargHeight
                                
                bAPISuccess = True
                
                 Dim DecimalNumber() As String
                DecimalNumber = Split(m_dVolumeTolerance, GetDecimalSeparator)
                vCndtPropValue = CStr(FormatNumber(dCandLength, Len(DecimalNumber(1)))) + "; " + CStr(FormatNumber(dCandWidth, Len(DecimalNumber(1)))) + "; " + CStr(FormatNumber(dCandHeight, Len(DecimalNumber(1))))
                vTgtPropValue = CStr(FormatNumber(dTargLength, Len(DecimalNumber(1)))) + "; " + CStr(FormatNumber(dTargWidth, Len(DecimalNumber(1)))) + "; " + CStr(FormatNumber(dTargHeight, Len(DecimalNumber(1))))
                
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
                
             Case "PartsGeometry"  ' Compares all plates, profiles, etc.,
            
                bAPISuccess = oCPSCompareHelper.AreCommon
            
                bCommon = bAPISuccess
                
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
            'Populate test command outputs
            
            Dim arr() As String
            arr = Split(GetPropertyValues(oCandidate, vCndtPropValue, vTgtPropValue, bAPISuccess, bCommon, strKeywordChecksColl(j), oCPSCompareHelperEx), "@")
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
    Set oCPSAssemblyHelper = Nothing
    
Exit Function
ErrorHandler:
    IJCommonEntityComparison_AreCommon = False
    strResultMessage = "Unexpected error while comparing " + strDisplayChecksColl(j)
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
        
        Case "WeightTolerance"
            m_dWeightTolerance = saPropertyValues(i, 1)
            
        Case "AngleTolerance"
            m_dAngleTolerance = saPropertyValues(i, 1)
    End Select
Next

Exit Sub
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Sub







