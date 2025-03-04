VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StdMemberRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'  Copyright (C) 2013, Intergraph Corporation. All rights reserved.
'
'  Project     : SharedContent\Src\Planning\Rules\CommonEntity\ComparisonRule\CommonEntityComparison.vbp
'  File        : StdMemberRule.cls
'
'  Description : Common Part Standard Members comparison rule
'
'  History     :
'   31st Jan 2013      Chaitanya     Initial creation
'**************************************************************************************

Option Explicit

Private Const IID_IJAssemblyChild As String = "{B447C9B4-FB74-11D1-8A49-00A0C9065DF6}"
Private Const IID_IJPlnProductionRouting As String = "{E9B1D9F6-C687-40BA-85E2-311534F4782A}"

Implements IJCommonEntityComparison
Implements IJStandardEntityComparison

Private Sub Class_Initialize()

    sSOURCEFILE = "StdMemberRule.cls"
    
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

Private Function IJCommonEntityComparison_AreCommon(ByVal pCandidate As Object, ByVal pTarget As Object, strDisplayChecksColl() As String, strKeywordChecksColl() As String, ByVal pManager As Object, strResultMessage As String, saCandCheckValues() As Variant, saTgtCheckValues() As Variant, saStatus() As Variant) As Boolean
                                                    
Const METHOD = "IJCommonEntityComparison_AreCommon"
On Error GoTo ErrorHandler

    Dim i                           As Long
    Dim j                           As Long
    Dim oCandidate                  As ISPSMemberPartPrismatic
    Dim oTarget                     As ISPSMemberPartPrismatic
    Dim bCalledByTestCmd            As Boolean
    Dim bCommon                     As Boolean
    Dim oCndtMaterial               As IJDMaterial
    Dim oTgtMaterial                As IJDMaterial
    
    Dim oConnectedObjPairs          As IJElements
    Dim oCndtConnectedObjects       As IJElements
    Dim oTgtConnectedObjects        As IJElements
    Dim oCndtProdRouting            As Object
    Dim oTgtProdRouting             As Object
    Dim oCndtProdRtgActions         As IJElements
    Dim oTgtProdRtgActions          As IJElements
    Dim oMatchedPCs                 As IJElements
    Dim bPCsAlreadyMatched          As Boolean
    
    Dim vCndtPropValue              As Variant
    Dim vTgtPropValue               As Variant
    Dim bAPISuccess                 As Boolean
    Dim bIsGeomSame                 As Boolean
    Dim strBevelInfo                As String
        
    bCommon = True
    bIsGeomSame = False
    bPCsAlreadyMatched = False
    m_bGeometriesCompared = False   'Indicates that the geometries are not compared yet.
    strResultMessage = vbNullString
    
    If pManager Is Nothing And m_bStandardPartRule = False Then
        bCalledByTestCmd = True
    Else
        bCalledByTestCmd = False
    End If

    Dim oCPSCommonHelper            As IJPlnCompareHelperEx
    Dim oCPSMemberHelper             As IJPlnCompareHelper
    
    Set oCPSMemberHelper = New CPlnMemberHelper
    Set oCPSCommonHelper = oCPSMemberHelper
    
    Set oCandidate = pCandidate
    Set oTarget = pTarget
    
    ReDim saCandCheckValues(UBound(strKeywordChecksColl)) As Variant
    ReDim saTgtCheckValues(UBound(strKeywordChecksColl)) As Variant
    ReDim saStatus(UBound(strKeywordChecksColl)) As Variant
    
    'Set candidate and Target
    oCPSMemberHelper.Candidate = pCandidate
    oCPSMemberHelper.Target = pTarget
        
    'Get Cand & Targ Cross-section definitions.
    Dim oCandCS As ISPSCrossSection
    Dim oTargCS As ISPSCrossSection
    
    Set oCandCS = oCandidate.CrossSection
    Set oTargCS = oTarget.CrossSection
    
    Dim oCandCSDefiniton As IJCrossSection
    Dim oTargCSDefiniton As IJCrossSection
    
    If Not oCandCS Is Nothing And Not oTargCS Is Nothing Then
        Set oCandCSDefiniton = oCandCS.Definition
        Set oTargCSDefiniton = oTargCS.Definition
    End If
    
    For j = LBound(strKeywordChecksColl) To UBound(strKeywordChecksColl)
            
        bAPISuccess = False
        Select Case strKeywordChecksColl(j)

            Case "Material"
                bAPISuccess = oCPSCommonHelper.GetRelatedObjects(CP_Material, oCndtMaterial, oTgtMaterial)

                If bAPISuccess Then
                    vCndtPropValue = oCndtMaterial.MaterialType + " | " + oCndtMaterial.MaterialGrade
                    vTgtPropValue = oTgtMaterial.MaterialType + " | " + oTgtMaterial.MaterialGrade

                    If oCndtMaterial.MaterialType <> oTgtMaterial.MaterialType Or _
                        oCndtMaterial.MaterialGrade <> oTgtMaterial.MaterialGrade Then
                        bCommon = False
                    End If
                End If

            'Check cross-section Definition
            Case "Section"
                If Not oCandCSDefiniton Is Nothing And Not oTargCSDefiniton Is Nothing Then
                    vCndtPropValue = oCandCSDefiniton.Type
                    vTgtPropValue = oTargCSDefiniton.Type
                    
                    If Not (oCandCSDefiniton Is oTargCSDefiniton) Then
                        bCommon = False
                    End If
                End If
            
            Case "DefinitionAttributes"
                Dim oCandAttributesCol As IJDAttributesCol
                Dim oTargAttributesCol As IJDAttributesCol
                
                Set oCandAttributesCol = oCandCSDefiniton.Attributes
                Set oTargAttributesCol = oTargCSDefiniton.Attributes
                            
                bCommon = CompareAttributes(oCandAttributesCol, oTargAttributesCol)
            
            Case "SectionName"
                If Not oCandCS Is Nothing And Not oTargCS Is Nothing Then
                    
                    'variables for populating test command.
                    vCndtPropValue = oCandCS.SectionName
                    vTgtPropValue = oTargCS.SectionName
                    If StrComp(oCandCS.SectionName, oTargCS.SectionName, vbTextCompare) <> 0 Then
                        bCommon = False
                    End If
                End If
                
            Case "SectionStandard"
                If Not oCandCS Is Nothing And Not oTargCS Is Nothing Then
                    'variables for populating test command.
                    vCndtPropValue = oCandCS.SectionStandard
                    vTgtPropValue = oTargCS.SectionStandard
                    If StrComp(oCandCS.SectionStandard, oTargCS.SectionStandard, vbTextCompare) <> 0 Then
                        bCommon = False
                    End If
                End If
                
            'CutLength comparison
            Case "CutLength"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("CutLength", vCndtPropValue, vTgtPropValue, "ISPSMemberPartPrismatic")

                If bAPISuccess Then
                    If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dDistanceTolerance Then
                        bCommon = False
                    End If
                End If

            Case "TypeCategory"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("TypeCategory", vCndtPropValue, vTgtPropValue, "ISPSMemberType")

                If bAPISuccess Then
                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            Case "Type"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("Type", vCndtPropValue, vTgtPropValue, "ISPSMemberType")

                If bAPISuccess Then
                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If
             
            Case "Priority"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("Priority", vCndtPropValue, vTgtPropValue, "ISPSMemberType")

                If bAPISuccess Then
                    If vCndtPropValue <> vTgtPropValue Then
                        bCommon = False
                    End If
                End If
              
             'Length comparison
            Case "Length"
                Dim oMemberSystem As SPSMembers.ISPSMemberSystem
                Dim oCurve As IJCurve
     
                Set oMemberSystem = oCandidate.MemberSystem
                Set oCurve = oMemberSystem
                
                vCndtPropValue = oCurve.length
                
                Set oMemberSystem = oTarget.MemberSystem
                Set oCurve = oMemberSystem
                
                vTgtPropValue = oCurve.length
                

                If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dDistanceTolerance Then
                    bCommon = False
                End If
            
            'DryWeight comparison
            Case "DryWeight"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("DryWeight", vCndtPropValue, vTgtPropValue, "IJWeightCG")
                              
                If bAPISuccess Then
                    If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dWeightTolerance Then
                        bCommon = False
                    End If
                End If
                
            'WetWeight comparison
            Case "WetWeight"
                bAPISuccess = oCPSCommonHelper.GetDirectPropertyValues("WetWeight", vCndtPropValue, vTgtPropValue, "IJWeightCG")

                If bAPISuccess Then
                        If Abs(CDbl(vCndtPropValue) - CDbl(vTgtPropValue)) > m_dWeightTolerance Then
                            bCommon = False
                        End If
                End If
            
            Case "StageCode"
                bAPISuccess = oCPSCommonHelper.GetRelatedObjectPropertyValues(CP_ProductionRouting, "StageCode", vCndtPropValue, vTgtPropValue) ', "IJPlnProductionRouting")

                If bAPISuccess Then
                    If StrComp(CStr(vCndtPropValue), CStr(vTgtPropValue), vbTextCompare) <> 0 Then
                        bCommon = False
                    End If
                End If

            Case "WorkCenter"
                bAPISuccess = oCPSCommonHelper.GetRelatedObjectPropertyValues(CP_ProductionRouting, "WorkCenter", vCndtPropValue, vTgtPropValue) ', "IJPlnProductionRouting")

                If bAPISuccess Then
                    If StrComp(CStr(vCndtPropValue), CStr(vTgtPropValue), vbTextCompare) <> 0 Then
                        bCommon = False
                    End If
                End If

            Case "Actions"
                bAPISuccess = oCPSCommonHelper.GetRelatedObjects(CP_ProductionRouting, oCndtProdRouting, oTgtProdRouting)

                If bAPISuccess Then 'And Not oCndtProdRouting Is Nothing Then ' And Not oTgtProdRouting Is Nothing Then
                    If Not oCndtProdRouting Is Nothing Then
                        Set oCndtProdRtgActions = GetRelatedObjects(oCndtProdRouting, IID_IJPlnProductionRouting, "RoutingAction")
                        vCndtPropValue = oCndtProdRtgActions.Count
                    Else
                        vCndtPropValue = ""
                    End If
                    
                    If Not oTgtProdRouting Is Nothing Then
                        Set oTgtProdRtgActions = GetRelatedObjects(oTgtProdRouting, IID_IJPlnProductionRouting, "RoutingAction")
                        vTgtPropValue = oTgtProdRtgActions.Count
                    Else
                        vTgtPropValue = ""
                    End If
                    
                    If CStr(vCndtPropValue) <> CStr(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            Case "Action"   'Need more inputs on this
            
            Case "Geometry"
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                bAPISuccess = True
                bCommon = bIsGeomSame
                
            'Checks whether the 2 members are aligned parallelly or not.
            Case "MemberDirection"
                Dim objCandLCS As IJLocalCoordinateSystem
                Dim objTargLCS As IJLocalCoordinateSystem
                
                Set objCandLCS = oCandidate
                Set objTargLCS = oTarget
                
                'Compare the direction vectors.
                Dim otempVector As IJDVector
                
                Set otempVector = objCandLCS.XAxis.Cross(objTargLCS.XAxis)
                
                If Not otempVector Is Nothing Then
                    If otempVector.length = 0 Then
                    
                        Set otempVector = objCandLCS.YAxis.Cross(objTargLCS.YAxis)
                        
                        If Not otempVector Is Nothing Then
                            If otempVector.length = 0 Then
                        
                                Set otempVector = objCandLCS.ZAxis.Cross(objTargLCS.ZAxis)
                                
                                If Not otempVector Is Nothing Then
                                    If otempVector.length <> 0 Then
                                    
                                        bCommon = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

            Case "PCsCount"
                bAPISuccess = oCPSCommonHelper.GetConnectedObjects(AllPhysicalConnections, oCndtConnectedObjects, oTgtConnectedObjects)

                If bAPISuccess Then
                    vCndtPropValue = oCndtConnectedObjects.Count
                    vTgtPropValue = oTgtConnectedObjects.Count

                    If CLng(vCndtPropValue) <> CLng(vTgtPropValue) Then
                        bCommon = False
                    End If
                End If

            'Bevels
            Case "ConnectionBevels"
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    If bPCsAlreadyMatched = False Then
                        bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(AllPhysicalConnections, oMatchedPCs)
                    Else
                        bAPISuccess = True
                    End If
                        
                    Dim bFlip As Boolean
                    bFlip = oCPSCommonHelper.Flipped
                    
                    If bAPISuccess And oMatchedPCs.Count > 0 Then bCommon = CompareBevels(oCandidate, oTarget, oMatchedPCs, bFlip, strBevelInfo)
                    If bAPISuccess And oMatchedPCs.Count = 0 Then bCommon = True
                End If

            'MountingAngles
            Case "MountingAngles"
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    If bPCsAlreadyMatched = False Then
                        bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(AllPhysicalConnections, oMatchedPCs)
                    Else
                        bAPISuccess = True
                    End If

                    If bAPISuccess Then
                        bPCsAlreadyMatched = True
                        bCommon = ArePCMountingAnglesCommon(oMatchedPCs, oCPSCommonHelper.TransMatrix)
                    End If
                End If
            
            Case "MarkingLines"
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(MarkingLines, oConnectedObjPairs)
    
                    If bAPISuccess Then
                        bCommon = AreConnectedMfgObjectsCommon(oConnectedObjPairs, oCPSCommonHelper.TransMatrix)
                    End If
                End If

            Case "Margins"
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(Margins, oConnectedObjPairs)

                    If bAPISuccess Then
                        bCommon = AreConnectedMfgObjectsCommon(oConnectedObjPairs, oCPSCommonHelper.TransMatrix)
                    End If
                 End If

            Case "Shrinkages"
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(Shrinkages, oConnectedObjPairs)
    
                    If bAPISuccess Then
                        bCommon = AreConnectedMfgObjectsCommon(oConnectedObjPairs, oCPSCommonHelper.TransMatrix)
                    End If
                End If

            Case "AssyShrinkages"
                If m_bGeometriesCompared = False Then
                    bIsGeomSame = ArePartGeometriesEqual(oCPSCommonHelper)
                End If
                
                If bIsGeomSame Then
                    bAPISuccess = oCPSCommonHelper.GetSortedConnectedObjects(AssemblyShrinkages, oConnectedObjPairs)
    
                    If bAPISuccess Then
                        bCommon = AreConnectedMfgObjectsCommon(oConnectedObjPairs, oCPSCommonHelper.TransMatrix)
                    End If
                End If
        End Select
        
        'Populate test command outputs
        If bCalledByTestCmd Then
            Dim arr() As String
            arr = Split(GetPropertyValues(oCandidate, vCndtPropValue, vTgtPropValue, bAPISuccess, bCommon, strKeywordChecksColl(j), oCPSCommonHelper, bIsGeomSame, oConnectedObjPairs, oMatchedPCs), "@")
            If strKeywordChecksColl(j) = "ConnectionBevels" Then
                saCandCheckValues(j) = strBevelInfo
                saTgtCheckValues(j) = strBevelInfo
            Else
                saCandCheckValues(j) = arr(0)
                saTgtCheckValues(j) = arr(1)
            End If
            saStatus(j) = CStr(arr(2))
        Else
            If bCommon = False Then
                Dim oCandidateNamedItem As IJNamedItem
                Set oCandidateNamedItem = pCandidate
            
                Dim oTargetNamedItem As IJNamedItem
                Set oTargetNamedItem = pTarget
                    
                If m_bGeometriesCompared = True And bIsGeomSame = False Then
                    ' It indicates that the geometry is not same
                    strResultMessage = "Geometry of Candidate part: " & oCandidateNamedItem.Name & " and Target Part: " & oTargetNamedItem.Name & " are not same"
                Else
                    strResultMessage = strDisplayChecksColl(j) + " of Candidate part: " & oCandidateNamedItem.Name & " and Target Part: " & oTargetNamedItem.Name & " are not same"
                End If
                
                Set oCandidateNamedItem = Nothing
                Set oTargetNamedItem = Nothing
                    
                Exit For
            End If
        End If
        
        vCndtPropValue = Null
        vTgtPropValue = Null
        bCommon = True
    Next j
    
WrapUp:
    Set oCandidate = Nothing
    Set oTarget = Nothing
    
    Set oCandCS = Nothing
    Set oTargCS = Nothing
    Set oCandCSDefiniton = Nothing
    Set oTargCSDefiniton = Nothing
    
    Set oCPSCommonHelper = Nothing
    Set oCPSMemberHelper = Nothing

    IJCommonEntityComparison_AreCommon = bCommon
    
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

Private Function IJStandardEntityComparison_IsStandardEntity(ByVal oCommonPartMember As Object, strDisplayChecksColl() As String, strKeywordChecksColl() As String, _
                                                                ByVal strStandardComparisonType As String, ByVal strStandardReferencePath As String, _
                                                                strStandardReferencePartName As String, strProcessPurpose As String) As Boolean
Const METHOD = "IJStandardEntityComparison_IsStandardEntity"
On Error GoTo ErrorHandler

    IJStandardEntityComparison_IsStandardEntity = False
        
    If strStandardComparisonType = "Model" Then
        ' get all the standard reference part collection from the input reference path
        Dim oStandardReferenceParts As IJDObjectCollection
        Set oStandardReferenceParts = GetStandardReferencePartsCollection(strStandardReferencePath)
        
        If oStandardReferenceParts Is Nothing Then
            Exit Function
        Else
            If oStandardReferenceParts.Count = 0 Then
                Exit Function
            End If
        End If
    
        Dim oStandardRefObj As Object
        
        ' Call AreCommon API for each standard reference part
        For Each oStandardRefObj In oStandardReferenceParts
            Dim strResultMsg  As String
            Dim saCandCheckValues() As Variant
            Dim saTgtCheckValues() As Variant
            Dim saStatus() As Variant
            Dim bCommon As Boolean
                        
            ' In case of Standard part rule execution, as Manager is not avaialable,
            ' needs to skip execution of AreCommon on first comparison failure so use m_bStandardPartRule as True
            m_bStandardPartRule = True
            bCommon = IJCommonEntityComparison_AreCommon(oCommonPartMember, oStandardRefObj, strDisplayChecksColl, strKeywordChecksColl, Nothing, _
                                                          strResultMsg, saCandCheckValues, saTgtCheckValues, saStatus)
                                                        
            ' If Common reference part is found, return refernce part name and TRUE to the caller
            If bCommon = True Then
                ' Return the standard part reference name
                Dim oStdPartNamedItem As IJNamedItem
                Set oStdPartNamedItem = oStandardRefObj
                strStandardReferencePartName = oStdPartNamedItem.Name
                
                ' ProcessPurpose codelist short description values => Common, StandardByModel, StandardByLookup
                strProcessPurpose = "StandardByModel"
                IJStandardEntityComparison_IsStandardEntity = True
                
                Exit For
            End If
        Next
    End If
    
    m_bStandardPartRule = False

Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

Private Function ArePartGeometriesEqual(oCPSCommonHelper As IJPlnCompareHelperEx) As Boolean
Const METHOD = "ArePartGeometriesEqual"
On Error GoTo ErrorHandler

    m_bGeometriesCompared = True
    ArePartGeometriesEqual = False
       
    ' Setting the Property values to the helper
    ' Setting the Property values to the helper
    ' Note that Common Part Service for Members needs only Distance Tolerance,
    ' Area Tolerance and Volume tolerance for geometry comparison
    oCPSCommonHelper.Properties(DistanceTolerance) = m_dDistanceTolerance
    oCPSCommonHelper.Properties(AreaTolerance) = m_dAreaTolerance
    oCPSCommonHelper.Properties(VolumeTolerance) = m_dVolumeTolerance
        
    ArePartGeometriesEqual = oCPSCommonHelper.IsGeometrySame(m_dDistanceTolerance)
    
Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

