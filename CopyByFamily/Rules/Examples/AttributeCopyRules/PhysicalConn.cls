VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PhysicalConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJCbFAttributeCopyRule

Private Sub IJCbFAttributeCopyRule_GetCustomizedAttributeInfo( _
        ByVal customizedAttributesInfo As RuleManagement.customizedAttributesInfo)
    
    With customizedAttributesInfo
        .Add "RefSideFirstBevelAngle"
        .Add "AntiRefSideFirstBevelAngle"
    End With
    
End Sub

Public Property Get RefSideFirstBevelAngle( _
                     ByVal interfaceName As String, _
                     ByVal sourceAttrList As IJAttributeLookup, _
                     ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set RefSideFirstBevelAngle = New AttributeRuleResult
    
    If AreReferenceSidesDifferent(interfaceName, sourceAttrList, targetAttrList) Then
        Dim sourceAntiRef As IJDAttribute
        Set sourceAntiRef = sourceAttrList.Lookup(interfaceName, "AntiRefSideFirstBevelAngle")
        
        If Not sourceAntiRef Is Nothing Then
            RefSideFirstBevelAngle.UpdateTargetValue sourceAntiRef.Value
        End If
    End If
End Property

Public Property Get AntiRefSideFirstBevelAngle( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set AntiRefSideFirstBevelAngle = New AttributeRuleResult
    
    If AreReferenceSidesDifferent(interfaceName, sourceAttrList, targetAttrList) Then
        Dim sourceRef As IJDAttribute
        Set sourceRef = sourceAttrList.Lookup(interfaceName, "RefSideFirstBevelAngle")
        
        If Not sourceRef Is Nothing Then
            AntiRefSideFirstBevelAngle.UpdateTargetValue sourceRef.Value
        End If
    End If
End Property

Private Function AreReferenceSidesDifferent( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As Boolean
                     
    Dim sourceRefSide As IJDAttribute
    Set sourceRefSide = sourceAttrList.Lookup(interfaceName, "ReferenceSide")
    
    Dim targetRefSide As IJDAttribute
    Set targetRefSide = targetAttrList.Lookup(interfaceName, "ReferenceSide")
    
    If (Not sourceRefSide Is Nothing) And (Not targetRefSide Is Nothing) Then
        AreReferenceSidesDifferent = (sourceRefSide.Value <> targetRefSide.Value)
    End If
End Function

