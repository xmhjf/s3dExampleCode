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

Private Const Attr_RefSideFirstBevelAngle As String = "RefSideFirstBevelAngle"
Private Const Attr_AntiRefSideFirstBevelAngle As String = "AntiRefSideFirstBevelAngle"
Private Const Attr_RefSideFirstBevelMethod As String = "RefSideFirstBevelMethod"
Private Const Attr_AntiRefSideFirstBevelMethod As String = "AntiRefSideFirstBevelMethod"
Private Const Attr_RefSideFirstBevelDepth As String = "RefSideFirstBevelDepth"
Private Const Attr_AntiRefSideFirstBevelDepth As String = "AntiRefSideFirstBevelDepth"
Private Const Attr_RefSideSecondBevelAngle As String = "RefSideSecondBevelAngle"
Private Const Attr_AntiRefSideSecondBevelAngle As String = "AntiRefSideSecondBevelAngle"
Private Const Attr_RefSideSecondBevelMethod As String = "RefSideSecondBevelMethod"
Private Const Attr_AntiRefSideSecondBevelMethod As String = "AntiRefSideSecondBevelMethod"
Private Const Attr_RefSideSecondBevelDepth As String = "RefSideSecondBevelDepth"
Private Const Attr_AntiRefSideSecondBevelDepth As String = "AntiRefSideSecondBevelDepth"

Private Const Attr_NRRefSideFirstBevelAngle As String = "NRRefSideFirstBevelAngle"
Private Const Attr_NRAntiRefSideFirstBevelAngle As String = "NRAntiRefSideFirstBevelAngle"
Private Const Attr_NRRefSideFirstBevelMethod As String = "NRRefSideFirstBevelMethod"
Private Const Attr_NRAntiRefSideFirstBevelMethod As String = "NRAntiRefSideFirstBevelMethod"
Private Const Attr_NRRefSideFirstBevelDepth As String = "NRRefSideFirstBevelDepth"
Private Const Attr_NRAntiRefSideFirstBevelDepth As String = "NRAntiRefSideFirstBevelDepth"
Private Const Attr_NRRefSideSecondBevelAngle As String = "NRRefSideSecondBevelAngle"
Private Const Attr_NRAntiRefSideSecondBevelAngle As String = "NRAntiRefSideSecondBevelAngle"
Private Const Attr_NRRefSideSecondBevelMethod As String = "NRRefSideSecondBevelMethod"
Private Const Attr_NRAntiRefSideSecondBevelMethod As String = "NRAntiRefSideSecondBevelMethod"
Private Const Attr_NRRefSideSecondBevelDepth As String = "NRRefSideSecondBevelDepth"
Private Const Attr_NRAntiRefSideSecondBevelDepth As String = "NRAntiRefSideSecondBevelDepth"

Private Const Attr_ReferenceSide As String = "ReferenceSide"
Private Const Attr_NRReferenceSide As String = "NRReferenceSide"

Private Sub IJCbFAttributeCopyRule_GetCustomizedAttributeInfo( _
        ByVal customizedAttributesInfo As RuleManagement.customizedAttributesInfo)
    
    With customizedAttributesInfo
        .Add Attr_RefSideFirstBevelAngle
        .Add Attr_AntiRefSideFirstBevelAngle
        
        .Add Attr_RefSideFirstBevelMethod
        .Add Attr_AntiRefSideFirstBevelMethod
        
        .Add Attr_RefSideFirstBevelDepth
        .Add Attr_AntiRefSideFirstBevelDepth
        
        .Add Attr_RefSideSecondBevelAngle
        .Add Attr_AntiRefSideSecondBevelAngle
        
        .Add Attr_RefSideSecondBevelMethod
        .Add Attr_AntiRefSideSecondBevelMethod
        
        .Add Attr_RefSideSecondBevelDepth
        .Add Attr_AntiRefSideSecondBevelDepth
        
        .Add Attr_NRRefSideFirstBevelAngle
        .Add Attr_NRAntiRefSideFirstBevelAngle
        
        .Add Attr_NRRefSideFirstBevelMethod
        .Add Attr_NRAntiRefSideFirstBevelMethod
        
        .Add Attr_NRRefSideFirstBevelDepth
        .Add Attr_NRAntiRefSideFirstBevelDepth
        
        .Add Attr_NRRefSideSecondBevelAngle
        .Add Attr_NRAntiRefSideSecondBevelAngle
        
        .Add Attr_NRRefSideSecondBevelMethod
        .Add Attr_NRAntiRefSideSecondBevelMethod
        
        .Add Attr_NRRefSideSecondBevelDepth
        .Add Attr_NRAntiRefSideSecondBevelDepth
    End With
    
End Sub

Public Property Get RefSideFirstBevelAngle( _
                     ByVal interfaceName As String, _
                     ByVal sourceAttrList As IJAttributeLookup, _
                     ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set RefSideFirstBevelAngle = UpdateTargetReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_AntiRefSideFirstBevelAngle)
    
End Property

Public Property Get AntiRefSideFirstBevelAngle( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

   Set AntiRefSideFirstBevelAngle = UpdateTargetReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_RefSideFirstBevelAngle)
        
End Property

Public Property Get RefSideFirstBevelMethod( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set RefSideFirstBevelMethod = UpdateTargetReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_AntiRefSideFirstBevelMethod)
    
End Property

Public Property Get AntiRefSideFirstBevelMethod( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set AntiRefSideFirstBevelMethod = UpdateTargetReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_RefSideFirstBevelMethod)
   
End Property

Public Property Get RefSideFirstBevelDepth( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set RefSideFirstBevelDepth = UpdateTargetReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_AntiRefSideFirstBevelDepth)
    
End Property

Public Property Get AntiRefSideFirstBevelDepth( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set AntiRefSideFirstBevelDepth = UpdateTargetReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_RefSideFirstBevelDepth)
    
End Property

Public Property Get RefSideSecondBevelAngle( _
                     ByVal interfaceName As String, _
                     ByVal sourceAttrList As IJAttributeLookup, _
                     ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set RefSideSecondBevelAngle = UpdateTargetReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_AntiRefSideSecondBevelAngle)
    
End Property

Public Property Get AntiRefSideSecondBevelAngle( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set AntiRefSideSecondBevelAngle = UpdateTargetReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_RefSideSecondBevelAngle)
        
End Property

Public Property Get RefSideSecondBevelMethod( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set RefSideSecondBevelMethod = UpdateTargetReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_AntiRefSideSecondBevelMethod)
    
End Property

Public Property Get AntiRefSideSecondBevelMethod( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set AntiRefSideSecondBevelMethod = UpdateTargetReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_RefSideSecondBevelMethod)
   
End Property

Public Property Get RefSideSecondBevelDepth( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set RefSideSecondBevelDepth = UpdateTargetReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_AntiRefSideSecondBevelDepth)
    
End Property

Public Property Get AntiRefSideSecondBevelDepth( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set AntiRefSideSecondBevelDepth = UpdateTargetReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_RefSideSecondBevelDepth)
    
End Property

Public Property Get NRRefSideFirstBevelAngle( _
                     ByVal interfaceName As String, _
                     ByVal sourceAttrList As IJAttributeLookup, _
                     ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set NRRefSideFirstBevelAngle = UpdateTargetNonReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_NRAntiRefSideFirstBevelAngle)
    
End Property

Public Property Get NRAntiRefSideFirstBevelAngle( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

   Set NRAntiRefSideFirstBevelAngle = UpdateTargetNonReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_NRRefSideFirstBevelAngle)
        
End Property

Public Property Get NRRefSideFirstBevelMethod( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set NRRefSideFirstBevelMethod = UpdateTargetNonReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_NRAntiRefSideFirstBevelMethod)
    
End Property

Public Property Get NRAntiRefSideFirstBevelMethod( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set NRAntiRefSideFirstBevelMethod = UpdateTargetNonReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_NRRefSideFirstBevelMethod)
   
End Property

Public Property Get NRRefSideFirstBevelDepth( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set NRRefSideFirstBevelDepth = UpdateTargetNonReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_NRAntiRefSideFirstBevelDepth)
    
End Property

Public Property Get NRAntiRefSideFirstBevelDepth( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set NRAntiRefSideFirstBevelDepth = UpdateTargetNonReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_NRRefSideFirstBevelDepth)

End Property

Public Property Get NRRefSideSecondBevelAngle( _
                     ByVal interfaceName As String, _
                     ByVal sourceAttrList As IJAttributeLookup, _
                     ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set NRRefSideSecondBevelAngle = UpdateTargetNonReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_NRAntiRefSideSecondBevelAngle)
    
End Property

Public Property Get NRAntiRefSideSecondBevelAngle( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set NRAntiRefSideSecondBevelAngle = UpdateTargetNonReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_NRRefSideSecondBevelAngle)
        
End Property

Public Property Get NRRefSideSecondBevelMethod( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set NRRefSideSecondBevelMethod = UpdateTargetNonReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_NRAntiRefSideSecondBevelMethod)
    
End Property

Public Property Get NRAntiRefSideSecondBevelMethod( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set NRAntiRefSideSecondBevelMethod = UpdateTargetNonReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_NRRefSideSecondBevelMethod)
   
End Property

Public Property Get NRRefSideSecondBevelDepth( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set NRRefSideSecondBevelDepth = UpdateTargetNonReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_NRAntiRefSideSecondBevelDepth)
    
End Property

Public Property Get NRAntiRefSideSecondBevelDepth( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As AttributeRuleResult

    Set NRAntiRefSideSecondBevelDepth = UpdateTargetNonReferencePartBevelParameter(interfaceName, sourceAttrList, targetAttrList, Attr_NRRefSideSecondBevelDepth)
    
End Property

Private Function AreReferenceSidesDifferent( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As Boolean
                     
    Dim sourceRefSide As IJDAttribute
    Set sourceRefSide = sourceAttrList.Lookup(interfaceName, Attr_ReferenceSide)
    
    Dim targetRefSide As IJDAttribute
    Set targetRefSide = targetAttrList.Lookup(interfaceName, Attr_ReferenceSide)
    
    If (Not sourceRefSide Is Nothing) And (Not targetRefSide Is Nothing) Then
        AreReferenceSidesDifferent = (sourceRefSide.Value <> targetRefSide.Value)
    End If
End Function

Private Function AreNRReferenceSidesDifferent( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup) As Boolean
                     
    Dim sourceNRRefSide As IJDAttribute
    Set sourceNRRefSide = sourceAttrList.Lookup(interfaceName, Attr_NRReferenceSide)
    
    Dim targetNRRefSide As IJDAttribute
    Set targetNRRefSide = targetAttrList.Lookup(interfaceName, Attr_NRReferenceSide)
    
    If (Not sourceNRRefSide Is Nothing) And (Not targetNRRefSide Is Nothing) Then
        AreNRReferenceSidesDifferent = (sourceNRRefSide.Value <> targetNRRefSide.Value)
    End If
End Function

Private Function UpdateTargetReferencePartBevelParameter( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup, ByVal sourceParameterName As String) As AttributeRuleResult
                    
    Set UpdateTargetReferencePartBevelParameter = New AttributeRuleResult
    
    If AreReferenceSidesDifferent(interfaceName, sourceAttrList, targetAttrList) Then
        Dim sourceAntiRef As IJDAttribute
        Set sourceAntiRef = sourceAttrList.Lookup(interfaceName, sourceParameterName)
        
        If Not sourceAntiRef Is Nothing Then
           UpdateTargetReferencePartBevelParameter.UpdateTargetValue sourceAntiRef.Value
        End If
    End If
End Function

Private Function UpdateTargetNonReferencePartBevelParameter( _
                    ByVal interfaceName As String, _
                    ByVal sourceAttrList As IJAttributeLookup, _
                    ByVal targetAttrList As IJAttributeLookup, ByVal sourceParameterName As String) As AttributeRuleResult
                    
    Set UpdateTargetNonReferencePartBevelParameter = New AttributeRuleResult
    
    If AreNRReferenceSidesDifferent(interfaceName, sourceAttrList, targetAttrList) Then
        Dim sourceAntiRef As IJDAttribute
        Set sourceAntiRef = sourceAttrList.Lookup(interfaceName, sourceParameterName)
        
        If Not sourceAntiRef Is Nothing Then
            UpdateTargetNonReferencePartBevelParameter.UpdateTargetValue sourceAntiRef.Value
        End If
    End If
End Function

