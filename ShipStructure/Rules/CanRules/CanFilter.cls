VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CanFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const MODULE = "CanFilter"

Public Function IsTube(element As Object) As Integer
    
    Const METHOD = "IsTube"
    On Error GoTo ErrorHandler
    Dim oXS As ISPSCrossSection
    
    IsTube = 0
    
    If Not element Is Nothing Then
        
        If TypeOf element Is ISPSFrameConnection Then
            Dim oFC As ISPSFrameConnection

            Set oFC = element
            Set oXS = oFC.MemberSystem.MemberPartAtEnd(oFC.WPO.portIndex).CrossSection
            If Not oXS Is Nothing Then
                ' boolean true = -1  so IsTube = oXS.CrossSectionTypeName = "CircTube"
                ' will not work
                If oXS.CrossSectionTypeName = "CircTube" Then
                    IsTube = 1
                End If
            End If

        ElseIf TypeOf element Is ISPSMemberSystemLinear Then
            Dim iSPSMemberSystem As iSPSMemberSystem
            
            Set iSPSMemberSystem = element
            IsTube = ContainsTube(iSPSMemberSystem.DesignParts)
        
        ElseIf TypeOf element Is ISPSAxisEndPort Then
            Dim oIJPort As IJPort
            Dim oMemberPart As ISPSMemberPartCommon
            Dim strCrossSectionTypeName As String
            
            Set oIJPort = element
            Set oMemberPart = oIJPort.Connectable
            Set oXS = oMemberPart.CrossSection
            strCrossSectionTypeName = oXS.CrossSectionTypeName
            If strCrossSectionTypeName = "BUCan" Or strCrossSectionTypeName = "CircTube" Then
                IsTube = 1
            End If
        
        ElseIf TypeOf element Is ISPSSplitMemberConnection Then
            Dim iSplitConn As ISPSSplitMemberConnection
            Dim iCanRule As ISPSCanRule
            Dim strIJGeometry As String, strColl As String
            
            strIJGeometry = "{96eb9676-6530-11d1-977f-080036754203}"
            strColl = "ConstructionForOutput"

            ' cannot accept a split connection that was created by a CanRule.
            Set iCanRule = GetTargetObject(element, strIJGeometry, strColl)
            If iCanRule Is Nothing Then
                Set iSplitConn = element
                IsTube = ContainsTube(iSplitConn.PartPorts)
            End If
        
        End If
    End If

    Exit Function

ErrorHandler:
    IsTube = 0
    WriteToErrorLog Err.Number, MODULE, METHOD, Err.Description
    Err.Clear
End Function


Private Function ContainsTube(elesPartsOrPorts As IJElements) As Integer

    Const METHOD = "ContainsTube"
    On Error GoTo ErrorHandler

    ContainsTube = 0

    Dim oObj As Object
    Dim oMemberPart As ISPSMemberPartCommon
    Dim oXS As ISPSCrossSection
    Dim lIndex As Long, lCount As Long
    Dim oPort As IJPort
    
    lCount = elesPartsOrPorts.count
    
    For lIndex = 1 To lCount
        Set oObj = elesPartsOrPorts.Item(lIndex)
        If TypeOf oObj Is ISPSMemberPartCommon Then
            Set oMemberPart = oObj
        Else
            Set oPort = oObj
            Set oMemberPart = oPort.Connectable
        End If
        Set oXS = oMemberPart.CrossSection
        If Not oXS Is Nothing Then
            ' boolean true = -1  so IsTube = oXS.CrossSectionTypeName = "CircTube"
            ' will not work
            If oXS.CrossSectionTypeName = "CircTube" Then
                ContainsTube = 1
                Set oXS = Nothing
                Exit Function
            End If
            Set oXS = Nothing
        End If
    Next lIndex

    Exit Function
ErrorHandler:
    ContainsTube = 0
    WriteToErrorLog Err.Number, MODULE, METHOD, Err.Description
    Err.Clear
End Function

Private Function GetTargetObject(oSource As Object, strInterface As String, strCollectionName As String) As Object
    Const METHOD = "GetTargetObject"
    On Error GoTo ErrorHandler
  
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Set pRelationHelper = oSource
    On Error Resume Next
    Set pCollectionHelper = pRelationHelper.CollectionRelations(strInterface, strCollectionName)
    On Error Resume Next
    If Not pCollectionHelper Is Nothing Then
        If pCollectionHelper.count > 0 Then
            Set GetTargetObject = pCollectionHelper.Item(1)
        End If
    End If

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
End Function

