VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepairMargin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 2009 Intergraph Corporation.  All Rights Reserved.
'
' File: RepairMargin.cls
'
' Author: J.Manasa
'
' Abstract: Check Manufacturability repair tool that deletes one of the margins in case of concurrency
'
'********************************************************************

Option Explicit

Private Const Module = "StrMfgCheckMfctyMargin.RepairMargin"
Private Const strProgID = "StrMfgCheckMfctyMargin.RepairMargin"

Private Const IID_IJDObject As String = "{05F65FD1-57B3-11D1-93FC-0060973D4777}"

Implements IJCheckMfctyRepair


Private Sub Class_Initialize()
    Const METHOD As String = "Class_Initialize"
    On Error GoTo ErrorHandler
    GoTo Cleanup

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD).Number

Cleanup:

End Sub

'******************************************************************************
' Routine: IJCheckMfctyRepair_Repair
'
' Abstract:
'   Implementation of the IJCheckMfcty interface
'
' Description: Perform a repair on the specified object.
'
' Inputs:
'     vObject:  The object to be repaired.
'******************************************************************************
Public Sub IJCheckMfctyRepair_Repair(ByRef vObject As Variant)
Const METHOD As String = "IJCheckMfctyRepair_Repair"
   On Error GoTo ErrorHandler
    
    Dim sMessage As String
    sMessage = "...Performing Repair"
    RepairObject vObject
    
   Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, sMessage).Number

Cleanup:

End Sub

Private Sub RepairObject(ByRef vObject As Variant)
    Dim oUnk As IUnknown
    Dim sMessage As String
    Dim eObject As Object
    
    Dim oPort As IJPort
    Dim oEligiblePort As IJPort
    
    Dim oStructConnectable  As IJStructConnectable
    Dim oMfgDefCol As Collection
    Dim oEnumPorts As IMSCoreCollections.IJElements
    Dim m_oEligiblePortColl As IJDObjectCollection
    Dim oFabMarginColl As Collection
    Dim oFabMargin As IJDObject
    Dim oPart As IJDObject
    Dim oObject As Object
    Dim oIJDObject As IJDObject
    Dim oIJDOtherObject As IJDObject
    
    Dim oMargin As IJDFabMargin
    Dim oOtherMargin As IJDFabMargin
    
    Dim i As Integer
    Dim n As Integer
    Dim j As Integer
    
        
    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper
    
    Set oEnumPorts = New JObjectCollection 
    
    'Check if current object is of type IJPlatePart or IJProfilePart
         If TypeOf vObject Is IJPlatePart Or TypeOf vObject Is IJProfilePart Then
            Set oStructConnectable = vObject
            'Get all the ports
            sMessage = "...Getting all the ports"
            oEnumPorts.AddElements oStructConnectable.enumAllPorts
            Set oStructConnectable = Nothing
         ElseIf TypeOf vObject Is ISPSMemberPartPrismatic Then
            Dim oPartSupport As IJPartSupport
            Set oPartSupport = New MemberPartSupport
            Set oPartSupport.Part = vObject
            Set oStructConnectable = oPartSupport
            oEnumPorts.AddElements oStructConnectable.enumAllPorts
            Set oStructConnectable = Nothing
            Set oPartSupport = Nothing
        Else
            Exit Sub
        End If
        
         Set m_oEligiblePortColl = New JObjectCollection
        
         'Get the eligible port collection
         sMessage = "...Getting the eligible port collection"

          For j = 1 To oEnumPorts.Count
            Set oEligiblePort = oEnumPorts.Item(j)
            
            
            If TypeOf vObject Is IJPlatePart Then
            
                If IsLateralFacePort(oEligiblePort) Then
                    m_oEligiblePortColl.Add oEligiblePort
                    Set oEligiblePort = Nothing
                End If
                
            ElseIf TypeOf vObject Is IJStructProfilePart Then
                
                If IsBaseOrOffsetFacePort(oEligiblePort) Then
                    m_oEligiblePortColl.Add oEligiblePort
                    Set oEligiblePort = Nothing
                
                End If
                
            End If
            
        Next j
        
         sMessage = "...Iterating through the eligible port collection"
        Dim oMfgUpdateService  As IJMfgUpdateService
        Set oMfgUpdateService = New MfgUpdateService
        
        
        Set oFabMarginColl = New Collection
        'Iterate through all the ports
        For Each eObject In m_oEligiblePortColl
                
                Set oPort = eObject
                
                'Get all the Mfgdefinitions on that port
                Set oMfgDefCol = oMfgRuleHelper.GetMfgDefinitions(oPort)
                
                For i = 1 To oMfgDefCol.Count
                     If (TypeOf oMfgDefCol.Item(i) Is IJDFabMargin) Then
                        Set oMargin = oMfgDefCol.Item(i)
                        For j = 1 To oMfgDefCol.Count
                            Set oOtherMargin = oMfgDefCol.Item(j)
                            If Not oOtherMargin Is oMargin Then
                                If CInt(oMargin.MarginType) = CInt(oOtherMargin.MarginType) And CInt(oMargin.MarginMode) = CInt(oOtherMargin.MarginMode) Then
                                    oMfgUpdateService.CheckForPartialUpdate oMargin, oUnk
                                    If oUnk Is Nothing Then  'Margins are not as a result of partial split migration
                                         oFabMarginColl.Add oMargin
                                    End If
                                End If
                            End If
                        Next 'oMfgDefCol j
                    End If
                    
                Next 'oMfgDefCol i
                
        Next 'm_oEligiblePortColl
                
           
      'TO DELETE THE MARGIN WHICH IS IN THE TO-DO-LIST
      
'      If m_FabMarginColl.Count > 1 Then
'        For Each oObject In m_FabMarginColl
'
'            Set oFabMargin = oObject
'
'            Dim oIJDAssocRel As IJDAssocRelation
'            Dim oIJDTargetObjColl As IJDTargetObjectCol
'
'            'Check if the margin is in the To-Do-List
'            Set oIJDAssocRel = oFabMargin
'            Set oIJDTargetObjColl = oIJDAssocRel.CollectionRelations(IID_IJDObject, "toErrorList")
'
'            If Not oIJDTargetObjColl Is Nothing Then 'The Margin is in the To-Do-List
'                'Remove it
'                If oIJDTargetObjColl.Count > 0 Then
'                    sMessage = "...Deleting the Margin"
'                    On Error Resume Next 'If there are no permissions on the Mfg object
'                    oFabMargin.Remove
'                End If
'            End If
'        Next
'      End If
        
      sMessage = "...Iterating through the margins to get the newest margin"
      Set oIJDObject = oFabMarginColl.Item(1)
      For i = 2 To oFabMarginColl.Count
         Set oIJDOtherObject = oFabMarginColl.Item(i)
            If oIJDOtherObject.DateLastModified > oIJDObject.DateLastModified Then
                 Set oIJDObject = oIJDOtherObject
                 Set oIJDOtherObject = Nothing
            End If
      Next
    
      sMessage = "...Deleting other margins"
      For i = 1 To oFabMarginColl.Count
        Set oIJDOtherObject = oFabMarginColl.Item(i)
           If Not oIJDOtherObject Is oIJDObject Then
                On Error Resume Next 'If there are no permissions on the Mfg object
                oIJDOtherObject.Remove
           End If
      Next
        
      'The CheckMfcty command changes the status of the object.
    'After repair,we have to put it back to working status
        
      Set oPart = vObject
      oPart.ApprovalStatus = Working
End Sub
Public Function IsLateralFacePort(pObj As Object) As Long
    Const METHOD As String = "IsLateralFacePort"
    On Error GoTo ErrorHandler
   
    IsLateralFacePort = 0

    '- return if it not a face port
    If TypeOf pObj Is IJPort Then
        Dim oPort As IJPort
        Set oPort = pObj
        If oPort.Type <> PortFace Then
            Set oPort = Nothing
            Exit Function
        End If
        Set oPort = Nothing
    End If

    '- check for lateral port
    If Not pObj Is Nothing Then
        If TypeOf pObj Is IJStructPort Then
            Dim oStructPort As IJStructPort
            Set oStructPort = pObj
            If oStructPort.ContextID = CTX_LATERAL_LFACE Then 'Or oStructPort.ContextID = CTX_LATERAL_LFACE_3DSPLIT Then
                IsLateralFacePort = 1
            End If
            Set oStructPort = Nothing
        End If
    End If

    Exit Function
ErrorHandler:
    IsLateralFacePort = 0
End Function

Public Function IsBaseOrOffsetFacePort(pObj As Object) As Long
    Const METHOD As String = "IsBaseOrOffsetFacePort"
    On Error GoTo ErrorHandler
   
    Dim sMessage As String
    
    IsBaseOrOffsetFacePort = 0
         
    '- check for lateral port
    sMessage = "...Checking the ContextID"
    If Not pObj Is Nothing Then
        If TypeOf pObj Is IJStructPort Then
            Dim oStructPort As IJStructPort
            Set oStructPort = pObj
            If oStructPort.ContextID = CTX_BASE Or oStructPort.ContextID = CTX_OFFSET Then
                IsBaseOrOffsetFacePort = 1
            End If
            Set oStructPort = Nothing
        End If
    End If

    Exit Function
ErrorHandler:
    IsBaseOrOffsetFacePort = 0
    Err.Raise LogError(Err, Module, METHOD, sMessage).Number
End Function

