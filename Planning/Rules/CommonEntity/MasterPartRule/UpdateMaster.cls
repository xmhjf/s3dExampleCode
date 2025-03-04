VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpdateMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------
'      Copyright (C) 2012, Intergraph Corporation. All rights reserved.
'
'      Project: MasterPart

'      Author: Shylaja
'-----------------------------------------------------------------------------

Option Explicit

Implements IJCommonEntityUpdateService

Public m_oErrors As IJEditErrors   ' To collect and propagate the errors.
Public m_oError As IJEditError     ' Defined here for convenience
Private Const MODULE = "PlnMasterPartRule.UpdateMaster"
Private Const RELATION_DELETED = &H1000000

Private Sub IJCommonEntityUpdateService_CustomUpdate(ByVal eType As CommonPartInterfacesLib.CommonPartType, ByVal pGroup As Object)
Const METHOD = "IJCommonEntityUpdateService_CustomUpdate"
    On Error GoTo ErrorHandler

     Dim oParts      As IJElements
     Dim oGroup      As IJCommonPartGroup
     Dim oPrevMaster As Object
     Dim lCount      As Long
     Dim i           As Integer
     Dim oIJDObject  As IJDObject

    'MasterDelete is the case where the master part is deleted..
    'MemberDelete is the case where the member is deleted..

    Set oGroup = pGroup

    If (eType = Master) Then

        Dim lAssocFlags As Long
        Dim pStructAssocCompute As IJStructAssocCompute
        Set pStructAssocCompute = New StructAssocTools

        oGroup.GetCommonParts oParts
        lCount = oParts.Count   'lcount represents the number of common parts under that Common part group..

        'i value should be less than or eaqual to lcount..

        For i = 1 To lCount
            i = 2 ' Set the i value based on conditions..
            If (i <= lCount) Then

                pStructAssocCompute.GetAssocFlags oParts.Item(i), lAssocFlags
                'If the part is in Deleted parts list then dont set that part as master.

                If Not (lAssocFlags And RELATION_DELETED) = RELATION_DELETED Then
                    oGroup.Master = oParts.Item(i)

                    Exit For
                Else
                    Exit For
                End If

            End If
        Next

        Set oPrevMaster = Nothing

    ElseIf (eType = Member) Then

        Set oIJDObject = oGroup
        If (oGroup.ProcessControl = 2) Then

            'ProcessControl = 2 represents Type of group i.e Manual Group
            'If the group is manual group then do nothing.

        ElseIf (oGroup.ProcessControl = 1) Then

            'ProcessControl = 1 represents Type of group i.e Automatic(rule based) Group
            'If the group is rule based then delete the group..
            oIJDObject.Remove
            Set oIJDObject = Nothing
        End If
    End If

    Set oGroup = Nothing
    Exit Sub

ErrorHandler:
   Set m_oError = m_oErrors.AddFromErr(Err, MODULE & " - " & METHOD)
   m_oError.Raise
End Sub

'This Rule is triggered in 2 cases:
    '1. When we delete the Common part.
    '2. When we delete the master part.
Private Function IJCommonEntityUpdateService_Setting(ByVal eType As CommonPartInterfacesLib.CommonPartType, ByVal pGroupManager As Object) As CommonPartInterfacesLib.CommonPartUpdate
Const METHOD = "IJCommonEntityUpdateService_Setting"
On Error GoTo ErrorHandler

    'MasterDelete is the case where the master part is deleted..
    'MemberDelete is the case where the member is deleted..

    'Check the Deletion Type it is Master Type or Member Type.
    If (eType = Master) Then
         ' 3 enum values have exposed on IJCommonEntityUpdateService interface(DeleteGroup,RetainGroup,CustomAssign)..
        'if enum is "DeleteGroup" then delete the master part and as well as Group and then move all the remaining parts to relocated part folder..
        'if enum is "RetainGroup" ,only the master part will be deleted, Not Common Part Group. And next part in Common Part Group will be set as master part..
        'if enum is "CustomAssign", only the master part will be deleted, and user can set his own part as master on Common part group..

        IJCommonEntityUpdateService_Setting = RetainGroup
    End If

    'Check the Deletion Type it is Master Type or Member Type.
    If (eType = Member) Then
         ' 3 enum values have exposed on IJCommonEntityUpdateService interface(DeleteGroup,RetainGroup,CustomAssign)..
        'if enum is "DeleteGroup" then delete the group and move the part that is left in the group  to relocated part folder..
        'if enum is "RetainGroup" ,Group will not be deleted, Group will be exists with a single part..
        'if enum is "CustomAssign", User has to select the option from "CustomUpadate" method..

        IJCommonEntityUpdateService_Setting = DeleteGroup
    End If

   Exit Function

ErrorHandler:
   Set m_oError = m_oErrors.AddFromErr(Err, MODULE & " - " & METHOD)
   m_oError.Raise

End Function
