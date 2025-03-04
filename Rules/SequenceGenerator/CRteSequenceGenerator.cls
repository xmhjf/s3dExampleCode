VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRteSequenceGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJSequenceGenerator


'If IJSequenceGenerator_ComparePositions returns
'    strId1 is less than strId2  ==>  -1
'    strId1 is equal to strId2  ==> 0
'    strId1 is greater than strId2  ==> 1

Private Function IJSequenceGenerator_ComparePositions(ByVal strId1 As String, ByVal strId2 As String) As Long

    Dim lId1 As Long
    Dim lId2 As Long
    
    If (Len(strId1) = 0) Then
        If (Len(strId2) = 0) Then
            IJSequenceGenerator_ComparePositions = 0
        Else
            IJSequenceGenerator_ComparePositions = -1
        End If
        Exit Function
    End If
    If (Len(strId2) = 0) Then
        IJSequenceGenerator_ComparePositions = 1
        Exit Function
    End If
    
    lId1 = CInt(strId1)
    lId2 = CInt(strId2)
    
    If lId1 > lId2 Then
        IJSequenceGenerator_ComparePositions = 1
    ElseIf lId1 = lId2 Then
        IJSequenceGenerator_ComparePositions = 0
    Else
        IJSequenceGenerator_ComparePositions = -1
    End If
    
    
End Function

Private Sub IJSequenceGenerator_GetNextSequenceId(ByVal eObjectType As GSCADRouteUtils.eSequenceMemberObjectTypes, strId As String)

    If (Len(strId) = 0) Then
        strId = "1"
    Else
        strId = CInt(strId) + 1
    End If
End Sub
