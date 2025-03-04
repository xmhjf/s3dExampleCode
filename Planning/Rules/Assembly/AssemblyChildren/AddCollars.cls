VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddCollars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MODULE = "PlnAssyChildrenRules.AddCollars"

Implements IJAssemblyChildrenRules


Private Function IJAssemblyChildrenRules_GetSecondaryObjects(ByVal pChildren As IMSCoreCollections.IJElements) As IJElements
Const METHOD As String = "IJAssemblyChildrenRules_GetSecondaryObjects"
On Error GoTo ErrorHelper

    Dim oSystemChild    As IJSystemChild
    Dim oObject         As Object
    Dim oPlatePart      As IJPlatePart
    Dim oParentSys      As IJPlateSystem
    Dim oCollarPart     As IJCollarPart
    Dim oSDOCollar      As StructDetailObjects.Collar
    Dim oPenetrated     As IJPlatePart
    Dim oCollorsColl    As IJElements
    Dim oCollarParts    As IJDObjectCollection
    
    Set oCollorsColl = New JObjectCollection

    For Each oObject In pChildren
        If (TypeOf oObject Is IJPlatePart) And Not (TypeOf oObject Is IJCollarPart) Then

            Set oPlatePart = oObject

            Set oSystemChild = oPlatePart
            Set oParentSys = oSystemChild.GetParent

            Set oSystemChild = Nothing

            Set oSystemChild = oParentSys
            Set oParentSys = Nothing
            Set oParentSys = oSystemChild.GetParent

            Set oCollarParts = GetCollarParts(oParentSys)

           If Not oCollarParts Is Nothing Then
                For Each oCollarPart In oCollarParts
                    Set oSDOCollar = New StructDetailObjects.Collar
                  Set oSDOCollar.object = oCollarPart

                  Set oPenetrated = oSDOCollar.Penetrated

                  If oPenetrated Is oPlatePart Then
                      oCollorsColl.Add oCollarPart
                  End If
                  Set oCollarPart = Nothing
                  Set oPenetrated = Nothing
                  Set oSDOCollar = Nothing
                Next
          End If
        End If
        Set oObject = Nothing
        Set oPlatePart = Nothing
        Set oSystemChild = Nothing
    Next
    
    Set IJAssemblyChildrenRules_GetSecondaryObjects = oCollorsColl
    
    Set oSystemChild = Nothing
    Set oPlatePart = Nothing
    Set oParentSys = Nothing
    Set oCollarPart = Nothing
    Set oSDOCollar = Nothing
    Set oPenetrated = Nothing
    Set oCollorsColl = Nothing
    
Exit Function
ErrorHelper:
    MsgBox Err.Description
End Function

Public Function GetCollarParts(oObject As Object) As IJDObjectCollection
Const METHOD As String = "GetCollarParts"
On Error GoTo ErrorHelper
    
    Dim oSystem         As IJSystem
    Dim oCollection     As IJDTargetObjectCol
    Dim oCollarPart     As IJCollarPart
    Dim oReturnColl     As IJDObjectCollection
    Dim i               As Long
    
    Set oReturnColl = New JObjectCollection
    
    If TypeOf oObject Is IJCollarPart Then
        Set oCollarPart = oObject
        oReturnColl.Add oCollarPart
        
        Set GetCollarParts = oReturnColl
        
        Set oCollarPart = Nothing
        Exit Function
    End If
    
    If TypeOf oObject Is IJSystem Then
        Set oSystem = oObject
    End If
    
    If Not (oSystem Is Nothing) Then
        Set oCollection = oSystem.GetChildren
        
        For i = 1 To oCollection.Count
            
            If TypeOf oCollection.Item(i) Is IJCollarPart Then
                Set oCollarPart = oCollection.Item(i)
                oReturnColl.Add oCollarPart
            Else
                Dim oTempColl As IJDObjectCollection
                Set oTempColl = GetCollarParts(oCollection.Item(i))
                
                If Not oTempColl Is Nothing And oTempColl.Count > 0 Then
                
                    For Each oCollarPart In oTempColl
                        oReturnColl.Add oCollarPart
                        Set oCollarPart = Nothing
                    Next
                    Set oTempColl = Nothing
                End If
            End If
            Set oCollarPart = Nothing
        Next i
    End If
    
    Set GetCollarParts = oReturnColl
                                                                    
CleanUp:
    Set oSystem = Nothing
    Set oCollection = Nothing
    Set oReturnColl = Nothing

Exit Function
ErrorHelper:
    MsgBox Err.Description
End Function
