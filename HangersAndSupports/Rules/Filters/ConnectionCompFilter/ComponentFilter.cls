VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComponentFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   ComponentFilter.cls
'   ProgID:         ConnCompFilter.ComponentFilter
'   Author:         Yan Ji
'   Creation Date:  28.Apr.2003
'   Description:
'
'
'   Change History:
'       28.Apr.2003             Yan Ji       Creation
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "ComponentFilter"  'Used for error messages

'==============================
'component filter for HgrBeam
'===============================
Public Function ComponentFilter1(pComp As Object) As Long
Const METHOD = "ComponentFilter1"
On Error GoTo ErrorHandler

    ComponentFilter1 = 0
    
    Dim pIJDPart As IJDPart
    Dim pIJDPartOcc As IJPartOcc
    Dim partClassName As String
    
    If TypeOf pComp Is IJPartOcc Then
        Set pIJDPartOcc = pComp
        
        If pIJDPartOcc Is Nothing Then
            Exit Function
        End If
        
        If Not TypeOf pComp Is IJHgrConnComponent Then
            Exit Function
        End If
        
        'pIJDPartOcc.GetPart pIJDPart
        'partClassName = pIJDPart.GetRelatedPartClassName
        'If partClassName = "HgrBeam" Then
            ComponentFilter1 = 1
        'End If
        
        Set pIJDPartOcc = Nothing
        Set pIJDPart = Nothing
    End If
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'==============================
'component filter for HgrProfile
'===============================
Public Function ComponentFilter2(pComp As Object) As Long
Const METHOD = "ComponentFilter2"
On Error GoTo ErrorHandler

    ComponentFilter2 = 0
    
    Dim pIJDPart As IJDPart
    Dim pIJDPartOcc As IJPartOcc
    Dim partClassName As String
    
    If TypeOf pComp Is IJPartOcc Then
        Set pIJDPartOcc = pComp
        
        If pIJDPartOcc Is Nothing Then
            Exit Function
        End If
        
        If Not TypeOf pComp Is IJHgrConnComponent Then
            Exit Function
        End If
        
        pIJDPartOcc.GetPart pIJDPart
        partClassName = pIJDPart.GetRelatedPartClassName
        If partClassName = "StructProfile" Or partClassName = "G4G_7878_01" Then
            ComponentFilter2 = 1
        End If
        
        Set pIJDPartOcc = Nothing
        Set pIJDPart = Nothing
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'====================================
'filter for IJPlanes of ISPSWallParts
'=====================================
Public Function WallPlaneFilter(pComp As Object) As Long

Const METHOD = "WallPlaneFilter"
On Error GoTo ErrorHandler

    WallPlaneFilter = 0
    
    Dim pIJPlane As IJPlane
    
    On Error Resume Next
    Set pIJPlane = pComp
    On Error GoTo ErrorHandler
    
       If Not pIJPlane Is Nothing Then
         Dim pIJPort As IJPort
        
         On Error Resume Next
         Set pIJPort = pIJPlane
         On Error GoTo ErrorHandler
        
            If Not pIJPort Is Nothing Then
                Dim pIJConnectable  As IJConnectable
                Set pIJConnectable = pIJPort.Connectable
                Dim pWallPart As ISPSWallPart
            
                On Error Resume Next
                Set pWallPart = pIJConnectable
                On Error GoTo ErrorHandler
            
                If Not pWallPart Is Nothing Then
                    WallPlaneFilter = 1
                End If
            
                Set pIJConnectable = Nothing
                Set pWallPart = Nothing

           End If
                Set pIJPort = Nothing
     Else
            
            Dim pWallPart2 As ISPSWallPart
            On Error Resume Next
            Set pWallPart2 = pComp
            On Error GoTo ErrorHandler
            
            If Not pWallPart2 Is Nothing Then
                WallPlaneFilter = 1
            End If
            
            Set pWallPart2 = Nothing
     End If
     
     Set pIJPlane = Nothing
     
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function
Public Function SlabPlaneFilter(pComp As Object) As Long

Const METHOD = "SlabPlaneFilter"
On Error GoTo ErrorHandler
   

    SlabPlaneFilter = 0
    
    Dim pIJPlane As IJPlane
    
    On Error Resume Next
    Set pIJPlane = pComp
    On Error GoTo ErrorHandler
    
       If Not pIJPlane Is Nothing Then
         Dim pIJPort As IJPort
        
         On Error Resume Next
         Set pIJPort = pIJPlane
         On Error GoTo ErrorHandler
        
            If Not pIJPort Is Nothing Then
                Dim pIJConnectable  As IJConnectable
                Set pIJConnectable = pIJPort.Connectable
                Dim pSlabPart As ISPSSlabEntity
            
                On Error Resume Next
                Set pSlabPart = pIJConnectable
                On Error GoTo ErrorHandler
            
                If Not pSlabPart Is Nothing Then
                    SlabPlaneFilter = 1
                End If
            
                Set pIJConnectable = Nothing
                Set pSlabPart = Nothing

           End If
                Set pIJPort = Nothing
     Else
            
            Dim pSlabPart2 As ISPSSlabEntity
            
            
            On Error Resume Next
            Set pSlabPart2 = pComp
            On Error GoTo ErrorHandler
            
            If Not pSlabPart2 Is Nothing Then
                SlabPlaneFilter = 1
            End If
            
            Set pSlabPart2 = Nothing
     End If
     
     Set pIJPlane = Nothing
     
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Function EndFeatureFilter(pFeature As Object) As Long
  Const METHOD = "EndFeatureFilter"
  On Error GoTo ErrorHandler
 
  EndFeatureFilter = 0
  Dim oFeat As IJRteEndPathFeat
  Dim bIsAnIEF As Boolean
  Set oFeat = pFeature
 If Not oFeat Is Nothing Then
 
   Dim oIJRtePathFeature As IJRtePathFeat
   Dim oElements As IJElements
   Dim lPartCnt As Long
                           
   Set oIJRtePathFeature = oFeat
   Set oElements = oIJRtePathFeature.GetBaseParts(lPartCnt)
              If (lPartCnt > 0) Then EndFeatureFilter = 1
              
  bIsAnIEF = oFeat.IsIntermediateFeature()
  
  If bIsAnIEF Then
    EndFeatureFilter = 1
   End If
 End If
  
  Set oFeat = Nothing
Exit Function

ErrorHandler:
 Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Function GenericSurfaceFilter(pComp As Object) As Long
Const METHOD = "GenericSurfaceFilter"

    
    GenericSurfaceFilter = 0
    
    Dim oPort As IJPort
    Dim bWallOrSlabOrTrenchObject As Boolean
    
    On Error Resume Next
    Set oPort = pComp
    If oPort Is Nothing Then
        bWallOrSlabOrTrenchObject = GenericPlaneFilter(pComp)
         
        If (bWallOrSlabOrTrenchObject = False) Then
            GenericSurfaceFilter = 1
        End If
    End If

End Function

Public Function GenericPlaneFilter(pComp As Object) As Long
Const METHOD = "GenericPlaneFilter"

    GenericPlaneFilter = 0
    
    Dim pIJPlane As IJPlane
    
    On Error Resume Next
    Set pIJPlane = pComp
    On Error GoTo ErrorHandler
    
     If Not pIJPlane Is Nothing Then
         Dim pIJPort As IJPort
        
         On Error Resume Next
         Set pIJPort = pIJPlane
         On Error GoTo ErrorHandler
        
            If Not pIJPort Is Nothing Then
                Dim pIJConnectable  As IJConnectable
                Set pIJConnectable = pIJPort.Connectable
                
                If Not pIJConnectable Is Nothing Then
                    Dim sSupportingObjType As String
                    sSupportingObjType = ""
                    
                    Select Case TypeName(pIJConnectable)
                    Case "ISPSWallPart"
                        sSupportingObjType = "Wall"
                    Case "ISPSSlabEntity"
                        sSupportingObjType = "Slab"
                    Case "IJTrenchPart"
                        sSupportingObjType = "Trench"
                    End Select
              
                    If (sSupportingObjType = "Wall") Or (sSupportingObjType = "Slab") Or (sSupportingObjType = "Trench") Then
                        GenericPlaneFilter = 1
                    End If
                End If
            
                Set pIJConnectable = Nothing

           End If
                Set pIJPort = Nothing
     Else
            
            Dim sSuppObjType As String
            sSuppObjType = ""
            
            Select Case TypeName(pComp)
            Case "ISPSWallPart"
                sSuppObjType = "Wall"
            Case "ISPSSlabEntity"
                sSuppObjType = "Slab"
            Case "IJTrenchPart"
                sSuppObjType = "Trench"
            End Select
            
            
            If (sSuppObjType = "Wall") Or (sSuppObjType = "Slab") Or (sSuppObjType = "Trench") Then
                GenericPlaneFilter = 1
            End If
            
     End If
     
     Set pIJPlane = Nothing
     
     Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Function

' TR#141961 ; check that it is not a logical comp and not comp of same support
Public Function IsValidSupportingObject(pCandidateObj As Object, SupArg As Object) As Long
On Error Resume Next
    IsValidSupportingObject = 1
    If Not SupArg Is Nothing Then
        Dim oSupport As IJHgrSupportInternalInfo
        Dim bValidSupporting As Boolean
        bValidSupporting = True
        Set oSupport = SupArg
        If Not oSupport Is Nothing Then
            bValidSupporting = oSupport.DoesCompBelongToSameSupport(pCandidateObj)
            If bValidSupporting = True Then
                IsValidSupportingObject = 1
            Else
                IsValidSupportingObject = 0
            End If
            Set oSupport = Nothing
        End If
    End If
End Function

