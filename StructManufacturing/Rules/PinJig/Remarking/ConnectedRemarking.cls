VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnectedRemarking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigRule
' Module: ConnectedRemarking
'
' Description:  Example to show the remarking lines due to plates/profiles
'               that are connected to the supported plates, on the pinjig
'
' Author: Anand Hariharan
'
'*******************************************************************************
Option Explicit
Private Const MODULE As String = "MfgPinJigRule.ConnectedRemarking"

Implements IJDMfgPinJigRule

Private Function IJDMfgPinJigRule_GetEntitiesForRemarking(ByVal PartialPinJig As Object) As GSCADMfgRulesDefinitions.IJElements
    Const METHOD = "IJDMfgPinJigRule_GetEntitiesForRemarking"
    On Error GoTo ErrorHandler

    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New PartSupport
    
    Dim ReturnColl As IJElements
    Set ReturnColl = New JObjectCollection
    
    Dim oMfgUtil As IJMfgUtilSurface
    Set oMfgUtil = New MfgUtilSurface

    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig
    
    Dim NormalX As Double
    Dim NormalY As Double
    Dim NormalZ As Double
    Dim RootPtX As Double
    Dim RootPtY As Double
    Dim RootPtZ As Double
    oPinJig.GetBasePlane NormalX, NormalY, NormalZ, _
                         RootPtX, RootPtY, RootPtZ
                         
    Dim oBasePlane As IJPlane
    Set oBasePlane = New Plane3d
    oBasePlane.DefineByPointNormal RootPtX, RootPtY, RootPtZ, _
                                   NormalX, NormalY, NormalZ
    
    Dim oPlateColl As IJElements
    Set oPlateColl = oPinJig.SupportedPlates

    Dim iter As Object
    ' For each supported plate ...
    For Each iter In oPlateColl
                
        Dim IsBaseSideSupported As Boolean
        oMfgUtil.GetPlateSideFacingPlane iter, oBasePlane, IsBaseSideSupported
        
        Set oPartSupport.Part = iter
        
        Dim ConnectedObjColl As Collection
        Dim ConnectionsColl As Collection
        Dim ThisPortColl As Collection
        Dim OtherPortColl As Collection
        ' ... get all objects connected to plate with a physical connection
        oPartSupport.GetConnectedObjects ConnectionPhysical, _
                                         ConnectedObjColl, ConnectionsColl, _
                                         ThisPortColl, OtherPortColl
        
        Dim i As Integer
        ' For each connected object ...
        For i = 1 To ConnectedObjColl.Count
            ' ... check if this connected object is another supported plate
            If oPlateColl.Contains(ConnectedObjColl.Item(i)) = False Then
            
                Dim oAppCon As IJAppConnection
                Set oAppCon = ConnectionsColl.Item(i)
                
                Dim ConnType As ContourConnectionType
                Dim WhichPartCross As Boolean
                oPartSupport.GetConnectionTypeForContour oAppCon, ConnType, WhichPartCross
                
            ' ... check if this connected object connects with supported plate with a TEE type connection
                If ConnType = PARTSUPPORT_CONNTYPE_TEE Then
                    
                    Dim oPort As IJStructPort
                    Set oPort = ThisPortColl.Item(i)
                    
            ' ... check if this connected object is on pin jig's remarking side
                    If ((oPort.ContextID & CTX_OFFSET) And IsBaseSideSupported = True) Or _
                       ((oPort.ContextID & CTX_BASE) And IsBaseSideSupported = False) Then
                       
                       ' IMPORTANT NOTE:
                       ' The Pin jig semantic expects *connections* not the /connected entities/.
                       ReturnColl.Add ConnectionsColl.Item(i)
                       
                    End If ' end check for remarking side
                    
                    Set oPort = Nothing
                    
                End If ' end check for TEE connection
                
                Set oAppCon = Nothing
                
            End If ' end check for supported plate
            
        Next ' end looping around connected objects
                
    Next ' end looping around supported plates
    
    Set IJDMfgPinJigRule_GetEntitiesForRemarking = ReturnColl
    
    Set oBasePlane = Nothing
    Set oMfgUtil = Nothing
    Set oPlateColl = Nothing
    Set ReturnColl = Nothing
    Set oPartSupport = Nothing
    Set oPinJig = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

Private Function IJDMfgPinJigRule_GetRemarkingGeometry(ByVal PartialPinJig As Object) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "IJDMfgPinJigRule_GetRemarkingGeometry"
    On Error GoTo ErrorHandler


    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

