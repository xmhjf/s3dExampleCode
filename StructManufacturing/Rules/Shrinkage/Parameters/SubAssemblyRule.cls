VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SubAssemblyRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgShrinkageRules
' Module: GrandBlockRule
'
' Description:  Determines the proposed settings for the GrandBlock type assembly
'
' Author:
'
' Comments:
' 2012.JUNE.21    Raman     New desgin of the class
'*******************************************************************************
Option Explicit

Const MODULE = "AssemblyRule:"

Implements IJDShrinkageRule

Private Function IJDShrinkageRule_GetConnectedObjectsForPart(ByVal pDispObj As Object) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
    Const METHOD = "IJDShrinkageRule_GetConnectedObjectsForPart"
    On Error GoTo ErrorHandler
    
    'No implementation
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub IJDShrinkageRule_GetDependentProfileShrinkageParameters(ByVal pDispObj As Object, ByVal pPlateColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal pPlatePrimaryFactorColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal pPlateSecondaryFactorColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal pPlatePrimaryAxisColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal pPlateSecondaryAxisColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, PrimaryFactor As Double)
    Const METHOD = "IJDShrinkageRule_GetDependentProfileShrinkageParameters"
    On Error GoTo ErrorHandler
        
    'No implementation
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function GetGlobalCoordinateSystem(ByVal pDispObj As Object) As Object
    Const METHOD = "GetGlobalCoordinateSystem"
    On Error GoTo ErrorHandler
    
    Dim oCSservice As New CMfgCoordSys
    
    Dim FSFact As IJElements
    Set FSFact = oCSservice.FrameSystemsFromCatalogRule("CScalingShr_FrameSystem", pDispObj)
   
    Dim nIndex As Long
    Dim FS As IHFrameSystem
   
    For nIndex = 1 To FSFact.Count
        Set FS = FSFact(nIndex)
        If Not FS Is Nothing Then
            If "CS_0" = FS.Name Then
                Set GetGlobalCoordinateSystem = FS
                Exit For
            End If
        End If
    Next nIndex

    If GetGlobalCoordinateSystem Is Nothing Then
        If FSFact.Count > 0 Then
            Set GetGlobalCoordinateSystem = FSFact(1)
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub IJDShrinkageRule_GetShrinkageParameters(ByVal pDispObj As Object, ShrinkageType As StrMfgShrinkageType, PrimaryAxis As Object, PrimaryFactor As Double, SecondaryAxis As Object, SecondaryFactor As Double, Optional TertiaryAxis As Object, Optional TertiaryFactor As Double)
    Const METHOD = "IJDShrinkageRule_GetShrinkageParameters"
    On Error GoTo ErrorHandler
    
    If TypeOf pDispObj Is IJAssembly Then
        Dim oAssemblyBase       As IJAssemblyBase
        
        Set oAssemblyBase = pDispObj
        
        If oAssemblyBase.Type = 1 Then 'Assembly
          
            'get the stiffened plate
            Dim oPlatePart As IJPlatePart
            Dim oAssembly As IJAssembly
            
            Set oAssembly = oAssemblyBase
            
            Dim oChildren As IJDTargetObjectCol
            Set oChildren = oAssembly.GetChildren
            
            If oChildren.Count > 0 Then
                'Set up elements list
                Dim Index As Long
                            
                'Get each assembly child and add child to elements list
                For Index = 1 To oChildren.Count
                    Dim oItem As Object
                    
                    'Get next item
                    Set oItem = oChildren.Item(Index)
                                                   
                    'Add item to elements list if type of item
                    If TypeOf oItem Is IJPlatePart Then
                        Set oPlatePart = oItem
                    End If
                Next
            Else
                'nothing to do
            End If

            
            If Not oPlatePart Is Nothing Then
                Dim oNumberL As Integer
                Dim oNumberT As Integer
                Dim oNumberV As Integer
                Dim oTempObj As Variant
                
                'Initialize values
                oNumberL = 0
                oNumberT = 0
                oNumberV = 0
                
                    
                Dim oSDPlatePart As New StructDetailObjects.PlatePart
                Set oSDPlatePart.object = oPlatePart
                Dim collConnectedObj As Collection
            
                Set collConnectedObj = oSDPlatePart.ConnectedObjects
            
            '    'Give me the number of physical connections
                
                ''Count the number of attached profiles in a certain direction, if they are not longitudenal then assume that they are transversal
                '
                For Each oTempObj In collConnectedObj
                    Dim oConData As ConnectionData
                    oConData = oTempObj
                    Set oTempObj = oConData.ToConnectable
                    Dim oStiffner  As IJStiffener
                    
                    If TypeOf oTempObj Is IJStiffener Then
                        Set oStiffner = oTempObj
                        If oStiffner.pType = sptLongitudinal Then
                            oNumberL = oNumberL + 1
                        ElseIf oStiffner.pType = sptTransversal Then
                            oNumberT = oNumberT + 1
                        ElseIf oStiffner.pType = sptVertical Then
                            oNumberV = oNumberV + 1
                        End If
                    End If
                Next
                
                Select Case oSDPlatePart.PlateType
                
                Case DeckPlate
                
                    If oNumberL > oNumberT Then
                        'deck stiffened along longitudinal
                        ShrinkageType = GlobalType
                        Dim oFrameSystem As IHFrameSystem
                        Set oFrameSystem = GetGlobalCoordinateSystem(oPlatePart)
                        
                        If Not oFrameSystem Is Nothing Then
                            Set PrimaryAxis = oFrameSystem.PrincipalXAxis
                            Set SecondaryAxis = oFrameSystem.PrincipalYAxis
                            'Set TertiaryAxis = oFrameSystem.PrincipalZAxis
                        End If
            
                        PrimaryFactor = 0.2
                        SecondaryFactor = 0.4
                        TertiaryFactor = 0#
                    Else
                        'deck stiffened along transversal
                        ShrinkageType = GlobalType
                        
                        If Not oFrameSystem Is Nothing Then
                            Set PrimaryAxis = oFrameSystem.PrincipalYAxis
                            Set SecondaryAxis = oFrameSystem.PrincipalXAxis
                            'Set TertiaryAxis = oFrameSystem.PrincipalZAxis
                        End If
            
                        PrimaryFactor = 0.2
                        SecondaryFactor = 0.4
                        TertiaryFactor = 0#
                    End If
                End Select
                
            Else
                ShrinkageType = ShrinkageUndefined 'No shirnkage needed
            End If
        Else
            ShrinkageType = ShrinkageUndefined 'No shirnkage needed
        End If
    Else
        ShrinkageType = ShrinkageUndefined 'No shirnkage needed
    End If

    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 7004, , "RULES")
End Sub





