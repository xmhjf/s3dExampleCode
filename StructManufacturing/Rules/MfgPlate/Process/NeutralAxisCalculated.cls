VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Calculated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess
' Module: AntiMoldedSide
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Private Const MODULE = "MfgPlateProcess.Calculated"

Implements IJDMfgPlateNeutralAxisRule

Public Function IsFlangedBracket(oPlatePart As IJPlatePart, ByRef dNARatio As Double, ByRef eInnerSurfSide As enumPlateSide) As Boolean
    Dim bIsFlanged As Boolean
    
    bIsFlanged = False
    dNARatio = 0#
    IsFlangedBracket = False
    eInnerSurfSide = BaseSide
    ' Get the root system of this part
    Dim oParentSystem As IJSystem
    Dim oStructDetailHelper As GSCADStructDetailUtil.StructDetailHelper
    Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
    On Error Resume Next
    ' True below means "recursive" - Navigate past design splits
    ' to retrieve the root system

    oStructDetailHelper.IsPartDerivedFromSystem oPlatePart, oParentSystem, True
    
'    Set oStructDetailHelper = Nothing
    If oParentSystem Is Nothing Then
        ' Part is a stand-alone plate part
        ' These cannot be flanged (yet)
        Exit Function
    End If
    Dim oPlateSystem As IJPlateSystem
    Set oPlateSystem = oParentSystem
    Dim oPlateFlange_AE As IJPlateFlange_AE
    
    ' Get the flange AE (if any)
    
    'reserved parameter below is to support eventual multiple
    ' flanges.  If a plate system has multiple free-edges, each
    ' could theoretically be flanged.  This is not yet supported
    Set oPlateFlange_AE = oPlateSystem.FlangeActiveEntity(Nothing)
    If oPlateFlange_AE Is Nothing Then
        ' system is not flanged - part cannot be
        GoTo CleanUp
        Exit Function
    End If
    Dim oFlangedLeafPlate As IJPlate
    
    ' ask the flange AE for the affected leaf system
    
    ' true means get the affected leaf system
    ' currently, flanges cannot cross design splits
    Set oFlangedLeafPlate = oPlateFlange_AE.GetAffectedPlate(True)
    'get the immediate parent (leaf system) of the input part
    Dim oLeafSystem As IJSystem
    oStructDetailHelper.IsPartDerivedFromSystem oPlatePart, oLeafSystem, False      ' Used to be True
    Dim oLeafPlate As IJPlate
    Set oLeafPlate = oLeafSystem

    ' see if this part's leaf system is the one that is flanged
    If oLeafPlate Is oFlangedLeafPlate Then
        dNARatio = oPlateFlange_AE.NeutralAxisRatio
        
        ' Validate the N.A.Ratio returned by the flanged bracket symbol
        If dNARatio < 0 Then
            If Abs(dNARatio) < 0.001 Then
                dNARatio = 0#
            Else
                dNARatio = 0.5
            End If
        ElseIf dNARatio > 1 Then
            If dNARatio < 1.001 Then
                dNARatio = 1#
            Else
                dNARatio = 0.5
            End If
        End If
        
        If oPlateFlange_AE.IsBaseFaceInnerFaceOfBend = True Then
            eInnerSurfSide = BaseSide
        Else
            eInnerSurfSide = OffsetSide
        End If
        bIsFlanged = True
    End If

    IsFlangedBracket = bIsFlanged
    
CleanUp:
    Set oLeafPlate = Nothing
    Set oStructDetailHelper = Nothing
    Set oFlangedLeafPlate = Nothing
    Set oPlateFlange_AE = Nothing
    Set oPlateSystem = Nothing
    Set oParentSystem = Nothing
    
End Function
Private Function IJDMfgPlateNeutralAxisRule_GetNeutralAxis(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As Double
    Const METHOD = "NeutralAxisCalculated: IJDMfgPlateNeutralAxisRule_GetNeutralAxis"
    On Error Resume Next
    
    ' Use the plate part as input
    Dim dNARatio As Double
    Dim eInnerSurfSide As enumPlateSide
    Dim oSDPlatePart As New StructDetailObjects.PlatePart
    Set oSDPlatePart.object = pPlatePart
    
    If IsFlangedBracket(pPlatePart, dNARatio, eInnerSurfSide) Then
        Dim oMfgPlateWrapper As MfgRuleHelpers.PlatePartHlpr
        Dim oMfgPlatePart As IJMfgPlatePart
        Set oMfgPlateWrapper = CreateObject("MfgRuleHelpers.PlatePartHlpr")
        Set oMfgPlateWrapper.object = pPlatePart
        If oMfgPlateWrapper.PlateHasMfgPart(oMfgPlatePart) Then
            Dim ePlateUpside As enumPlateSide
            
            ePlateUpside = oMfgPlatePart.UpSide
            
            If ePlateUpside = eInnerSurfSide Then
                IJDMfgPlateNeutralAxisRule_GetNeutralAxis = dNARatio
            Else
                IJDMfgPlateNeutralAxisRule_GetNeutralAxis = 1# - dNARatio
            End If

            Set oMfgPlateWrapper = Nothing
            Exit Function
        End If
        Set oMfgPlateWrapper = Nothing
    End If
      
    Dim Neutral As Double
    Dim Material As String
        
    Dim Thickness As Double
    Neutral = 0.5
    Material = oSDPlatePart.MaterialType
    ' Convert from m to mm
    Thickness = oSDPlatePart.PlateThickness * 1000#
        
    Select Case Material
    Case "FE"
            If Thickness < 20 Then
                IJDMfgPlateNeutralAxisRule_GetNeutralAxis = 0.5
            Else
                If Thickness < 30 Then
                  IJDMfgPlateNeutralAxisRule_GetNeutralAxis = 0.48
                Else
                  IJDMfgPlateNeutralAxisRule_GetNeutralAxis = 0.45
               End If
            End If
     Case "ALU"
            If Thickness < 10 Then
                IJDMfgPlateNeutralAxisRule_GetNeutralAxis = 0.5
            Else
                If Thickness < 20 Then
                  IJDMfgPlateNeutralAxisRule_GetNeutralAxis = 0.48
                Else
                  IJDMfgPlateNeutralAxisRule_GetNeutralAxis = 0.45
                End If
            End If
    Case Else
        IJDMfgPlateNeutralAxisRule_GetNeutralAxis = 0.5
    End Select
    
CleanUp:
    Set oSDPlatePart = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1044, , "RULES")
    GoTo CleanUp
End Function
 
