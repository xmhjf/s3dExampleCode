VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FittingSelRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
' Copyright (C) 2006 Intergraph Corporation.  All Rights Reserved.
'
' File: FittingSelRules.cls
'
' Author: sypark@ship.samsung.co.kr
'
' Abstract: Set the proper fitting type to HoleTrace according to the selection rule.
'
' Description:
'   To allow the customization of fittings, the selection rule will always be
'   required when using the place fitting command in Hole Mgmt. This rule will
'   provide the following functionality:
'
'       1. Determine the appropriate fitting type (if user requests)
'       2. Determine the appropriate catalog part (if user requests)
'       3. Create the actual fittings based on selected type
'       4. Update the actual fittings based on selected type
'       5. Flip the actual fittings based on selected type
'
'   If the user has selected fitting by rule, the middle tier will call the
'   fitting rule to determine the fitting type. If the rule decides a catalog
'   fitting is required, the middle tier will call the rule again to pick the
'   appropriate catalog part.
'
'   When it comes time to create the actual fitting, the middle tier will once
'   again call the rule to create the fitting based on the type.
'
'   If the hole changes for any reason, the middle tier will notify the rule
'   to allow the update of the fitting if it is needed.
'
'   If the user requires that the fitting be flipped, the middle tier will
'   also notify the rule to flip the fitting.
'
' Changes:
'   Date    By         Reason
' 06/13/06  PML       Updated for changes made in V7
'
'******************************************************************************

Option Explicit

Private Const Module = "HMFittingSelectRules.FittingSelRules"
Private Const PROG_ID = "HMFittingSelectRules.FittingSelRules"

Private m_oErrors As New IMSErrorLog.JServerErrors

Implements IJHMFittingSelRule

'//////////////////////////////////////////////////////////////////////////////
'******************************************************************************
'Method: GetType
'
'Interface: IJHMFittingSelRule
'
'Abstract: Return the proper fitting type based on the input structure
'          and outfitting of the hole trace
'
'   Input:
'       oHoleTrace - hole trace
'
'   Output:
'       eFittingType - fitting type
'
'Attention :  selection rule is followings.
'
'                Tight Deck      Non-tight Deck      Tight Wall     Non-tight Wall      Others
'   Pipe       Center Flange    On deck coaming     Center Flange       None              None
'   HVAC       Center Flange    On deck coaming     Center Flange       None              None
'   Cableway   Cable coaming    Cable coaming       Cable coaming   Cable coaming    Cable coaming
'   Equipment      None               None              None            None              None
'   None           None         On deck coaming         None            None              None
'
'******************************************************************************
Public Sub IJHMFittingSelRule_GetType(ByVal oHoleTrace As HoleMgmtRulesItfLib.IJHoleTraceAE, _
                                      ByRef eFittingType As HoleMgmtRulesItfLib.HMFittingType)
    Const METHOD = "IJHMFittingSelRule_GetType"
    On Error GoTo ErrorHandler
    
    ' Get the outfitting to check what kind of hole this is
    Dim oOutfitting As IJDObjectCollection
    Set oOutfitting = oHoleTrace.GetParentOutfitting

    ' Get the name of outfitting application
    Dim eOutfittingApp As HMHoleRequestApp
    eOutfittingApp = GetRequestApp(oOutfitting)

    ' Get the structure the hole is on
    Dim oStructure As Object
    Set oStructure = oHoleTrace.GetParentStructure

    ' If the structure isn't a plate, fitting type should be set as none
    Dim oPlate As IJPlate
    If TypeOf oStructure Is IJPlate Then
        Set oPlate = oStructure
    Else
        eFittingType = HM_FittingNone
        GoTo Cleanup
    End If

    ' Get the tightness and type of plate
    Dim ePlateTightness As StructPlateTightness
    Dim ePlateType As StructPlateType
     
    ePlateTightness = oPlate.Tightness
    ePlateType = oPlate.plateType

'******************************************************************************
' Piping Application
'******************************************************************************
    
    If eOutfittingApp = HM_PipingApp Then
        ' Plate type is Deck
        If ePlateType = DeckPlate Then
            ' Plate Tightness is Non-Tight
            If ePlateTightness = NonTight Then
                eFittingType = HM_FittingCoaming
            ' Plate Tightness Is Tight
            Else
                eFittingType = HM_FittingCenterFlange
            End If
        ' Plate type is Bulkhead
        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
            ' Plate Tightness is Non-Tight
            If ePlateTightness = NonTight Then
                eFittingType = HM_FittingNone
            ' Plate Tightness Is Tight
            Else
                eFittingType = HM_FittingCenterFlange
            End If
        ' Hull or other type of plate
        Else
            eFittingType = HM_FittingNone
        End If
    
'******************************************************************************
' HVAC Application
'******************************************************************************
    
    ElseIf eOutfittingApp = HM_HVACApp Then
        ' Plate type is Deck
        If ePlateType = DeckPlate Then
            ' Plate Tightness is Non-Tight
            If ePlateTightness = NonTight Then
                eFittingType = HM_FittingCoaming
            ' Plate Tightness Is Tight
            Else
                eFittingType = HM_FittingCenterFlange
            End If
        ' Plate type is Bulkhead
        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
            ' Plate Tightness is Non-Tight
            If ePlateTightness = NonTight Then
                eFittingType = HM_FittingNone
            ' Plate Tightness Is Tight
            Else
                eFittingType = HM_FittingCenterFlange
            End If
        ' Hull or other type of plate
        Else
            eFittingType = HM_FittingNone
        End If
    
'******************************************************************************
' Cableway Application
'******************************************************************************
    
    ElseIf eOutfittingApp = HM_CableApp Then
        eFittingType = HM_FittingCatalogPart
        
'******************************************************************************
' Equipment Application
'******************************************************************************
    
    ElseIf eOutfittingApp = HM_EquipmentApp Then
        eFittingType = HM_FittingNone
        
'******************************************************************************
' Multiple Applications
'******************************************************************************
    
    ElseIf eOutfittingApp = HM_OtherApp Then
        eFittingType = HM_FittingNone
        
'******************************************************************************
' Stand-Alone Hole
'******************************************************************************
    
    ElseIf eOutfittingApp = HM_StandAloneApp Then
        ' Plate type is Deck
        If ePlateType = DeckPlate Then
            ' Plate Tightness is Non-Tight
            If ePlateTightness = NonTight Then
                eFittingType = HM_FittingCoaming
            ' Plate Tightness Is Tight
            Else
                eFittingType = HM_FittingNone
            End If
        ' Plate type is Bulkhead
        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
            ' Plate Tightness is Non-Tight
            If ePlateTightness = NonTight Then
                eFittingType = HM_FittingNone
            ' Plate Tightness Is Tight
            Else
                eFittingType = HM_FittingNone
            End If
        ' Hull or other type of plate
        Else
            eFittingType = HM_FittingNone
        End If
        
'******************************************************************************
' Conduit Application
'******************************************************************************
    
    ElseIf eOutfittingApp = HM_ConduitApp Then
        ' Plate type is Deck
        If ePlateType = DeckPlate Then
            ' Plate Tightness is Non-Tight
            If ePlateTightness = NonTight Then
                eFittingType = HM_FittingCoaming
            'Plate Tightness Is Tight
            Else
                eFittingType = HM_FittingCatalogPart
            End If
        'Plate type is Bulkhead
        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
            ' Plate Tightness is Non-Tight
            If ePlateTightness = NonTight Then
                eFittingType = HM_FittingNone
            'Plate Tightness Is Tight
            Else
                eFittingType = HM_FittingCatalogPart
            End If
        ' Hull or other type of plate
        Else
            eFittingType = HM_FittingNone
        End If
        
'******************************************************************************
' Unknown Application
'******************************************************************************
    
    'In Case of Unknown Application
    ElseIf eOutfittingApp = HM_UnknownApp Then
        eFittingType = HM_FittingNone
        
'******************************************************************************
' Default
'******************************************************************************
    
    Else
        eFittingType = HM_FittingNone
    End If
   
Cleanup:
    Set oOutfitting = Nothing
    Set oStructure = Nothing
    Set oPlate = Nothing
    
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, Module & " - " & METHOD, Err.Description
    GoTo Cleanup
End Sub

'//////////////////////////////////////////////////////////////////////////////
'******************************************************************************
'Method: GetPart
'
'Interface: IJHMFittingSelRule
'
'Abstract: Return the proper part class and part name based on the
'          input structure and outfitting of the hole trace
'
'   Input:
'       oHoleTrace - holetrace
'
'   Output:
'       strPartClassName - partclass name of catalog part
'       strPartClass     - part name of catalog part
'
'Attention :
'
'******************************************************************************
Public Sub IJHMFittingSelRule_GetPart(ByVal oHoleTrace As IJHoleTraceAE, _
                                      ByRef strPartClassName As String, _
                                      ByRef strPartName As String)
    Const METHOD = "IJHMFittingSelRule_GetPart"
    On Error GoTo ErrorHandler
  
    Dim oObject As Object
    Dim oOutfitting As IJDObjectCollection
    
    'Get the Collection of Outfittings
    Set oOutfitting = oHoleTrace.GetParentOutfitting
    For Each oObject In oOutfitting

'******************************************************************************
' Cableway rules
'******************************************************************************
        
        If TypeOf oObject Is IJRteCablewayPathFeat Then
            'Get the Outfitting, Only Cable Way should be taken.
            Dim oSelectedType As CrossSectionShapeTypes, oWidth As Double, oDepth As Double
            GetWidthDepthOfCableway oHoleTrace, oSelectedType, oWidth, oDepth
            
            'Get the information of structure.
            Dim ePlateType As StructPlateType, ePlateTightness As StructPlateTightness
            GetStructInfo oHoleTrace, ePlateType, ePlateTightness
            
            'Check the CrossSection type of CW and check the width and Depth
            
'******************************************************************************
' Rectangular or FlatOval cableway
'******************************************************************************
            
            If oSelectedType = Rectangular Or oSelectedType = FlatOval Then
               
'******************************************************************************
' oWidth > 0 And oWidth <= 0.2
'******************************************************************************
               
               If oWidth > 0 And oWidth <= 0.2 Then
                    If oDepth > 0 And oDepth <= 0.15 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-2"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-2"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-2"
                            strPartClassName = "CableCoamingCAType"
                        End If
        
                    ElseIf oDepth > 0.15 And oDepth <= 0.25 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-22"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-22"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-22"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.25 And oDepth <= 0.4 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-23"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-23"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-23"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.4 And oDepth <= 0.55 Then
                         If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-24"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-24"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-24"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.55 And oDepth <= 0.7 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-25"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-25"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-25"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    End If
                
'******************************************************************************
' oWidth > 0.2 And oWidth <= 0.3
'******************************************************************************
                
                ElseIf oWidth > 0.2 And oWidth <= 0.3 Then
                    If oDepth > 0 And oDepth <= 0.15 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-3"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-3"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-3"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.15 And oDepth <= 0.25 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-32"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-32"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-32"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.25 And oDepth <= 0.4 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-33"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-33"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-33"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.4 And oDepth <= 0.55 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-34"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-34"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-34"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.55 And oDepth <= 0.7 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-35"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-35"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-35"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    End If
                
'******************************************************************************
' oWidth > 0.3 And oWidth <= 0.4
'******************************************************************************
                
                ElseIf oWidth > 0.3 And oWidth <= 0.4 Then
                    If oDepth > 0 And oDepth <= 0.15 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-4"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-4"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-4"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.15 And oDepth <= 0.25 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-42"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-42"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-42"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.25 And oDepth <= 0.4 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-43"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-43"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-43"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.4 And oDepth <= 0.55 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-44"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-44"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-44"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.55 And oDepth <= 0.7 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-45"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-45"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-45"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    End If
                
'******************************************************************************
' oWidth > 0.4 And oWidth <= 0.5
'******************************************************************************
                
                ElseIf oWidth > 0.4 And oWidth <= 0.5 Then
                    If oDepth > 0 And oDepth <= 0.15 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-5"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-5"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-5"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.15 And oDepth <= 0.25 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-52"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-52"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-52"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.25 And oDepth <= 0.4 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-53"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-53"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-53"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.4 And oDepth <= 0.55 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-54"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-54"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-54"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.55 And oDepth <= 0.7 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-55"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-55"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-55"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    End If
               
'******************************************************************************
' oWidth > 0.5 And oWidth <= 0.6
'******************************************************************************
                
                ElseIf oWidth > 0.5 And oWidth <= 0.6 Then
                    If oDepth > 0 And oDepth <= 0.15 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-6"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-6"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-6"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.15 And oDepth <= 0.25 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-62"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-62"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-62"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.25 And oDepth <= 0.4 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-63"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-63"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-63"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.4 And oDepth <= 0.55 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-64"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-64"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-64"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.55 And oDepth <= 0.7 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-65"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-65"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-65"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    End If
                
'******************************************************************************
' oWidth > 0.6 And oWidth <= 0.7
'******************************************************************************
                
                ElseIf oWidth > 0.6 And oWidth <= 0.7 Then
                    If oDepth > 0 And oDepth <= 0.15 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-7"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-7"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-7"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.15 And oDepth <= 0.25 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-72"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-72"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-72"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.25 And oDepth <= 0.4 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-73"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-73"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-73"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.4 And oDepth <= 0.55 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-74"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-74"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-74"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.55 And oDepth <= 0.7 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-75"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-75"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-75"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    End If
                
'******************************************************************************
' oWidth > 0.7 And oWidth <= 0.8
'******************************************************************************
                
                ElseIf oWidth > 0.7 And oWidth <= 0.8 Then
                    If oDepth > 0 And oDepth <= 0.15 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-8"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-8"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-8"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.15 And oDepth <= 0.25 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-82"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-82"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-82"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.25 And oDepth <= 0.4 Then
                         If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-83"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-83"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-83"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.4 And oDepth <= 0.55 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-84"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-84"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-84"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.55 And oDepth <= 0.7 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-85"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-85"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-85"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    End If
          
'******************************************************************************
' oWidth > 0.8 And oWidth <= 0.9
'******************************************************************************
                
                ElseIf oWidth > 0.8 And oWidth <= 0.9 Then
                    If oDepth > 0 And oDepth <= 0.15 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-9"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-9"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-9"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.15 And oDepth <= 0.25 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-92"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-92"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-92"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.25 And oDepth <= 0.4 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-93"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-93"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-93"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.4 And oDepth <= 0.55 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-94"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-94"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-94"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.55 And oDepth <= 0.7 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-95"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-95"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-95"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    End If
          
'******************************************************************************
' oWidth > 0.9 And oWidth <= 1
'******************************************************************************
                
                ElseIf oWidth > 0.9 And oWidth <= 1 Then
                    If oDepth > 0 And oDepth <= 0.15 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-10"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-10"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-10"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.15 And oDepth <= 0.25 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-102"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-102"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-102"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.25 And oDepth <= 0.4 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-103"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-103"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-103"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.4 And oDepth <= 0.55 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-104"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-104"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-104"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    
                    ElseIf oDepth > 0.55 And oDepth <= 0.7 Then
                        If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                strPartName = "C-105"
                                strPartClassName = "CableCoamingCType"
                            Else
                                strPartName = "K-105"
                                strPartClassName = "CableCoamingKType"
                            End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-105"
                            strPartClassName = "CableCoamingCAType"
                        End If
                    End If
                End If
                
'******************************************************************************
' Round cableway
'******************************************************************************
            
            ElseIf oSelectedType = 4 Then
                If oWidth = 0.0605 Then
                    If ePlateType = DeckPlate Then
                            If ePlateTightness = NonTight Then
                                    strPartName = "C-50"
                                    strPartClassName = "CableCoamingCType"
                            Else
                                    strPartName = "K-50"
                                    strPartClassName = "CableCoamingKType"
                            End If
                    ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                            strPartName = "CA-50"
                            strPartClassName = "CableCoamingCAType"
                    End If
                ElseIf oWidth = 0.0891 Then
                        If ePlateType = DeckPlate Then
                                If ePlateTightness = NonTight Then
                                        strPartName = "C-80"
                                        strPartClassName = "CableCoamingCType"
                                Else
                                        strPartName = "K-80"
                                        strPartClassName = "CableCoamingKType"
                                End If
                        ElseIf ePlateType = LBulkheadPlate Or ePlateType = TBulkheadPlate Then
                                strPartName = "CA-80"
                                strPartClassName = "CableCoamingCAType"
                        End If
                End If
            End If
  
'******************************************************************************
' Conduit
'******************************************************************************
    
        ElseIf TypeOf oObject Is IJRteConduitPathFeat Then
            Dim dNominalDia As Double
        
            'Get the Nominal Dia of Conduit
            dNominalDia = GetNominalDiaofConduit(oHoleTrace)
        
            'Decide the Part Name based on Size
            If dNominalDia >= 15# And dNominalDia <= 20# Then
                strPartName = "RS50"
            ElseIf dNominalDia >= 25# And dNominalDia <= 32# Then
                strPartName = "RS75"
            ElseIf dNominalDia >= 40# And dNominalDia <= 50# Then
                strPartName = "RS100"
            ElseIf dNominalDia >= 65# And dNominalDia <= 80 Then
                strPartName = "RS125"
            ElseIf dNominalDia >= 90# And dNominalDia <= 100# Then
                strPartName = "RS150"
                'None of the cases above set the Fitting type to HM_FittingNone
            Else
                oHoleTrace.FittingType = HM_FittingNone
                Exit Sub
            End If
            strPartClassName = "ConduitSealType"
        End If
    
        Exit For
    Next
    
Cleanup:
    Set oObject = Nothing
    Set oOutfitting = Nothing
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, Module & " - " & METHOD, Err.Description
    GoTo Cleanup
End Sub

'//////////////////////////////////////////////////////////////////////////////
'******************************************************************************
'Method: CreateFitting
'
'Interface: IJHMFittingSelRule
'
'Abstract: create the actual fitting based on the fitting type
'          of the hole trace
'
'   Input:
'       oHoleTrace - hole trace
'
'   Output:
'       oFitting - the fitting created
'
'******************************************************************************
Private Sub IJHMFittingSelRule_CreateFitting(ByVal oHoleTrace As IJHoleTraceAE, _
                                             ByRef oFitting As Object)
    Const METHOD = "IJHMFittingSelRule_CreateFitting"
    On Error GoTo ErrorHandler

    ' need to get the fitting active entity from the hole trace
    Dim oFittingAE As IJHoleFitting
    On Error Resume Next
    Set oFittingAE = oHoleTrace.GetHoleFitting
    On Error GoTo ErrorHandler
    
    ' if it wasn't there - return now
    If oFittingAE Is Nothing Then
        GoTo Cleanup
    End If
    
    'get the resource manager
    Dim oResourceMgr As Object
    Set oResourceMgr = GetResourceMgrFromObject(oHoleTrace)
    If oResourceMgr Is Nothing Then GoTo Cleanup
    
    ' need a wrapper for the hole trace - it will do all the work
'    Dim oWrapHole As HoleMgmtObjects.HoleTrace
    Dim oWrapHole As Object
    Set oWrapHole = SP3DCreateObject("HoleMgmtObjects.HoleTrace")
    Set oWrapHole.object = oHoleTrace
    
    ' create the fitting based on fitting type
    Dim eFittingType As HMFittingType
    eFittingType = oFittingAE.FittingType
    
    ' also need to know the application
    Dim eRequestingApp As HMHoleRequestApp
    eRequestingApp = oHoleTrace.RequestingApp
    
    Dim oER As Object
    Dim oXSect As IJCrossSection
    
    Dim oMatrix As IJDT4x4
    Dim oPartOcc As IJPartOcc
    Dim oRules As IJElements
    Dim oNameRuleAE As IJNameRuleAE
    Dim oNamingObject As IJDNamingRulesHelper
    
    Dim oHMConnection As Object
    
    ' create the actual fitting
    Select Case eFittingType

'******************************************************************************
' No fitting selected
'******************************************************************************
        
        Case HM_FittingNone
            GoTo Cleanup
        
'******************************************************************************
' Coaming
'******************************************************************************
        
        Case HM_FittingCoaming
            ' create a cross section to use - example of FB cross section type
            Set oXSect = CreateCrossSection("FB")
            If oXSect Is Nothing Then
                GoTo Cleanup
            End If
            
            ' create the edge reinforcement
            Set oER = oWrapHole.PlaceEdgeReinforcement(oXSect)
            If Not oER Is Nothing Then
                Set oFitting = oER
            End If
            
'******************************************************************************
' Sloped Coaming - currently same as Coaming
'******************************************************************************
        
        Case HM_FittingCoamingSloped
            ' create a cross section to use - example of FB cross section type
            Set oXSect = CreateCrossSection("FB")
            If oXSect Is Nothing Then
                GoTo Cleanup
            End If
            
            ' create the edge reinforcement
            Set oER = oWrapHole.PlaceEdgeReinforcement(oXSect)
            If Not oER Is Nothing Then
                Set oFitting = oER
            End If
        
'******************************************************************************
' Penetrated Coaming
'******************************************************************************
        
        Case HM_FittingPenetrated
            ' create a cross section to use - example of FB cross section type
            Set oXSect = CreateCrossSection("FB")
            If oXSect Is Nothing Then
                GoTo Cleanup
            End If
            
            ' create the edge reinforcement
            Set oER = oWrapHole.PlaceEdgeReinforcement(oXSect)
            If Not oER Is Nothing Then
                Set oFitting = oER
            End If
        
'******************************************************************************
' Sloped Penetrated Coaming - currently same as Penetrated Coaming
'******************************************************************************
        
        Case HM_FittingPenetratedSloped
            ' create a cross section to use - example of FB cross section type
            Set oXSect = CreateCrossSection("FB")
            If oXSect Is Nothing Then
                GoTo Cleanup
            End If
            
            ' create the edge reinforcement
            Set oER = oWrapHole.PlaceEdgeReinforcement(oXSect)
            If Not oER Is Nothing Then
                Set oFitting = oER
            End If
        
'******************************************************************************
' Center Flange
'******************************************************************************
        
        Case HM_FittingCenterFlange
            ' create edge reinforcement plate part
            Set oER = oWrapHole.PlacePlateFitting()
            If Not oER Is Nothing Then
                Set oFitting = oER
            End If
    
            ' if piping app - create the connection for pipe spooling
            If eRequestingApp = HM_PipingApp Then
                'going to pass the call on to the fitting active entity
                Set oHMConnection = oFittingAE.ConstructHMConnection(oResourceMgr, oFitting)
            End If
        
'******************************************************************************
' Double Ring
'******************************************************************************
        
        Case HM_FittingDoubleRing
            ' create edge reinforcement plate part
            Set oER = oWrapHole.PlacePlateFitting()
            If Not oER Is Nothing Then
                Set oFitting = oER
            End If
        
'******************************************************************************
' Wall Attached
'******************************************************************************
        
        Case HM_FittingWallAttached
            ' create a cross section to use - example of FB cross section type
            Set oXSect = CreateCrossSection("FB")
            If oXSect Is Nothing Then
                GoTo Cleanup
            End If
    
            ' create the edge reinforcement
            Set oER = oWrapHole.PlaceProfileFitting(oXSect)
            If Not oER Is Nothing Then
                Set oFitting = oER
            End If
        
'******************************************************************************
' Catalog Part
'******************************************************************************
        
        Case HM_FittingCatalogPart
            ' create the part occurrence
            Set oPartOcc = oWrapHole.PlaceCatalogFitting()
            If oPartOcc Is Nothing Then
                GoTo Cleanup
            End If
    
            ' if piping app with a Sleeve fitting
            If eRequestingApp = HM_PipingApp Then
                ' Assuming all PipingApps using Catalog Parts are Pipe Sleeves
                Set oMatrix = oWrapHole.HoleMatrix
                AdjustFittingToAlongPipeSleeve oHoleTrace, oPartOcc, oMatrix
                oPartOcc.SetMatrix oMatrix
            End If
    
            ' retrieve first default naming rule
            Set oNamingObject = New NamingRulesHelper
            oNamingObject.GetEntityNamingRulesGivenName "CPEquipment", oRules
            If oRules.Count >= 1 Then
                oNamingObject.AddNamingRelations oPartOcc, oRules.Item(1), oNameRuleAE
            End If
    
            ' send the fitting back
            Set oFitting = oPartOcc
    
            ' if piping app - create the connection for pipe spooling
            If eRequestingApp = HM_PipingApp Then
                'going to pass the call on to the fitting active entity
                Set oHMConnection = oFittingAE.ConstructHMConnection(oResourceMgr, oFitting)
            End If
    
    End Select
    
Cleanup:
    Set oER = Nothing
    Set oXSect = Nothing
    
    Set oPartOcc = Nothing
    Set oRules = Nothing
    Set oNameRuleAE = Nothing
    Set oNamingObject = Nothing
    
    Set oWrapHole = Nothing
    Set oFittingAE = Nothing
    
    Set oResourceMgr = Nothing
    Set oHMConnection = Nothing
    Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, Module & " - " & METHOD, Err.Description
    GoTo Cleanup
End Sub

'//////////////////////////////////////////////////////////////////////////////
'******************************************************************************
'Method: UpdateFitting
'
'Interface: IJHMFittingSelRule
'
'Abstract: the hole trace has been updated - allow the actual fitting
'          to update as well
'
'   Input:
'       oHoleTrace - hole trace
'
'   Output:
'       oFitting - the update fitting
'
'******************************************************************************
Private Sub IJHMFittingSelRule_UpdateFitting(ByVal oHoleTrace As IJHoleTraceAE, _
                                             ByRef oFitting As Object)
    Const METHOD = "IJHMFittingSelRule_UpdateFitting"
    On Error GoTo ErrorHandler

    ' need to get the fitting active entity from the hole trace
    Dim oFittingAE As IJHoleFitting
    On Error Resume Next
    Set oFittingAE = oHoleTrace.GetHoleFitting
    On Error GoTo ErrorHandler
    
    ' if it wasn't there - return now
    If oFittingAE Is Nothing Then
        GoTo Cleanup
    End If
    
    ' get the current fitting
    On Error Resume Next
    Set oFitting = oFittingAE.HoleFitting
    On Error GoTo ErrorHandler
    
    ' if it wasn't there - return
    If oFitting Is Nothing Then
        GoTo Cleanup
    End If
    
    'get the resource manager
    Dim oResourceMgr As Object
    Set oResourceMgr = GetResourceMgrFromObject(oHoleTrace)
    If oResourceMgr Is Nothing Then GoTo Cleanup
    
    ' need a wrapper for the hole trace - it will do all the work
'    Dim oWrapHole As HoleMgmtObjects.HoleTrace
'    Set oWrapHole = New HoleMgmtObjects.HoleTrace
'
    Dim oWrapHole As Object
    Set oWrapHole = SP3DCreateObject("HoleMgmtObjects.HoleTrace")
    Set oWrapHole.object = oHoleTrace
    
    ' going to need a middle helper
    Dim oHMHelper As IJHMMiddleHelper
    Set oHMHelper = New CHMMiddleHelper
    
    ' update the fitting based on fitting type
    Dim eFittingType As HMFittingType
    eFittingType = oFittingAE.FittingType
    
    ' also need to know the application
    Dim eRequestingApp As HMHoleRequestApp
    eRequestingApp = oHoleTrace.RequestingApp
    
    Dim oPenetratedPlate As IJPlate
    Dim oFittingPlate As IJPlate
    Dim oDimensions As IJDPlateDimensions
    Dim oPenetratedMaterial As IJStructureMaterial
    Dim oFittingMaterial As IJStructureMaterial
    Dim oPlateMaterial As Object
    
    Dim oInnerCuts As IJDObjectCollection
    
    Dim oPartOcc As IJPartOcc
    Dim oMatrix As IJDT4x4
                
    ' update the actual fitting
    Select Case eFittingType

'******************************************************************************
' No fitting selected
'******************************************************************************
        
        Case HM_FittingNone
            GoTo Cleanup
        
'******************************************************************************
' Coaming
'******************************************************************************
        
        Case HM_FittingCoaming
            ' nothing to be updated at this time ...
            GoTo Cleanup
            
'******************************************************************************
' Sloped Coaming - currently same as Coaming
'******************************************************************************
        
        Case HM_FittingCoamingSloped
            ' nothing to be updated at this time ...
            GoTo Cleanup
        
'******************************************************************************
' Penetrated Coaming
'******************************************************************************
        
        Case HM_FittingPenetrated
            ' nothing to be updated at this time ...
            GoTo Cleanup
        
'******************************************************************************
' Sloped Penetrated Coaming - currently same as Penetrated Coaming
'******************************************************************************
        
        Case HM_FittingPenetratedSloped
            ' nothing to be updated at this time ...
            GoTo Cleanup
        
'******************************************************************************
' Center Flange
'******************************************************************************
        
        Case HM_FittingCenterFlange
            ' get the cutting plate part and copy the dimension and material
            Set oPenetratedPlate = oWrapHole.CuttingPlatePart
            Set oFittingPlate = oFitting

            ' copy the dimensions onto the fitting plate
            If Not oPenetratedPlate Is Nothing Then
                Set oDimensions = oPenetratedPlate.Dimensions
                oFittingPlate.Dimensions = oDimensions
            End If
            
            ' copy the material
            Set oFittingMaterial = oFittingPlate
            Set oPenetratedMaterial = oPenetratedPlate
        
            If Not oPenetratedMaterial Is Nothing Then
                Set oPlateMaterial = oPenetratedMaterial.Material
                oFittingMaterial.Material = oPlateMaterial
            End If
        
            ' update the cuts on the center flange
            Set oInnerCuts = oWrapHole.CreateCenterFlangeInnerCuts(oResourceMgr)
            oHMHelper.CutCenterFlange oResourceMgr, _
                                      oFittingPlate, _
                                      oInnerCuts, _
                                      oHoleTrace, _
                                      oPenetratedPlate
        
            ' also need to update the spooling connection
            oFittingAE.UpdateHMConnection

'******************************************************************************
' Double Ring
'******************************************************************************
        
        Case HM_FittingDoubleRing
            ' get the cutting plate part and copy the dimension and material
            Set oPenetratedPlate = oWrapHole.CuttingPlatePart
            Set oFittingPlate = oFitting

            ' copy the dimensions onto the fitting plate
            If Not oPenetratedPlate Is Nothing Then
                Set oDimensions = oPenetratedPlate.Dimensions
                oFittingPlate.Dimensions = oDimensions
            End If
            
            ' copy the material
            Set oFittingMaterial = oFittingPlate
            Set oPenetratedMaterial = oPenetratedPlate
        
            If Not oPenetratedMaterial Is Nothing Then
                Set oPlateMaterial = oPenetratedMaterial.Material
                oFittingMaterial.Material = oPlateMaterial
            End If
        
'******************************************************************************
' Wall Attached
'******************************************************************************
        
        Case HM_FittingWallAttached
            ' nothing to be updated at this time ...
            GoTo Cleanup
        
'******************************************************************************
' Catalog Part
'******************************************************************************
        
        Case HM_FittingCatalogPart
            If TypeOf oFitting Is IJPartOcc Then
                Set oPartOcc = oFitting
                
                ' did the user select a different catalog part ...
                Dim oOldPart As IJDPart
                oPartOcc.GetPart oOldPart
                
                Dim oNewPart As IJDPart
                Set oNewPart = oFittingAE.FittingPart
                
                If Not oOldPart Is oNewPart Then
                    oPartOcc.SetPart2 oNewPart, True
                End If
                
                ' need to update the matrix of the part occurrence
                Set oMatrix = oWrapHole.HoleMatrix

                ' if piping app with a Sleeve fitting
                If eRequestingApp = HM_PipingApp Then
                    ' Assuming all PipingApps using Catalog Parts are Pipe Sleeves
                    AdjustFittingToAlongPipeSleeve oHoleTrace, oPartOcc, oMatrix
                End If

                ' Rotate the matrix if needed
                Dim dAngle As Double
                Dim oVector As DVector
                
                dAngle = oWrapHole.OrientationAngle
                If dAngle > 0 Then
                    Set oVector = New DVector
                    oVector.Set 0, 0, 1
                    oMatrix.Rotate dAngle, oVector
                   Set oVector = Nothing
                End If

                oPartOcc.SetMatrix oMatrix
            End If

            ' also need to update the spooling connection
            oFittingAE.UpdateHMConnection
    
    End Select
    
Cleanup:
    Set oPartOcc = Nothing
    Set oMatrix = Nothing
    
    Set oHMHelper = Nothing
    Set oResourceMgr = Nothing
    
    Set oDimensions = Nothing
    Set oPenetratedPlate = Nothing
    Set oFittingPlate = Nothing
    Set oPenetratedMaterial = Nothing
    Set oFittingMaterial = Nothing
    Set oPlateMaterial = Nothing

    Set oInnerCuts = Nothing
    
    Set oWrapHole = Nothing
    Set oFittingAE = Nothing
    Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, Module & " - " & METHOD, Err.Description
    GoTo Cleanup
End Sub

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: FlipFitting
'
'Interface: IJHMFittingSelRule
'
'Abstract: flip the fitting to the other side of the structure
'
'   Input:
'       oHoleTrace - hole trace
'
'   Output:
'       oFitting - the update fitting
'
'********************************************************************
Private Sub IJHMFittingSelRule_FlipFitting(ByVal oHoleTrace As IJHoleTraceAE, _
                                           ByRef oFitting As Object)
    Const METHOD = "IJHMFittingSelRule_FlipFitting"
    On Error GoTo ErrorHandler

    ' need to get the fitting active entity from the hole trace
    Dim oFittingAE As IJHoleFitting
    On Error Resume Next
    Set oFittingAE = oHoleTrace.GetHoleFitting
    On Error GoTo ErrorHandler
    
    ' if it wasn't there - return now
    If oFittingAE Is Nothing Then
        GoTo Cleanup
    End If
    
    ' get the current fitting
    On Error Resume Next
    Set oFitting = oFittingAE.HoleFitting
    On Error GoTo ErrorHandler
    
    ' if it wasn't there - return
    If oFitting Is Nothing Then
        GoTo Cleanup
    End If

    ' flip the fitting based on fitting type
    Dim eFittingType As HMFittingType
    eFittingType = oFittingAE.FittingType

    Dim oPartOcc As IJPartOcc
    Dim oMatrix As IJDT4x4
    
    ' update the actual fitting
    Select Case eFittingType

'******************************************************************************
' No fitting selected
'******************************************************************************
        
        Case HM_FittingNone
            GoTo Cleanup
        
'******************************************************************************
' Coaming
'******************************************************************************
        
        Case HM_FittingCoaming
            FlipProfileFitting oFitting
            
'******************************************************************************
' Sloped Coaming - currently same as Coaming
'******************************************************************************
        
        Case HM_FittingCoamingSloped
            FlipProfileFitting oFitting
        
'******************************************************************************
' Penetrated Coaming
'******************************************************************************
        
        Case HM_FittingPenetrated
            FlipProfileFitting oFitting
        
'******************************************************************************
' Sloped Penetrated Coaming - currently same as Penetrated Coaming
'******************************************************************************
        
        Case HM_FittingPenetratedSloped
            FlipProfileFitting oFitting
        
'******************************************************************************
' Center Flange
'******************************************************************************
        
        Case HM_FittingCenterFlange
            FlipPlateFitting oFitting
        
'******************************************************************************
' Double Ring
'******************************************************************************
        
        Case HM_FittingDoubleRing
            FlipPlateFitting oFitting
        
'******************************************************************************
' Wall Attached
'******************************************************************************
        
        Case HM_FittingWallAttached
            FlipProfileFitting oFitting
        
'******************************************************************************
' Catalog Part
'******************************************************************************
        
        Case HM_FittingCatalogPart
            If TypeOf oFitting Is IJPartOcc Then
                Set oPartOcc = oFitting
                oPartOcc.GetMatrix oMatrix
                
                Dim values(16) As Double
                oMatrix.Get values(0)

                values(0) = 0# - values(0)
                values(1) = 0# - values(1)
                values(2) = 0# - values(2)

                values(8) = 0# - values(8)
                values(9) = 0# - values(9)
                values(10) = 0# - values(10)

                oMatrix.Set values(0)
                oPartOcc.SetMatrix oMatrix
            End If
    
    End Select

Cleanup:
    Set oPartOcc = Nothing
    Set oMatrix = Nothing
    
    Set oFittingAE = Nothing
    Set oFitting = Nothing
    
    Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, Module & " - " & METHOD, Err.Description
    GoTo Cleanup
End Sub

'//////////////////////////////////////////////////////////////////////////////
'******************************************************************************
' Private methods follow
'******************************************************************************

'******************************************************************************
' METHOD:  CreateCrossSection
'
' DESCRIPTION:  create a cross section to be applied to the profile
'******************************************************************************
Private Function CreateCrossSection(ByVal strName As String) As IJCrossSection
    On Error GoTo ErrorHandler
    Const METHOD = "CreateCrossSection"
    
    Dim oSRDQuery As IJSRDStructQuery
    Dim oXSectColl As IJDCollection
    Dim oXSect As IJCrossSection
    
    Set oSRDQuery = New SRDQuery
    Set oXSectColl = oSRDQuery.GetSections("ShipShapes", strName)
    If oXSectColl.Size > 0 Then
        Set oXSect = oXSectColl.Item(1)
        If Not oXSect Is Nothing Then
            Set CreateCrossSection = oXSect
        End If
    End If
    
Cleanup:
    Set oSRDQuery = Nothing
    Set oXSectColl = Nothing
    Set oXSect = Nothing
    
    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, Module & " - " & METHOD, Err.Description
    GoTo Cleanup
End Function

'******************************************************************************
' METHOD:  FlipPlateFitting
'
' DESCRIPTION:  flip a fitting made from a stand-alone plate part
'******************************************************************************
Private Sub FlipPlateFitting(oFitting As Object)
    On Error GoTo ErrorHandler
    Const METHOD = "FlipPlateFitting"

    Dim oPlate As IJDStandAlonePlatePart
    Set oPlate = oFitting
    
    Dim thicknessDir As StructMoldedDirection
    thicknessDir = oPlate.plateThicknessDirection

    Select Case thicknessDir
        Case PortDir
            oPlate.plateThicknessDirection = StarDir
        
        Case StarDir
            oPlate.plateThicknessDirection = PortDir
        
        Case ForeDir
            oPlate.plateThicknessDirection = AftDir
        
        Case AftDir
            oPlate.plateThicknessDirection = ForeDir
        
        Case AboveDir
            oPlate.plateThicknessDirection = BelowDir
        
        Case BelowDir
            oPlate.plateThicknessDirection = AboveDir
        
        Case InboardDir
            oPlate.plateThicknessDirection = OutboardDir
        
        Case OutboardDir
            oPlate.plateThicknessDirection = InboardDir
    
    End Select
    
Cleanup:
    Set oPlate = Nothing
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, Module & " - " & METHOD, Err.Description
    GoTo Cleanup
End Sub

'******************************************************************************
' METHOD:  FlipProfileFitting
'
' DESCRIPTION:  flip a fitting made from a stand-alone profile part
'******************************************************************************
Private Sub FlipProfileFitting(oFitting As Object)
    On Error GoTo ErrorHandler
    Const METHOD = "FlipProfileFitting"

    Dim oConventions As IJDProfileMoldedConventions
    Set oConventions = oFitting

    Dim orientation As StructMoldedOrientation
    orientation = oConventions.orientation

    Select Case orientation
        Case PortOrient
            oConventions.orientation = StarboardOrient
        
        Case StarboardOrient
            oConventions.orientation = PortOrient
        
        Case ForeOrient
            oConventions.orientation = AftOrient
        
        Case AftOrient
            oConventions.orientation = ForeOrient
        
        Case AboveOrient
            oConventions.orientation = BelowOrient
        
        Case BelowOrient
            oConventions.orientation = AboveOrient
        
        Case InboardOrient
            oConventions.orientation = OutboardOrient
        
        Case OutboardOrient
            oConventions.orientation = InboardOrient
    
    End Select

Cleanup:
    Set oConventions = Nothing
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, Module & " - " & METHOD, Err.Description
    GoTo Cleanup
End Sub
 
