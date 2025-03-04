VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CablewayHoleSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
' Copyright (C) 1998-2002 Intergraph Corporation. All Rights Reserved.
'
' Project: S:\HoleMgmt\Data\SmartOccurrence\HoleRules
'
' File: CablewayHoleSel.cls
'
' Author: Hole Mgmt Team
'
' Abstract: selector for cableway hole trace smart class
'
' Description: This is the selector for the cableway hole trace smart class.
'              This class will be chosen if the outfitting for the hole trace
'              is a single cableway.
'
' Changes:
'   Date    By    Reason
' 05/15/03  CRS  Modified AddValidCWType to add cableway types that are close
'                in size to the one selected rather than adding all cableway
'                types.
'******************************************************************************

Option Explicit

Const m_SelectorProgid As String = "HoleRules.CablewayHoleSel"
Const m_SelectorName As String = "HoleRules.CablewayHoleSel"
Const m_FamilyProgid As String = ""

Private sERROR As String
Private sMETHOD As String
Private Const MODULE = "M:\SharedContent\Src\HoleMgmt\SmartOccurrence\HoleRules\CablewayHoleSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    'the hole trace is the only input
    pIH.SetInput INPUT_HOLETRACE

    Exit Sub

ErrorHandler:
    pIH.ReportError
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
 
    'none at this time
   
    Exit Sub

ErrorHandler:
    pQH.ReportError
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
    On Error GoTo ErrorHandler
 
'    Const PI = 3.14159265
    
    'Get the HoleSmartOcc and HoleTraceAE
    Dim oHoleSmartOcc As IJHoleSmartOcc
    Set oHoleSmartOcc = pSLH.SmartOccurrence
    Dim oHoleTrace As IJHoleTraceAE
    Set oHoleTrace = oHoleSmartOcc.GetHoleTraceAE
        
    'Get the Outfitting, Only Cable Way should be taken.
    Dim eSelectedType As CrossSectionShapeTypes, dWidth As Double, dDepth As Double
    GetWidthDepthOfCableway oHoleTrace, eSelectedType, dWidth, dDepth
    
    'Check the CrossSection type of CW and check the width and Depth
    Dim strCrossSection As String
    
    If eSelectedType = CrossSectionShapeTypes.Rectangular Then    'Rectangular = 1
        If dWidth > 0 And dWidth <= 0.2 Then
            If dDepth > 0 And dDepth <= 0.15 Then
                strCrossSection = "Rect284x124"
            ElseIf dDepth > 0.15 And dDepth <= 0.25 Then
                strCrossSection = "Rect284x224"
            ElseIf dDepth > 0.25 And dDepth <= 0.4 Then
                strCrossSection = "Rect284x374"
            ElseIf dDepth > 0.4 And dDepth <= 0.55 Then
                strCrossSection = "Rect284x524"
            ElseIf dDepth > 0.55 And dDepth <= 0.7 Then
                strCrossSection = "Rect284x674"
            Else
                strCrossSection = "Rect284x124"
                UpdateStatusBar "No support found for selected Cableway, set the support to Rect284x124"
            End If
        
        ElseIf dWidth > 0.2 And dWidth <= 0.3 Then
            If dDepth > 0 And dDepth <= 0.15 Then
                strCrossSection = "Rect384x124"
            ElseIf dDepth > 0.15 And dDepth <= 0.25 Then
                strCrossSection = "Rect384x224"
            ElseIf dDepth > 0.25 And dDepth <= 0.4 Then
                strCrossSection = "Rect384x374"
            ElseIf dDepth > 0.4 And dDepth <= 0.55 Then
                strCrossSection = "Rect384x524"
            ElseIf dDepth > 0.55 And dDepth <= 0.7 Then
                strCrossSection = "Rect384x674"
            Else
                strCrossSection = "Rect384x124"
                UpdateStatusBar "No support found for selected Cableway, set the support to Rect384x124"
            End If
        
        ElseIf dWidth > 0.3 And dWidth <= 0.4 Then
            If dDepth > 0 And dDepth <= 0.15 Then
                strCrossSection = "Rect484x124"
            ElseIf dDepth > 0.15 And dDepth <= 0.25 Then
                strCrossSection = "Rect484x224"
            ElseIf dDepth > 0.25 And dDepth <= 0.4 Then
                strCrossSection = "Rect484x374"
            ElseIf dDepth > 0.4 And dDepth <= 0.55 Then
                strCrossSection = "Rect484x524"
            ElseIf dDepth > 0.55 And dDepth <= 0.7 Then
                strCrossSection = "Rect484x674"
            Else
                strCrossSection = "Rect484x124"
                UpdateStatusBar "No support found for selected Cableway, set the support to Rect484x124"
            End If
        
        ElseIf dWidth > 0.4 And dWidth <= 0.5 Then
            If dDepth > 0 And dDepth <= 0.15 Then
                strCrossSection = "Rect584x124"
            ElseIf dDepth > 0.15 And dDepth <= 0.25 Then
                strCrossSection = "Rect584x224"
            ElseIf dDepth > 0.25 And dDepth <= 0.4 Then
                strCrossSection = "Rect584x374"
            ElseIf dDepth > 0.4 And dDepth <= 0.55 Then
                strCrossSection = "Rect584x524"
            ElseIf dDepth > 0.55 And dDepth <= 0.7 Then
                strCrossSection = "Rect584x674"
            Else
                strCrossSection = "Rect584x124"
                UpdateStatusBar "No support found for selected Cableway, set the support to Rect584x124"
            End If
       
        ElseIf dWidth > 0.5 And dWidth <= 0.6 Then
            If dDepth > 0 And dDepth <= 0.15 Then
                strCrossSection = "Rect684x124"
            ElseIf dDepth > 0.15 And dDepth <= 0.25 Then
                strCrossSection = "Rect684x224"
            ElseIf dDepth > 0.25 And dDepth <= 0.4 Then
                strCrossSection = "Rect684x374"
            ElseIf dDepth > 0.4 And dDepth <= 0.55 Then
                strCrossSection = "Rect684x524"
            ElseIf dDepth > 0.55 And dDepth <= 0.7 Then
                strCrossSection = "Rect684x674"
            Else
                strCrossSection = "Rect684x124"
                UpdateStatusBar "No support found for selected Cableway, set the support to Rect684x124"
            End If
        
        ElseIf dWidth > 0.6 And dWidth <= 0.7 Then
            If dDepth > 0 And dDepth <= 0.15 Then
                strCrossSection = "Rect784x124"
            ElseIf dDepth > 0.15 And dDepth <= 0.25 Then
                strCrossSection = "Rect784x224"
            ElseIf dDepth > 0.25 And dDepth <= 0.4 Then
                strCrossSection = "Rect784x374"
            ElseIf dDepth > 0.4 And dDepth <= 0.55 Then
                strCrossSection = "Rect784x524"
            ElseIf dDepth > 0.55 And dDepth <= 0.7 Then
                strCrossSection = "Rect784x674"
            Else
                strCrossSection = "Rect784x124"
                UpdateStatusBar "No support found for selected Cableway, set the support to Rect784x124"
            End If
        
        ElseIf dWidth > 0.7 And dWidth <= 0.8 Then
            If dDepth > 0 And dDepth <= 0.15 Then
                strCrossSection = "Rect884x124"
            ElseIf dDepth > 0.15 And dDepth <= 0.25 Then
                strCrossSection = "Rect884x224"
            ElseIf dDepth > 0.25 And dDepth <= 0.4 Then
                strCrossSection = "Rect884x374"
            ElseIf dDepth > 0.4 And dDepth <= 0.55 Then
                strCrossSection = "Rect884x524"
            ElseIf dDepth > 0.55 And dDepth <= 0.7 Then
                strCrossSection = "Rect884x674"
            Else
                strCrossSection = "Rect884x124"
                UpdateStatusBar "No support found for selected Cableway, set the support to Rect884x124"
            End If
  
        ElseIf dWidth > 0.8 And dWidth <= 0.9 Then
            If dDepth > 0 And dDepth <= 0.15 Then
                strCrossSection = "Rect984x124"
            ElseIf dDepth > 0.15 And dDepth <= 0.25 Then
                strCrossSection = "Rect984x224"
            ElseIf dDepth > 0.25 And dDepth <= 0.4 Then
                strCrossSection = "Rect984x374"
            ElseIf dDepth > 0.4 And dDepth <= 0.55 Then
                strCrossSection = "Rect984x524"
            ElseIf dDepth > 0.55 And dDepth <= 0.7 Then
                strCrossSection = "Rect984x674"
            Else
                strCrossSection = "Rect984x124"
                UpdateStatusBar "No support found for selected Cableway, set the support to Rect984x124"
            End If
  
        ElseIf dWidth > 0.9 And dWidth <= 1 Then
            If dDepth > 0 And dDepth <= 0.15 Then
                strCrossSection = "Rect1084x124"
            ElseIf dDepth > 0.15 And dDepth <= 0.25 Then
                strCrossSection = "Rect1084x224"
            ElseIf dDepth > 0.25 And dDepth <= 0.4 Then
                strCrossSection = "Rect1084x374"
            ElseIf dDepth > 0.4 And dDepth <= 0.55 Then
                strCrossSection = "Rect1084x524"
            ElseIf dDepth > 0.55 And dDepth <= 0.7 Then
                strCrossSection = "Rect1084x674"
            Else
                strCrossSection = "Rect1084x124"
                UpdateStatusBar "No support found for selected Cableway, set the support to Rect1084x124"
            End If
        
        Else
            strCrossSection = "Rect284x124"
            UpdateStatusBar "No support found for selected Cableway, set the support to Rect284x124"
        End If
        
    ElseIf eSelectedType = CrossSectionShapeTypes.Round Then
        If dWidth = 0.06 Then
            strCrossSection = "Round62"
        ElseIf dWidth = 0.089 Then
            strCrossSection = "Round92"
        Else
            strCrossSection = "Round62"
            UpdateStatusBar "No support found for selected Cableway, set the support to Round62"
        End If
    
    ElseIf eSelectedType = CrossSectionShapeTypes.FlatOval Then
     
        Dim dClearance As Double
        dClearance = 2# * 0.012
        
        ' Consider only cableway height less than that which will fit
        ' the FlatOval symbols already specified in the catalog.  i.e.
        ' depth must be less than 0.124 - 2 * dClearance.
        If dDepth > 0# And dDepth <= 0.124 - dClearance Then
            If dWidth > 0# And dWidth <= 0.284 - dClearance Then
                strCrossSection = "FlatOval284x124"
            ElseIf dWidth > 0.284 - dClearance And dWidth <= 0.384 - dClearance Then
                strCrossSection = "FlatOval384x124"
            ElseIf dWidth > 0.384 - dClearance And dWidth <= 0.484 - dClearance Then
                strCrossSection = "FlatOval484x124"
            ElseIf dWidth > 0.484 - dClearance And dWidth <= 0.584 - dClearance Then
                strCrossSection = "FlatOval584x124"
            ElseIf dWidth > 0.584 - dClearance And dWidth <= 0.684 - dClearance Then
                strCrossSection = "FlatOval684x124"
            ElseIf dWidth > 0.684 - dClearance And dWidth <= 0.784 - dClearance Then
                strCrossSection = "FlatOval784x124"
            ElseIf dWidth > 0.784 - dClearance And dWidth <= 0.884 - dClearance Then
                strCrossSection = "FlatOval884x124"
            ElseIf dWidth > 0.884 - dClearance And dWidth <= 0.984 - dClearance Then
                strCrossSection = "FlatOval984x124"
            ElseIf dWidth > 0.984 - dClearance And dWidth <= 1.084 - dClearance Then
                strCrossSection = "FlatOval1084x124"
            Else
                strCrossSection = "FlatOval284x124"
                UpdateStatusBar "No support found for selected Cableway, set the support to FlatOval284x124"
            End If
            
        ElseIf dDepth > 0.124 - dClearance And dDepth <= 0.284 - dClearance Then
            If dWidth > 0# And dWidth <= 0.284 - dClearance Then
                strCrossSection = "FlatOval284x284"
            ElseIf dWidth > 0.284 - dClearance And dWidth <= 0.384 - dClearance Then
                strCrossSection = "FlatOval384x284"
            ElseIf dWidth > 0.384 - dClearance And dWidth <= 0.484 - dClearance Then
                strCrossSection = "FlatOval484x284"
            ElseIf dWidth > 0.484 - dClearance And dWidth <= 0.584 - dClearance Then
                strCrossSection = "FlatOval584x284"
            ElseIf dWidth > 0.584 - dClearance And dWidth <= 0.684 - dClearance Then
                strCrossSection = "FlatOval684x284"
            ElseIf dWidth > 0.684 - dClearance And dWidth <= 0.784 - dClearance Then
                strCrossSection = "FlatOval784x284"
            ElseIf dWidth > 0.784 - dClearance And dWidth <= 0.884 - dClearance Then
                strCrossSection = "FlatOval884x284"
            ElseIf dWidth > 0.884 - dClearance And dWidth <= 0.984 - dClearance Then
                strCrossSection = "FlatOval984x284"
            ElseIf dWidth > 0.984 - dClearance And dWidth <= 1.084 - dClearance Then
                strCrossSection = "FlatOval1084x284"
            Else
                strCrossSection = "FlatOval284x284"
                UpdateStatusBar "No support found for selected Cableway, set the support to FlatOval284x284"
            End If
        ElseIf dDepth > 0.284 - dClearance And dDepth <= 0.384 - dClearance Then
            If dWidth > 0# And dWidth <= 0.284 - dClearance Then
                strCrossSection = "FlatOval284x384"
            ElseIf dWidth > 0.284 - dClearance And dWidth <= 0.384 - dClearance Then
                strCrossSection = "FlatOval384x384"
            ElseIf dWidth > 0.384 - dClearance And dWidth <= 0.484 - dClearance Then
                strCrossSection = "FlatOval484x384"
            ElseIf dWidth > 0.484 - dClearance And dWidth <= 0.584 - dClearance Then
                strCrossSection = "FlatOval584x384"
            ElseIf dWidth > 0.584 - dClearance And dWidth <= 0.684 - dClearance Then
                strCrossSection = "FlatOval684x384"
            ElseIf dWidth > 0.684 - dClearance And dWidth <= 0.784 - dClearance Then
                strCrossSection = "FlatOval784x384"
            ElseIf dWidth > 0.784 - dClearance And dWidth <= 0.884 - dClearance Then
                strCrossSection = "FlatOval884x384"
            ElseIf dWidth > 0.884 - dClearance And dWidth <= 0.984 - dClearance Then
                strCrossSection = "FlatOval984x384"
            ElseIf dWidth > 0.984 - dClearance And dWidth <= 1.084 - dClearance Then
                strCrossSection = "FlatOval1084x384"
            Else
                strCrossSection = "FlatOval284x384"
                UpdateStatusBar "No support found for selected Cableway, set the support to FlatOval284x384"
            End If
        ElseIf dDepth > 0.384 - dClearance And dDepth <= 0.484 - dClearance Then
            If dWidth > 0# And dWidth <= 0.284 - dClearance Then
                strCrossSection = "FlatOval284x484"
            ElseIf dWidth > 0.284 - dClearance And dWidth <= 0.384 - dClearance Then
                strCrossSection = "FlatOval384x484"
            ElseIf dWidth > 0.384 - dClearance And dWidth <= 0.484 - dClearance Then
                strCrossSection = "FlatOval484x484"
            ElseIf dWidth > 0.484 - dClearance And dWidth <= 0.584 - dClearance Then
                strCrossSection = "FlatOval584x484"
            ElseIf dWidth > 0.584 - dClearance And dWidth <= 0.684 - dClearance Then
                strCrossSection = "FlatOval684x484"
            ElseIf dWidth > 0.684 - dClearance And dWidth <= 0.784 - dClearance Then
                strCrossSection = "FlatOval784x484"
            ElseIf dWidth > 0.784 - dClearance And dWidth <= 0.884 - dClearance Then
                strCrossSection = "FlatOval884x484"
            ElseIf dWidth > 0.884 - dClearance And dWidth <= 0.984 - dClearance Then
                strCrossSection = "FlatOval984x484"
            ElseIf dWidth > 0.984 - dClearance And dWidth <= 1.084 - dClearance Then
                strCrossSection = "FlatOval1084x484"
            Else
                strCrossSection = "FlatOval284x484"
                UpdateStatusBar "No support found for selected Cableway, set the support to FlatOval284x484"
            End If
        ElseIf dDepth > 0.484 - dClearance And dDepth <= 0.584 - dClearance Then
            If dWidth > 0# And dWidth <= 0.284 - dClearance Then
                strCrossSection = "FlatOval284x584"
            ElseIf dWidth > 0.284 - dClearance And dWidth <= 0.384 - dClearance Then
                strCrossSection = "FlatOval384x584"
            ElseIf dWidth > 0.384 - dClearance And dWidth <= 0.484 - dClearance Then
                strCrossSection = "FlatOval484x584"
            ElseIf dWidth > 0.484 - dClearance And dWidth <= 0.584 - dClearance Then
                strCrossSection = "FlatOval584x584"
            ElseIf dWidth > 0.584 - dClearance And dWidth <= 0.684 - dClearance Then
                strCrossSection = "FlatOval684x584"
            ElseIf dWidth > 0.684 - dClearance And dWidth <= 0.784 - dClearance Then
                strCrossSection = "FlatOval784x584"
            ElseIf dWidth > 0.784 - dClearance And dWidth <= 0.884 - dClearance Then
                strCrossSection = "FlatOval884x584"
            ElseIf dWidth > 0.884 - dClearance And dWidth <= 0.984 - dClearance Then
                strCrossSection = "FlatOval984x584"
            ElseIf dWidth > 0.984 - dClearance And dWidth <= 1.084 - dClearance Then
                strCrossSection = "FlatOval1084x584"
            Else
                strCrossSection = "FlatOval284x584"
                UpdateStatusBar "No support found for selected Cableway, set the support to FlatOval284x584"
            End If
        ElseIf dDepth > 0.584 - dClearance And dDepth <= 0.684 - dClearance Then
            If dWidth > 0# And dWidth <= 0.284 - dClearance Then
                strCrossSection = "FlatOval284x684"
            ElseIf dWidth > 0.284 - dClearance And dWidth <= 0.384 - dClearance Then
                strCrossSection = "FlatOval384x684"
            ElseIf dWidth > 0.384 - dClearance And dWidth <= 0.484 - dClearance Then
                strCrossSection = "FlatOval484x684"
            ElseIf dWidth > 0.484 - dClearance And dWidth <= 0.584 - dClearance Then
                strCrossSection = "FlatOval584x684"
            ElseIf dWidth > 0.584 - dClearance And dWidth <= 0.684 - dClearance Then
                strCrossSection = "FlatOval684x684"
            ElseIf dWidth > 0.684 - dClearance And dWidth <= 0.784 - dClearance Then
                strCrossSection = "FlatOval784x684"
            ElseIf dWidth > 0.784 - dClearance And dWidth <= 0.884 - dClearance Then
                strCrossSection = "FlatOval884x684"
            ElseIf dWidth > 0.884 - dClearance And dWidth <= 0.984 - dClearance Then
                strCrossSection = "FlatOval984x684"
            ElseIf dWidth > 0.984 - dClearance And dWidth <= 1.084 - dClearance Then
                strCrossSection = "FlatOval1084x684"
            Else
                strCrossSection = "FlatOval284x684"
                UpdateStatusBar "No support found for selected Cableway, set the support to FlatOval284x684"
            End If
        ElseIf dDepth > 0.684 - dClearance And dDepth <= 0.784 - dClearance Then
            If dWidth > 0# And dWidth <= 0.284 - dClearance Then
                strCrossSection = "FlatOval284x784"
            ElseIf dWidth > 0.284 - dClearance And dWidth <= 0.384 - dClearance Then
                strCrossSection = "FlatOval384x784"
            ElseIf dWidth > 0.384 - dClearance And dWidth <= 0.484 - dClearance Then
                strCrossSection = "FlatOval484x784"
            ElseIf dWidth > 0.484 - dClearance And dWidth <= 0.584 - dClearance Then
                strCrossSection = "FlatOval584x784"
            ElseIf dWidth > 0.584 - dClearance And dWidth <= 0.684 - dClearance Then
                strCrossSection = "FlatOval684x784"
            ElseIf dWidth > 0.684 - dClearance And dWidth <= 0.784 - dClearance Then
                strCrossSection = "FlatOval784x784"
            ElseIf dWidth > 0.784 - dClearance And dWidth <= 0.884 - dClearance Then
                strCrossSection = "FlatOval884x784"
            ElseIf dWidth > 0.884 - dClearance And dWidth <= 0.984 - dClearance Then
                strCrossSection = "FlatOval984x784"
            ElseIf dWidth > 0.984 - dClearance And dWidth <= 1.084 - dClearance Then
                strCrossSection = "FlatOval1084x784"
            Else
                strCrossSection = "FlatOval284x784"
                UpdateStatusBar "No support found for selected Cableway, set the support to FlatOval284x784"
            End If
        ElseIf dDepth > 0.784 - dClearance And dDepth <= 0.884 - dClearance Then
            If dWidth > 0# And dWidth <= 0.284 - dClearance Then
                strCrossSection = "FlatOval284x884"
            ElseIf dWidth > 0.284 - dClearance And dWidth <= 0.384 - dClearance Then
                strCrossSection = "FlatOval384x884"
            ElseIf dWidth > 0.384 - dClearance And dWidth <= 0.484 - dClearance Then
                strCrossSection = "FlatOval484x884"
            ElseIf dWidth > 0.484 - dClearance And dWidth <= 0.584 - dClearance Then
                strCrossSection = "FlatOval584x884"
            ElseIf dWidth > 0.584 - dClearance And dWidth <= 0.684 - dClearance Then
                strCrossSection = "FlatOval684x884"
            ElseIf dWidth > 0.684 - dClearance And dWidth <= 0.784 - dClearance Then
                strCrossSection = "FlatOval784x884"
            ElseIf dWidth > 0.784 - dClearance And dWidth <= 0.884 - dClearance Then
                strCrossSection = "FlatOval884x884"
            ElseIf dWidth > 0.884 - dClearance And dWidth <= 0.984 - dClearance Then
                strCrossSection = "FlatOval984x884"
            ElseIf dWidth > 0.984 - dClearance And dWidth <= 1.084 - dClearance Then
                strCrossSection = "FlatOval1084x884"
            Else
                strCrossSection = "FlatOval284x884"
                UpdateStatusBar "No support found for selected Cableway, set the support to FlatOval284x884"
            End If
        ElseIf dDepth > 0.884 - dClearance And dDepth <= 0.984 - dClearance Then
            If dWidth > 0# And dWidth <= 0.284 - dClearance Then
                strCrossSection = "FlatOval284x984"
            ElseIf dWidth > 0.284 - dClearance And dWidth <= 0.384 - dClearance Then
                strCrossSection = "FlatOval384x984"
            ElseIf dWidth > 0.384 - dClearance And dWidth <= 0.484 - dClearance Then
                strCrossSection = "FlatOval484x984"
            ElseIf dWidth > 0.484 - dClearance And dWidth <= 0.584 - dClearance Then
                strCrossSection = "FlatOval584x984"
            ElseIf dWidth > 0.584 - dClearance And dWidth <= 0.684 - dClearance Then
                strCrossSection = "FlatOval684x984"
            ElseIf dWidth > 0.684 - dClearance And dWidth <= 0.784 - dClearance Then
                strCrossSection = "FlatOval784x984"
            ElseIf dWidth > 0.784 - dClearance And dWidth <= 0.884 - dClearance Then
                strCrossSection = "FlatOval884x984"
            ElseIf dWidth > 0.884 - dClearance And dWidth <= 0.984 - dClearance Then
                strCrossSection = "FlatOval984x984"
            ElseIf dWidth > 0.984 - dClearance And dWidth <= 1.084 - dClearance Then
                strCrossSection = "FlatOval1084x984"
            Else
                strCrossSection = "FlatOval284x984"
                UpdateStatusBar "No support found for selected Cableway, set the support to FlatOval284x984"
            End If
        ElseIf dDepth > 0.984 - dClearance And dDepth <= 1.084 - dClearance Then
            If dWidth > 0# And dWidth <= 0.284 - dClearance Then
                strCrossSection = "FlatOval284x1084"
            ElseIf dWidth > 0.284 - dClearance And dWidth <= 0.384 - dClearance Then
                strCrossSection = "FlatOval384x1084"
            ElseIf dWidth > 0.384 - dClearance And dWidth <= 0.484 - dClearance Then
                strCrossSection = "FlatOval484x1084"
            ElseIf dWidth > 0.484 - dClearance And dWidth <= 0.584 - dClearance Then
                strCrossSection = "FlatOval584x1084"
            ElseIf dWidth > 0.584 - dClearance And dWidth <= 0.684 - dClearance Then
                strCrossSection = "FlatOval684x1084"
            ElseIf dWidth > 0.684 - dClearance And dWidth <= 0.784 - dClearance Then
                strCrossSection = "FlatOval784x1084"
            ElseIf dWidth > 0.784 - dClearance And dWidth <= 0.884 - dClearance Then
                strCrossSection = "FlatOval884x1084"
            ElseIf dWidth > 0.884 - dClearance And dWidth <= 0.984 - dClearance Then
                strCrossSection = "FlatOval984x1084"
            ElseIf dWidth > 0.984 - dClearance And dWidth <= 1.084 - dClearance Then
                strCrossSection = "FlatOval1084x1084"
            Else
                strCrossSection = "FlatOval284x1084"
                UpdateStatusBar "No support found for selected Cableway, set the support to FlatOval284x1084"
            End If
        Else
            strCrossSection = "FlatOval284x124"
            UpdateStatusBar "No support found for selected Cableway, set the support to FlatOval284x124"
        End If

    Else
        strCrossSection = "Rect284x124"
        UpdateStatusBar "No support found for selected Cableway, set the support to Rect284x124"
    End If
           
    pSLH.Add strCrossSection
    AddValidCWType pSLH, strCrossSection
           
    GoTo Cleanup
    
ErrorHandler:
    pSLH.ReportError "SelectorLogic"
    
Cleanup:
    Set oHoleTrace = Nothing
    Set oHoleSmartOcc = Nothing
    
End Sub

'******************************************************************************
' Routine: AddValidCWType
'
' Abstract: Add the valid cw type except the selected CW hole type
'
' Notes:
'   It has been decided to use only other types close in size to the one
'   selected rather than the entire set of types.  Previously all
'   cableway types were added but then if a new type were selected the
'   Smart Occurrence would always reject it because the old type was in
'   the list.
'
'******************************************************************************
Private Sub AddValidCWType(pSLH As IJDSelectorLogic, sType As String)
    On Error GoTo ErrorHandler
    
    sMETHOD = "AddValidCWType"
    sERROR = "General Error"
    
    If sType = "Rect284x124" Then
        pSLH.Add "Rect284x224"      ' higher
        pSLH.Add "Rect384x124"      ' wider

    ElseIf sType = "Rect284x224" Then
        pSLH.Add "Rect284x374"      ' higher
        pSLH.Add "Rect284x124"      ' lower
        pSLH.Add "Rect384x224"      ' wider

    ElseIf sType = "Rect284x374" Then
        pSLH.Add "Rect284x524"      ' higher
        pSLH.Add "Rect284x224"      ' lower
        pSLH.Add "Rect384x374"      ' wider

    ElseIf sType = "Rect284x524" Then
        pSLH.Add "Rect284x674"      ' higher
        pSLH.Add "Rect284x374"      ' lower
        pSLH.Add "Rect384x524"      ' wider

    ElseIf sType = "Rect284x674" Then
        pSLH.Add "Rect284x524"      ' lower
        pSLH.Add "Rect384x674"      ' wider

    ElseIf sType = "Rect384x124" Then
        pSLH.Add "Rect384x224"      ' higher
        pSLH.Add "Rect484x124"      ' wider
        pSLH.Add "Rect284x124"      ' narrower

    ElseIf sType = "Rect384x224" Then
        pSLH.Add "Rect384x374"      ' higher
        pSLH.Add "Rect384x124"      ' lower
        pSLH.Add "Rect484x224"      ' wider
        pSLH.Add "Rect284x224"      ' narrower

    ElseIf sType = "Rect384x374" Then
        pSLH.Add "Rect384x524"      ' higher
        pSLH.Add "Rect384x224"      ' lower
        pSLH.Add "Rect484x374"      ' wider
        pSLH.Add "Rect284x374"      ' narrower

    ElseIf sType = "Rect384x524" Then
        pSLH.Add "Rect384x674"      ' higher
        pSLH.Add "Rect384x374"      ' lower
        pSLH.Add "Rect484x524"      ' wider
        pSLH.Add "Rect284x524"      ' narrower

    ElseIf sType = "Rect384x674" Then
        pSLH.Add "Rect384x524"      ' lower
        pSLH.Add "Rect484x674"      ' wider
        pSLH.Add "Rect284x674"      ' narrower

    ElseIf sType = "Rect484x124" Then
        pSLH.Add "Rect484x224"      ' higher
        pSLH.Add "Rect584x124"      ' wider
        pSLH.Add "Rect384x124"      ' narrower

    ElseIf sType = "Rect484x224" Then
        pSLH.Add "Rect484x374"      ' higher
        pSLH.Add "Rect484x124"      ' lower
        pSLH.Add "Rect584x224"      ' wider
        pSLH.Add "Rect384x224"      ' narrower

    ElseIf sType = "Rect484x374" Then
        pSLH.Add "Rect484x524"      ' higher
        pSLH.Add "Rect484x224"      ' lower
        pSLH.Add "Rect584x374"      ' wider
        pSLH.Add "Rect384x374"      ' narrower

    ElseIf sType = "Rect484x524" Then
        pSLH.Add "Rect484x674"      ' higher
        pSLH.Add "Rect484x374"      ' lower
        pSLH.Add "Rect584x524"      ' wider
        pSLH.Add "Rect384x524"      ' narrower

    ElseIf sType = "Rect484x674" Then
        pSLH.Add "Rect484x524"      ' lower
        pSLH.Add "Rect584x674"      ' wider
        pSLH.Add "Rect384x674"      ' narrower

    ElseIf sType = "Rect584x124" Then
        pSLH.Add "Rect584x224"      ' higher
        pSLH.Add "Rect684x124"      ' wider
        pSLH.Add "Rect484x124"      ' narrower

    ElseIf sType = "Rect584x224" Then
        pSLH.Add "Rect584x374"      ' higher
        pSLH.Add "Rect584x124"      ' lower
        pSLH.Add "Rect684x224"      ' wider
        pSLH.Add "Rect484x224"      ' narrower

    ElseIf sType = "Rect584x374" Then
        pSLH.Add "Rect584x524"      ' higher
        pSLH.Add "Rect584x224"      ' lower
        pSLH.Add "Rect684x374"      ' wider
        pSLH.Add "Rect484x374"      ' narrower

    ElseIf sType = "Rect584x524" Then
        pSLH.Add "Rect584x674"      ' higher
        pSLH.Add "Rect584x374"      ' lower
        pSLH.Add "Rect684x524"      ' wider
        pSLH.Add "Rect484x524"      ' narrower

    ElseIf sType = "Rect584x674" Then
        pSLH.Add "Rect584x524"      ' lower
        pSLH.Add "Rect684x674"      ' wider
        pSLH.Add "Rect484x674"      ' narrower

    ElseIf sType = "Rect684x124" Then
        pSLH.Add "Rect684x224"      ' higher
        pSLH.Add "Rect784x124"      ' wider
        pSLH.Add "Rect584x124"      ' narrower

    ElseIf sType = "Rect684x224" Then
        pSLH.Add "Rect684x374"      ' higher
        pSLH.Add "Rect684x124"      ' lower
        pSLH.Add "Rect784x224"      ' wider
        pSLH.Add "Rect584x224"      ' narrower

    ElseIf sType = "Rect684x374" Then
        pSLH.Add "Rect684x524"      ' higher
        pSLH.Add "Rect684x224"      ' lower
        pSLH.Add "Rect784x374"      ' wider
        pSLH.Add "Rect584x374"      ' narrower

    ElseIf sType = "Rect684x524" Then
        pSLH.Add "Rect684x674"      ' higher
        pSLH.Add "Rect684x374"      ' lower
        pSLH.Add "Rect784x524"      ' wider
        pSLH.Add "Rect584x524"      ' narrower

    ElseIf sType = "Rect684x674" Then
        pSLH.Add "Rect684x524"      ' lower
        pSLH.Add "Rect784x674"      ' wider
        pSLH.Add "Rect584x674"      ' narrower

    ElseIf sType = "Rect784x124" Then
        pSLH.Add "Rect784x224"      ' higher
        pSLH.Add "Rect884x124"      ' wider
        pSLH.Add "Rect684x124"      ' narrower

    ElseIf sType = "Rect784x224" Then
        pSLH.Add "Rect784x374"      ' higher
        pSLH.Add "Rect784x124"      ' lower
        pSLH.Add "Rect884x224"      ' wider
        pSLH.Add "Rect684x224"      ' narrower

    ElseIf sType = "Rect784x374" Then
        pSLH.Add "Rect784x524"      ' higher
        pSLH.Add "Rect784x224"      ' lower
        pSLH.Add "Rect884x374"      ' wider
        pSLH.Add "Rect684x374"      ' narrower

    ElseIf sType = "Rect784x524" Then
        pSLH.Add "Rect784x674"      ' higher
        pSLH.Add "Rect784x374"      ' lower
        pSLH.Add "Rect884x524"      ' wider
        pSLH.Add "Rect684x524"      ' narrower

    ElseIf sType = "Rect784x674" Then
        pSLH.Add "Rect784x524"      ' lower
        pSLH.Add "Rect884x674"      ' wider
        pSLH.Add "Rect684x674"      ' narrower

    ElseIf sType = "Rect884x124" Then
        pSLH.Add "Rect884x224"      ' higher
        pSLH.Add "Rect984x124"      ' wider
        pSLH.Add "Rect784x124"      ' narrower

    ElseIf sType = "Rect884x224" Then
        pSLH.Add "Rect884x374"      ' higher
        pSLH.Add "Rect884x124"      ' lower
        pSLH.Add "Rect984x224"      ' wider
        pSLH.Add "Rect784x224"      ' narrower

    ElseIf sType = "Rect884x374" Then
        pSLH.Add "Rect884x524"      ' higher
        pSLH.Add "Rect884x224"      ' lower
        pSLH.Add "Rect984x374"      ' wider
        pSLH.Add "Rect784x374"      ' narrower

    ElseIf sType = "Rect884x524" Then
        pSLH.Add "Rect884x674"      ' higher
        pSLH.Add "Rect884x374"      ' lower
        pSLH.Add "Rect984x524"      ' wider
        pSLH.Add "Rect784x524"      ' narrower

    ElseIf sType = "Rect884x674" Then
        pSLH.Add "Rect884x524"      ' lower
        pSLH.Add "Rect984x674"      ' wider
        pSLH.Add "Rect784x674"      ' narrower

    ElseIf sType = "Rect984x124" Then
        pSLH.Add "Rect984x224"      ' higher
        pSLH.Add "Rect1084x124"     ' wider
        pSLH.Add "Rect884x124"      ' narrower

    ElseIf sType = "Rect984x224" Then
        pSLH.Add "Rect984x374"      ' higher
        pSLH.Add "Rect984x124"      ' lower
        pSLH.Add "Rect1084x224"     ' wider
        pSLH.Add "Rect884x224"      ' narrower

    ElseIf sType = "Rect984x374" Then
        pSLH.Add "Rect984x524"      ' higher
        pSLH.Add "Rect984x224"      ' lower
        pSLH.Add "Rect1084x374"     ' wider
        pSLH.Add "Rect884x374"      ' narrower

    ElseIf sType = "Rect984x524" Then
        pSLH.Add "Rect984x674"      ' higher
        pSLH.Add "Rect984x374"      ' lower
        pSLH.Add "Rect1084x524"     ' wider
        pSLH.Add "Rect884x524"      ' narrower

    ElseIf sType = "Rect984x674" Then
        pSLH.Add "Rect984x524"      ' lower
        pSLH.Add "Rect1084x674"     ' wider
        pSLH.Add "Rect884x674"      ' narrower

    ElseIf sType = "Rect1084x124" Then
        pSLH.Add "Rect1084x224"     ' higher
        pSLH.Add "Rect984x124"      ' narrower

    ElseIf sType = "Rect1084x224" Then
        pSLH.Add "Rect1084x374"     ' higher
        pSLH.Add "Rect1084x124"     ' lower
        pSLH.Add "Rect984x224"      ' narrower

    ElseIf sType = "Rect1084x374" Then
        pSLH.Add "Rect1084x524"     ' higher
        pSLH.Add "Rect1084x224"     ' lower
        pSLH.Add "Rect984x374"      ' narrower

    ElseIf sType = "Rect1084x524" Then
        pSLH.Add "Rect1084x674"     ' higher
        pSLH.Add "Rect1084x374"     ' lower
        pSLH.Add "Rect984x524"      ' narrower

    ElseIf sType = "Rect1084x674" Then
        pSLH.Add "Rect1084x524"     ' lower
        pSLH.Add "Rect984x674"      ' narrower

    ElseIf sType = "Round62" Then
        pSLH.Add "Round92"

    ElseIf sType = "Round92" Then
        pSLH.Add "Round62"
        
    ElseIf sType = "FlatOval284x124" Then
        pSLH.Add "FlatOval284x284"      ' higher
        pSLH.Add "FlatOval384x124"      ' wider

    ElseIf sType = "FlatOval284x284" Then
        pSLH.Add "FlatOval284x384"      ' higher
        pSLH.Add "FlatOval284x124"      ' lower
        pSLH.Add "FlatOval384x284"      ' wider

    ElseIf sType = "FlatOval284x384" Then
        pSLH.Add "FlatOval284x584"      ' higher
        pSLH.Add "FlatOval284x284"      ' lower
        pSLH.Add "FlatOval384x384"      ' wider

    ElseIf sType = "FlatOval284x584" Then
        pSLH.Add "FlatOval284x684"      ' higher
        pSLH.Add "FlatOval284x384"      ' lower
        pSLH.Add "FlatOval384x584"      ' wider

    ElseIf sType = "FlatOval284x684" Then
        pSLH.Add "FlatOval284x584"      ' lower
        pSLH.Add "FlatOval384x684"      ' wider

    ElseIf sType = "FlatOval384x124" Then
        pSLH.Add "FlatOval384x284"      ' higher
        pSLH.Add "FlatOval484x124"      ' wider
        pSLH.Add "FlatOval284x124"      ' narrower

    ElseIf sType = "FlatOval384x284" Then
        pSLH.Add "FlatOval384x384"      ' higher
        pSLH.Add "FlatOval384x124"      ' lower
        pSLH.Add "FlatOval484x284"      ' wider
        pSLH.Add "FlatOval284x284"      ' narrower

    ElseIf sType = "FlatOval384x384" Then
        pSLH.Add "FlatOval384x584"      ' higher
        pSLH.Add "FlatOval384x284"      ' lower
        pSLH.Add "FlatOval484x384"      ' wider
        pSLH.Add "FlatOval284x384"      ' narrower

    ElseIf sType = "FlatOval384x584" Then
        pSLH.Add "FlatOval384x684"      ' higher
        pSLH.Add "FlatOval384x384"      ' lower
        pSLH.Add "FlatOval484x584"      ' wider
        pSLH.Add "FlatOval284x584"      ' narrower

    ElseIf sType = "FlatOval384x684" Then
        pSLH.Add "FlatOval384x584"      ' lower
        pSLH.Add "FlatOval484x684"      ' wider
        pSLH.Add "FlatOval284x684"      ' narrower

    ElseIf sType = "FlatOval484x124" Then
        pSLH.Add "FlatOval484x224"      ' higher
        pSLH.Add "FlatOval584x124"      ' wider
        pSLH.Add "FlatOval384x124"      ' narrower

    ElseIf sType = "FlatOval484x284" Then
        pSLH.Add "FlatOval484x384"      ' higher
        pSLH.Add "FlatOval484x124"      ' lower
        pSLH.Add "FlatOval584x284"      ' wider
        pSLH.Add "FlatOval384x284"      ' narrower

    ElseIf sType = "FlatOval484x384" Then
        pSLH.Add "FlatOval484x584"      ' higher
        pSLH.Add "FlatOval484x284"      ' lower
        pSLH.Add "FlatOval584x384"      ' wider
        pSLH.Add "FlatOval384x384"      ' narrower

    ElseIf sType = "FlatOval484x584" Then
        pSLH.Add "FlatOval484x684"      ' higher
        pSLH.Add "FlatOval484x384"      ' lower
        pSLH.Add "FlatOval584x584"      ' wider
        pSLH.Add "FlatOval384x584"      ' narrower

    ElseIf sType = "FlatOval484x684" Then
        pSLH.Add "FlatOval484x584"      ' lower
        pSLH.Add "FlatOval584x684"      ' wider
        pSLH.Add "FlatOval384x684"      ' narrower

    ElseIf sType = "FlatOval584x124" Then
        pSLH.Add "FlatOval584x284"      ' higher
        pSLH.Add "FlatOval684x124"      ' wider
        pSLH.Add "FlatOval484x124"      ' narrower

    ElseIf sType = "FlatOval584x284" Then
        pSLH.Add "FlatOval584x384"      ' higher
        pSLH.Add "FlatOval584x124"      ' lower
        pSLH.Add "FlatOval684x284"      ' wider
        pSLH.Add "FlatOval484x284"      ' narrower

    ElseIf sType = "FlatOval584x384" Then
        pSLH.Add "FlatOval584x584"      ' higher
        pSLH.Add "FlatOval584x284"      ' lower
        pSLH.Add "FlatOval684x384"      ' wider
        pSLH.Add "FlatOval484x384"      ' narrower

    ElseIf sType = "FlatOval584x584" Then
        pSLH.Add "FlatOval584x684"      ' higher
        pSLH.Add "FlatOval584x384"      ' lower
        pSLH.Add "FlatOval684x584"      ' wider
        pSLH.Add "FlatOval484x584"      ' narrower

    ElseIf sType = "FlatOval584x684" Then
        pSLH.Add "FlatOval584x584"      ' lower
        pSLH.Add "FlatOval684x684"      ' wider
        pSLH.Add "FlatOval484x684"      ' narrower

    ElseIf sType = "FlatOval684x124" Then
        pSLH.Add "FlatOval684x284"      ' higher
        pSLH.Add "FlatOval784x124"      ' wider
        pSLH.Add "FlatOval584x124"      ' narrower

    ElseIf sType = "FlatOval684x284" Then
        pSLH.Add "FlatOval684x384"      ' higher
        pSLH.Add "FlatOval684x124"      ' lower
        pSLH.Add "FlatOval784x284"      ' wider
        pSLH.Add "FlatOval584x284"      ' narrower

    ElseIf sType = "FlatOval684x384" Then
        pSLH.Add "FlatOval684x584"      ' higher
        pSLH.Add "FlatOval684x284"      ' lower
        pSLH.Add "FlatOval784x384"      ' wider
        pSLH.Add "FlatOval584x384"      ' narrower

    ElseIf sType = "FlatOval684x584" Then
        pSLH.Add "FlatOval684x684"      ' higher
        pSLH.Add "FlatOval684x384"      ' lower
        pSLH.Add "FlatOval784x584"      ' wider
        pSLH.Add "FlatOval584x584"      ' narrower

    ElseIf sType = "FlatOval684x684" Then
        pSLH.Add "FlatOval684x584"      ' lower
        pSLH.Add "FlatOval784x684"      ' wider
        pSLH.Add "FlatOval584x684"      ' narrower

    ElseIf sType = "FlatOval784x124" Then
        pSLH.Add "FlatOval784x284"      ' higher
        pSLH.Add "FlatOval884x124"      ' wider
        pSLH.Add "FlatOval684x124"      ' narrower

    ElseIf sType = "FlatOval784x284" Then
        pSLH.Add "FlatOval784x384"      ' higher
        pSLH.Add "FlatOval784x124"      ' lower
        pSLH.Add "FlatOval884x284"      ' wider
        pSLH.Add "FlatOval684x284"      ' narrower

    ElseIf sType = "FlatOval784x384" Then
        pSLH.Add "FlatOval784x584"      ' higher
        pSLH.Add "FlatOval784x284"      ' lower
        pSLH.Add "FlatOval884x384"      ' wider
        pSLH.Add "FlatOval684x384"      ' narrower

    ElseIf sType = "FlatOval784x584" Then
        pSLH.Add "FlatOval784x684"      ' higher
        pSLH.Add "FlatOval784x384"      ' lower
        pSLH.Add "FlatOval884x584"      ' wider
        pSLH.Add "FlatOval684x584"      ' narrower

    ElseIf sType = "FlatOval784x684" Then
        pSLH.Add "FlatOval784x584"      ' lower
        pSLH.Add "FlatOval884x684"      ' wider
        pSLH.Add "FlatOval684x684"      ' narrower

    ElseIf sType = "FlatOval884x124" Then
        pSLH.Add "FlatOval884x284"      ' higher
        pSLH.Add "FlatOval984x124"      ' wider
        pSLH.Add "FlatOval784x124"      ' narrower

    ElseIf sType = "FlatOval884x284" Then
        pSLH.Add "FlatOval884x384"      ' higher
        pSLH.Add "FlatOval884x124"      ' lower
        pSLH.Add "FlatOval984x284"      ' wider
        pSLH.Add "FlatOval784x284"      ' narrower

    ElseIf sType = "FlatOval884x384" Then
        pSLH.Add "FlatOval884x584"      ' higher
        pSLH.Add "FlatOval884x284"      ' lower
        pSLH.Add "FlatOval984x384"      ' wider
        pSLH.Add "FlatOval784x384"      ' narrower

    ElseIf sType = "FlatOval884x584" Then
        pSLH.Add "FlatOval884x684"      ' higher
        pSLH.Add "FlatOval884x384"      ' lower
        pSLH.Add "FlatOval984x584"      ' wider
        pSLH.Add "FlatOval784x584"      ' narrower

    ElseIf sType = "FlatOval884x684" Then
        pSLH.Add "FlatOval884x584"      ' lower
        pSLH.Add "FlatOval984x684"      ' wider
        pSLH.Add "FlatOval784x684"      ' narrower

    ElseIf sType = "FlatOval984x124" Then
        pSLH.Add "FlatOval984x284"      ' higher
        pSLH.Add "FlatOval1084x124"     ' wider
        pSLH.Add "FlatOval884x124"      ' narrower

    ElseIf sType = "FlatOval984x284" Then
        pSLH.Add "FlatOval984x384"      ' higher
        pSLH.Add "FlatOval984x124"      ' lower
        pSLH.Add "FlatOval1084x284"     ' wider
        pSLH.Add "FlatOval884x284"      ' narrower

    ElseIf sType = "FlatOval984x384" Then
        pSLH.Add "FlatOval984x584"      ' higher
        pSLH.Add "FlatOval984x284"      ' lower
        pSLH.Add "FlatOval1084x384"     ' wider
        pSLH.Add "FlatOval884x384"      ' narrower

    ElseIf sType = "FlatOval984x584" Then
        pSLH.Add "FlatOval984x684"      ' higher
        pSLH.Add "FlatOval984x384"      ' lower
        pSLH.Add "FlatOval1084x584"     ' wider
        pSLH.Add "FlatOval884x584"      ' narrower

    ElseIf sType = "FlatOval984x684" Then
        pSLH.Add "FlatOval984x584"      ' lower
        pSLH.Add "FlatOval1084x684"     ' wider
        pSLH.Add "FlatOval884x684"      ' narrower

    ElseIf sType = "FlatOval1084x124" Then
        pSLH.Add "FlatOval1084x284"     ' higher
        pSLH.Add "FlatOval984x124"      ' narrower

    ElseIf sType = "FlatOval1084x284" Then
        pSLH.Add "FlatOval1084x384"     ' higher
        pSLH.Add "FlatOval1084x124"     ' lower
        pSLH.Add "FlatOval984x284"      ' narrower

    ElseIf sType = "FlatOval1084x384" Then
        pSLH.Add "FlatOval1084x584"     ' higher
        pSLH.Add "FlatOval1084x284"     ' lower
        pSLH.Add "FlatOval984x384"      ' narrower

    ElseIf sType = "FlatOval1084x584" Then
        pSLH.Add "FlatOval1084x684"     ' higher
        pSLH.Add "FlatOval1084x384"     ' lower
        pSLH.Add "FlatOval984x584"      ' narrower

    ElseIf sType = "FlatOval1084x684" Then
        pSLH.Add "FlatOval1084x584"     ' lower
        pSLH.Add "FlatOval984x684"      ' narrower

    End If
    
    GoTo Cleanup
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
    
Cleanup:
    
End Sub

'******************************************************************************
' Routine: AllValidCWType
'
' Abstract: Set the all type of CW to Collection for Modify
'
' Description: User can change the CW hole type among followings.
'
' Notes:
'   This method is no longer called.  But it does provide a nice list of
'   all cableway types.
'
'******************************************************************************
Private Function AllValidCWType() As Collection
    On Error GoTo ErrorHandler

    sMETHOD = "AllValidCWType"
    sERROR = "General Error"
    
    Dim colValidCWs As New Collection
    colValidCWs.Add "Rect284x124"
    colValidCWs.Add "Rect284x224"
    colValidCWs.Add "Rect284x374"
    colValidCWs.Add "Rect284x524"
    colValidCWs.Add "Rect284x674"
    colValidCWs.Add "Rect384x124"
    colValidCWs.Add "Rect384x224"
    colValidCWs.Add "Rect384x374"
    colValidCWs.Add "Rect384x524"
    colValidCWs.Add "Rect384x674"
    colValidCWs.Add "Rect484x124"
    colValidCWs.Add "Rect484x224"
    colValidCWs.Add "Rect484x374"
    colValidCWs.Add "Rect484x524"
    colValidCWs.Add "Rect484x674"
    colValidCWs.Add "Rect584x124"
    colValidCWs.Add "Rect584x224"
    colValidCWs.Add "Rect584x374"
    colValidCWs.Add "Rect584x524"
    colValidCWs.Add "Rect584x674"
    colValidCWs.Add "Rect684x124"
    colValidCWs.Add "Rect684x224"
    colValidCWs.Add "Rect684x374"
    colValidCWs.Add "Rect684x524"
    colValidCWs.Add "Rect684x674"
    colValidCWs.Add "Rect784x124"
    colValidCWs.Add "Rect784x224"
    colValidCWs.Add "Rect784x374"
    colValidCWs.Add "Rect784x524"
    colValidCWs.Add "Rect784x674"
    colValidCWs.Add "Rect884x124"
    colValidCWs.Add "Rect884x224"
    colValidCWs.Add "Rect884x374"
    colValidCWs.Add "Rect884x524"
    colValidCWs.Add "Rect884x674"
    colValidCWs.Add "Rect984x124"
    colValidCWs.Add "Rect984x224"
    colValidCWs.Add "Rect984x374"
    colValidCWs.Add "Rect984x524"
    colValidCWs.Add "Rect984x674"
    colValidCWs.Add "Rect1084x124"
    colValidCWs.Add "Rect1084x224"
    colValidCWs.Add "Rect1084x374"
    colValidCWs.Add "Rect1084x524"
    colValidCWs.Add "Rect1084x674"
    colValidCWs.Add "Round62"
    colValidCWs.Add "Round92"
     
    Set AllValidCWType = colValidCWs
    Set colValidCWs = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

'******************************************************************************
' Routine: GetWidthDepthOfCableway
'
' Abstract: Get width and depth of Cable way
'
' Description:
'
'******************************************************************************
Public Sub GetWidthDepthOfCableway(ByVal oHoleTrace As IJHoleTraceAE, oSelectedType As CrossSectionShapeTypes, _
                                    oWidth As Double, oDepth As Double)
    'Get the Outfitting, Only Cable Way should be taken.
    Dim oOutfitting As IJDObjectCollection
    Set oOutfitting = oHoleTrace.GetParentOutfitting
    
    Dim oObj As Object
    Dim oRtePathFeat As IJRtePathFeat
    Dim oCableRad As Double
    Dim bOutderDia As Boolean
'    Dim oAngle As Double, oTempWidth As Double
'    Const PI = 3.14159265
    
    'For Loop input object ( 1 .. icount .. n ),
    'In case of CW, only one CW should be taken and hook up the proper GSCAD symbol file.
    For Each oObj In oOutfitting
        If TypeOf oObj Is IJRteCablewayPathFeat Then

            Set oRtePathFeat = oObj
            Dim oRteFeatUtility As IJRtePathCrossSectUtility
            Set oRteFeatUtility = oRtePathFeat
                            
            'Get the type , width , Depth from cableway....
            If Not oRteFeatUtility Is Nothing Then
                oRteFeatUtility.GetCrossSectionData Nothing, False, oSelectedType, _
                                    oWidth, oDepth, oCableRad, bOutderDia
            End If

            Set oRteFeatUtility = Nothing
            Set oRtePathFeat = Nothing
            Exit For
        End If
    Next oObj
    Set oObj = Nothing
    Set oOutfitting = Nothing

End Sub

'******************************************************************************
' Routine: GetOrientationAngle
'
' Abstract: Get the cross section's orientation angle
'
' Description: Orientation angle means the angle how much the CableWay itself
'              rotates at the same position. oUVector and oVVector are
'              the vectors of cross section.
'******************************************************************************
Private Function GetOrientationAngle(oUVector As IJDVector, oVVector As IJDVector) As Double
       
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oNormal As New DVector
    Dim oNewUVector As New DVector
    Dim oPlane As IJPlane
    
    Dim Ux As Double, Uy As Double, Uz As Double
    Dim Nx As Double, Ny As Double, Nz As Double
        
    oUVector.Get Ux, Uy, Uz
    
    'get normal vector of uvector and vvector
    Set oNormal = oUVector.Cross(oVVector)
    oNormal.Get Nx, Ny, Nz
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
            
    'If new plane is created, uvector and vvector are changed to be mapped to the origin axis.
    Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, Ux, Uy, Uz, Nx, Ny, Nz)
    oPlane.GetUDirection Ux, Uy, Uz
    
    'create a comparing uvector by using the previous uvector
    oNewUVector.Set Ux, Uy, Uz
    
    'get angle
    GetOrientationAngle = oNewUVector.Angle(oUVector, oNormal)
    
    Set oGeomFactory = Nothing
    Set oNormal = Nothing
    Set oNewUVector = Nothing
    Set oPlane = Nothing
    
End Function

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  pSelector.IJDInputs.RemoveAllInput
  pSelector.IJDRepresentations.RemoveAllRepresentation
  
  Dim pSelectorFact As New DefinitionFactory
  pSelectorFact.InitAbstractSelector pSelector
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.definition = pSelector
  pIH.InitAs m_FamilyProgid
  SelectorInputs pIH
  Dim pQH As IJDQuestionsHelper
  Set pQH = New QuestionHelper
  pQH.Selector = pSelector
  SelectorQuestions pQH
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pSelectorFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pSelectorFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMSelector(pRep As IJDRepresentation)
  Dim pSLH As IJDSelectorLogic
  Set pSLH = New SelectorLogic
  pSLH.Representation = pRep
  SelectorLogic pSLH
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
 
