VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomBracketGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------------
' Copyright (C) 2011 Intergraph Corporation. All rights reserved.
'
'
' Abstract
'     Custom Bracket GUI Browser
'
' Notes
'
'----------------------------------------------------------------------------------------

Option Explicit

Implements ICustomBracketGUI

Private Const MODULE = "mfCustomBracketGUI.CustomBracketGUI"
Private m_frmGUI As frmCustomBracketGUI
Private m_oCustomGUI As ICustomBracketGUI

Dim m_lOfferId As Long

Private Sub Class_Initialize()
    Set m_frmGUI = New frmCustomBracketGUI
    Set m_oCustomGUI = m_frmGUI
    
    Dim oTrader As Trader
    Set oTrader = New Trader
    
    oTrader.Export Me, TKCustomBracketGUI, "", m_lOfferId
    Set oTrader = Nothing
End Sub

Private Sub Class_Terminate()
    Const METHOD_NAME = "Class_Terminate"
    DEBUG_MSG "Entering " & METHOD_NAME
    
    Dim oTrader As Trader
    Set oTrader = New Trader
    oTrader.Withdraw m_lOfferId
    Set oTrader = Nothing
    
    If Not m_frmGUI Is Nothing Then
        DEBUG_MSG METHOD_NAME & ": frmGUI is nothing, visible=" & m_frmGUI.Visible
'        If m_frmGUI.Visible = True Then
            Unload m_frmGUI
'        End If
    End If
    
    Set m_oCustomGUI = Nothing
    Set m_frmGUI = Nothing
    
    DEBUG_MSG "Exiting " & METHOD_NAME
End Sub

Private Sub ICustomBracketGUI_ApplyChanges()
    m_oCustomGUI.ApplyChanges
End Sub

Private Property Let ICustomBracketGUI_EventHandler(ByVal pEventHandler As stDefinitionGUI.ICustomBracketGUIEventHandler)
    m_oCustomGUI.EventHandler = pEventHandler
End Property

Private Property Let ICustomBracketGUI_ExecuteMode(ByVal bCreateMode As Boolean)
    m_oCustomGUI.ExecuteMode = bCreateMode
End Property

Private Property Let ICustomBracketGUI_FinishEnabled(ByVal bFinishEnabled As Boolean)
    m_oCustomGUI.FinishEnabled = bFinishEnabled
End Property

Private Property Let ICustomBracketGUI_FlipEnabled(ByVal bFlipEnabled As Boolean)
    m_oCustomGUI.FlipEnabled = bFlipEnabled
End Property

Private Sub ICustomBracketGUI_HideBrowser()
    If Not m_frmGUI Is Nothing Then _
        m_frmGUI.Hide
End Sub

Private Property Let ICustomBracketGUI_OverrideRuleStatus(ByVal bChecked As Boolean)
    m_oCustomGUI.OverrideRuleStatus = bChecked
End Property

Private Property Let ICustomBracketGUI_PlaneMethod(ByVal ePlaneDefinitionMethod As stDefinitionGUI.stBktPlaneDefinitionMethod)
    m_oCustomGUI.PlaneMethod = ePlaneDefinitionMethod
End Property

Private Sub ICustomBracketGUI_PopulateOffsets()
    m_oCustomGUI.PopulateOffsets
End Sub

Private Sub ICustomBracketGUI_Reset()
    m_oCustomGUI.Reset
End Sub

Private Sub ICustomBracketGUI_SetUpForNextBracket()
    m_oCustomGUI.SetUpForNextBracket
End Sub

Private Sub ICustomBracketGUI_ShowBrowser()
    If m_frmGUI Is Nothing Then
        Set m_frmGUI = New frmCustomBracketGUI
        Set m_oCustomGUI = m_frmGUI
    End If
    
    m_frmGUI.Show
    
End Sub

Private Sub ICustomBracketGUI_SupportMultiEdit(ByVal bAllSameOverride As Boolean, ByVal bAllSamePlaneMethod As Boolean, ByVal bAllSameSymbol As Boolean)
    m_oCustomGUI.SupportMultiEdit bAllSameOverride, bAllSamePlaneMethod, bAllSameSymbol
End Sub

Private Property Let ICustomBracketGUI_Symbol(ByVal pSymbol As Object)
    m_oCustomGUI.Symbol = pSymbol
End Property

Private Sub ICustomBracketGUI_UpdateSelectedSupports(ByVal supportIndex As Integer, ByVal pSupport As Object)
    m_oCustomGUI.UpdateSelectedSupports supportIndex, pSupport
End Sub

Private Property Let ICustomBracketGUI_UVOffsetsOverridden(ByVal bAreUVOffsetsOverridden As Boolean)
    m_oCustomGUI.UVOffsetsOverridden = bAreUVOffsetsOverridden
End Property
