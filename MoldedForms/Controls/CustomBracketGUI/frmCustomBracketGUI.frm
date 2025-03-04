VERSION 5.00
Object = "{AEC1B8F7-0AA6-11D5-BC47-0800360103C4}#2.0#0"; "ImageViewCtl.ocx"
Object = "{4A06DC35-06ED-474D-BF22-9CD9F034E8C5}#1.1#0"; "CheckBoxView.ocx"
Object = "{10F53B7A-D365-4A44-834E-459EE9262794}#10.0#0"; "mfCustomTBOverride.ocx"
Object = "{53EEF7FB-A96A-45DE-B3B4-469E26676F59}#3.0#0"; "mfCustomPropCtl.ocx"
Object = "{82376281-D667-47D9-9D60-CB0B794E8353}#2.0#0"; "SOParameterGrid.ocx"
Object = "{D9EF0076-779A-11D4-B648-08003604A303}#2.0#0"; "ComboViewCtl.ocx"
Object = "{FB3CD98B-F7F7-4AF2-B355-ECEA06EF04B1}#1.1#0"; "QADIViewCtl.ocx"
Object = "{E1133060-2987-4390-A247-439CDD510B46}#1.0#0"; "StructCatalogPalette.ocx"
Begin VB.Form frmCustomBracketGUI 
   Caption         =   "Define Bracket Properties"
   ClientHeight    =   11430
   ClientLeft      =   165
   ClientTop       =   570
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   11430
   ScaleWidth      =   11925
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   11000
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   52
      Width           =   12700
      Begin VB.Frame Frame 
         Height          =   721
         Index           =   4
         Left            =   158
         TabIndex        =   12
         Top             =   158
         Width           =   5647
         Begin ImageViewCtl.ImageView imvReinforce 
            Height          =   330
            Left            =   3723
            TabIndex        =   21
            ToolTipText     =   "Bracket Reinforcement Type"
            Top             =   248
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
         End
         Begin CheckBoxView.CheckBoxViewCtl chkvwByRule 
            Height          =   255
            Left            =   1538
            TabIndex        =   20
            ToolTipText     =   "Check for rule based reinforcement selection"
            Top             =   185
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
         End
         Begin VB.CommandButton cmdFlip 
            Height          =   315
            Left            =   1103
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   248
            Width           =   345
         End
         Begin VB.CheckBox chkByRule 
            Caption         =   "By Rule"
            Height          =   255
            Left            =   158
            TabIndex        =   17
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Height          =   195
            Index           =   3
            Left            =   4560
            TabIndex        =   29
            Top             =   0
            Width           =   45
         End
         Begin VB.Label lblReinforceByRule 
            Caption         =   "Reinforcement By Rule"
            Height          =   255
            Left            =   1883
            TabIndex        =   22
            Top             =   248
            Width           =   1750
         End
      End
      Begin VB.Frame Frame 
         Height          =   3180
         Index           =   3
         Left            =   6200
         TabIndex        =   10
         Top             =   7109
         Width           =   5650
         Begin VB.Frame Frame 
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   24
            Top             =   2520
            Width           =   5415
            Begin ComboViewCtl.ComboBoxView cbvOffset 
               Height          =   315
               Index           =   1
               Left            =   3360
               TabIndex        =   28
               Top             =   120
               Width           =   1400
               _ExtentX        =   2461
               _ExtentY        =   556
            End
            Begin ComboViewCtl.ComboBoxView cbvOffset 
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   25
               ToolTipText     =   "U Vector Length"
               Top             =   120
               Width           =   1400
               _ExtentX        =   2461
               _ExtentY        =   556
            End
            Begin VB.Label lblVVec 
               Alignment       =   2  'Center
               Caption         =   "V"
               Height          =   255
               Left            =   4800
               TabIndex        =   26
               ToolTipText     =   "V Vector Length"
               Top             =   150
               Width           =   255
            End
            Begin VB.Label lblUVec 
               Alignment       =   2  'Center
               Caption         =   "U"
               Height          =   255
               Left            =   1700
               TabIndex        =   27
               Top             =   150
               Width           =   255
            End
            Begin VB.Label Label 
               Height          =   375
               Index           =   1
               Left            =   2040
               TabIndex        =   30
               Top             =   0
               Width           =   1135
            End
         End
         Begin VB.CheckBox chkOverrideUVOffsets 
            Caption         =   "Override UV Points"
            Enabled         =   0   'False
            Height          =   255
            Left            =   158
            TabIndex        =   19
            Top             =   2250
            Width           =   1695
         End
         Begin mfCustomTBOverride.TBOverride grdTBOverride 
            Height          =   2400
            Left            =   158
            TabIndex        =   11
            ToolTipText     =   "Override Connect Points"
            Top             =   248
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   4233
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Height          =   195
            Index           =   0
            Left            =   2400
            TabIndex        =   32
            Top             =   0
            Width           =   45
         End
      End
      Begin VB.Frame Frame 
         Height          =   3180
         Index           =   1
         Left            =   158
         TabIndex        =   5
         Top             =   7109
         Width           =   5950
         Begin mfCustomPropCtl.CustomPropCtl ctlCustomPropCtl 
            Height          =   2400
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   4233
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Height          =   195
            Index           =   4
            Left            =   4320
            TabIndex        =   31
            Top             =   0
            Width           =   45
         End
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Finish"
         Height          =   315
         Index           =   0
         Left            =   7905
         TabIndex        =   3
         Top             =   10354
         Width           =   1125
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Apply"
         Height          =   315
         Index           =   1
         Left            =   9120
         TabIndex        =   2
         Top             =   10354
         Width           =   1125
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Close"
         Height          =   315
         Index           =   2
         Left            =   10335
         TabIndex        =   1
         Top             =   10354
         Width           =   1125
      End
      Begin VB.Frame Frame 
         Height          =   6050
         Index           =   5
         Left            =   158
         TabIndex        =   6
         Top             =   969
         Visible         =   0   'False
         Width           =   11575
         Begin QADIViewCtl.SmartOccurrenceView ctlSO 
            Height          =   5880
            Left            =   120
            TabIndex        =   23
            Top             =   120
            Width           =   11145
            _ExtentX        =   19659
            _ExtentY        =   10372
         End
      End
      Begin VB.Frame Frame 
         DragMode        =   1  'Automatic
         Height          =   11450
         Index           =   2
         Left            =   158
         TabIndex        =   4
         Top             =   869
         Width           =   11750
         Begin StructCatalogPalette.CatalogPaletteCtl ctlCatalogPalette 
            Height          =   5370
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   9472
         End
         Begin StructCatalogPalette.CatalogPreviewCtl ctlCatalogPreviewCtl 
            Height          =   3660
            Left            =   5940
            TabIndex        =   9
            Tag             =   " "
            Top             =   200
            Width           =   5277
            _ExtentX        =   9313
            _ExtentY        =   6456
         End
         Begin SOParameterGrid.SymbolParameterGrid grdSOSymbolParamGrid 
            Height          =   2052
            Left            =   5820
            TabIndex        =   15
            Top             =   3908
            Width           =   5397
            _ExtentX        =   9525
            _ExtentY        =   3625
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Height          =   195
            Index           =   2
            Left            =   2040
            TabIndex        =   33
            Top             =   0
            Width           =   45
         End
         Begin VB.Label lblSymName 
            Height          =   480
            Left            =   2520
            TabIndex        =   8
            Top             =   5400
            Visible         =   0   'False
            Width           =   3705
         End
         Begin VB.Label lblCategory 
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   2000
         End
      End
      Begin VB.Label lblToggle 
         Caption         =   "Toggle Dir"
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Top             =   2640
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCustomBracketGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------------------------------------------------
' Copyright (C) 2011 Intergraph Corporation. All rights reserved.
'
'
' Abstract
'     Custom Bracket GUI Browser Form
'
' Notes
'
'----------------------------------------------------------------------------------------

Option Explicit

Implements ICustomBracketGUI

' Windows API call/constants
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Const HWND_TOP                  As Long = 0
Private Const HWND_BOTTOM               As Long = 1
Private Const HWND_TOPMOST              As Long = -1
Private Const HWND_NOTOPMOST            As Long = -2
Private Const SWP_NOSIZE                As Long = &H1
Private Const SWP_NOMOVE                As Long = &H2
Private Const SWP_NOACTIVATE            As Long = &H10
Private Const SWP_SHOWWINDOW            As Long = &H40

'************************
' PRIVATE CONSTANTS
'************************
Private Const MODULE = "mfCustomBracketGUI.frmCustomBracketGUI"
Private Const BPS_PALETTE_XML_RELATIVE_PATH = "Xml\Structure\BracketPlateSystemRootNode.xml"
Private Const DEFAULT_HEIGHT = 10200
Private Const DEFAULT_WIDTH = 12060

'************************
' PRIVATE VARIABLES
'************************
Private m_oSymbol As Object
Private m_bCreateMode As Boolean
Private m_bFormActivated As Boolean
Private m_PlaneMethod As PlaneDefinitionMethodEnum
Private m_bNotifyRibbonBar As Boolean
Private m_sPaletteMsg1 As String
Private m_sPaletteMsg2 As String
Private m_sParameterMsg As String
Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single

Private m_bFormFirstLoading As Boolean
Private m_bPaletteStarted As Boolean
Private m_bParameterGridHasSO As Boolean
Private m_bProcessingFinish As Boolean
Private m_bSelectedFromPalette As Boolean
Private m_bProcessingInternalCompute As Boolean

Private m_oEventHandler As ICustomBracketGUIEventHandler
Private m_oSupports As IJElements
Private m_bMultiEditMode As Boolean
Private m_bAllSameOverride As Boolean
Private m_bAllSameSymbol As Boolean
Private m_bHaveClicked As Boolean

'************************
' PRIVATE ENUMERATORS
'************************

Private Enum ButtonIndex
    FinishBtn = 0
    ApplyBtn = 1
    CloseBtn = 2
End Enum

Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Private Enum eControlContainerFrame
    ccfMain = 0
    ccfProperties = 1
    ccfPalette = 2
    ccfUVOffsets = 3
    ccfStepControls = 4
    ccfSmartOccurrence = 5
    ccfVectorLengths = 6
End Enum

Private Sub chkByRule_Click()
    Const METHOD = "chkByRule_Click"
    DEBUG_MSG "Entering " & METHOD

        
    ' when switching to/from by rule, even if there is an SO
    ' existing, it is out of date as soon as the method is
    ' switched so then the parameter grid is out of date.
    ' For switching to by rule, Compute will update the SO and
    ' then set up the symbol and associated correct parameter grid.
    ' For switching to pre-selected, the ribbon bar synchronization
    ' will reset the symbol if there is a remembered symbol.
    'Notify Ribbon Bar through the control, if this event wasn't triggered by the Ribbon Bar
    m_bParameterGridHasSO = False
    ClearParameterGrid
    
    If chkByRule.Value = vbChecked Then
        If m_bNotifyRibbonBar = True Then
           m_oEventHandler.SymbolByRule True
           m_oEventHandler.Accept
        End If
        ShowSmartOccurrenceView True
    Else
        ShowSmartOccurrenceView False
        
        ' when resetting back to pre-selected, verify that the symbol stored in the custom GUI
        ' and the symbol displayed by the icon browser are in synch
        If Not m_oSymbol Is Nothing And m_bPaletteStarted = True Then
            ctlCatalogPalette.SelectedCatalogObject = m_oSymbol
        End If
        
        If m_bNotifyRibbonBar = True Then m_oEventHandler.SymbolByRule False
    End If
    
    
            
    DEBUG_MSG "Exiting " & METHOD
    Exit Sub

End Sub

Private Sub chkOverrideUVOffsets_Click()
    Const METHOD = "chkOverrideUVOffsets_Click"
    DEBUG_MSG "Entering " & METHOD
    
    grdTBOverride.CheckEnabledOverrideBoxes chkOverrideUVOffsets.Value
    
    'Notify Ribbon Bar through the control, if this event wasn't triggered by the Ribbon Bar
    If m_bNotifyRibbonBar = True Then _
        m_oEventHandler.OffsetOverridden
    
    DEBUG_MSG "Exiting " & METHOD

End Sub

Private Sub cmdButton_Click(Index As Integer)
    Const METHOD = "cmdButton_Click"
    DEBUG_MSG "Entering " & METHOD & ": index = " & Index
    On Error GoTo ErrorHandler

    Select Case Index
           
        Case ButtonIndex.ApplyBtn
            ApplyBracketChanges True
            
        Case ButtonIndex.CloseBtn
            Me.Hide
            
        Case ButtonIndex.FinishBtn
            m_bProcessingFinish = True
            ApplyBracketChanges
            m_oEventHandler.Finish
            m_bProcessingFinish = False ' just in case Reset hasn't already done this
       
            
    End Select
    DEBUG_MSG "Exiting " & METHOD
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Sub cmdFlip_Click()
    Const METHOD = "cmdButton_Click"
    DEBUG_MSG "Entering " & METHOD
    
    m_oEventHandler.Flip
    
    DEBUG_MSG "Exiting " & METHOD
End Sub

Private Sub ctlCatalogPalette_ItemSelected()
    Const METHOD = "ctlCatalogPalette_ItemSelected"
    DEBUG_MSG "Entering " & METHOD
    On Error Resume Next

    Dim oPartWrapper As Object
    If ctlCatalogPalette.SelectedCatalogObject Is Nothing Then
        MsgBox m_sPaletteMsg2
    End If
    Set oPartWrapper = ctlCatalogPalette.SelectedCatalogObject
    If oPartWrapper Is Nothing Then
        MsgBox m_sPaletteMsg1
        Exit Sub
    End If
    
    m_bHaveClicked = True
    Dim oSmartItem As IJSmartItem

    If Not oPartWrapper Is Nothing Then
        Dim oObj As Object
        Set oObj = ctlCatalogPalette.SelectedCatalogObject

        If TypeOf oObj Is IJSmartItem Then
           
            Set oSmartItem = oObj
            ctlCatalogPreviewCtl.Visible = True
            
            'Only care if selected from palette initially,
            'so can compute symbol without getting in a loop w/ RB
            m_bSelectedFromPalette = True
            UpdateSmartItemDisplay oSmartItem
            m_bSelectedFromPalette = False
            
            Set oSmartItem = Nothing
            Set oObj = Nothing
            Exit Sub
        End If
        
        Set oObj = Nothing
    End If
    
    DEBUG_MSG "Exiting " & METHOD
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub ctlSO_NewSmartItemSelection(oSmartItem As IJSmartItem)
    Const METHOD = "ctlSO_NewSmartItemSelection"
    DEBUG_MSG "Entering " & METHOD
    On Error GoTo ErrorHandler

    If Not oSmartItem Is Nothing And m_bSelectedFromPalette = False Then
        DEBUG_MSG METHOD & ": Smart Item Selected=" & oSmartItem.Name
        m_bSelectedFromPalette = True
        If (m_bMultiEditMode = True) Then
            UpdateSmartItemDisplay oSmartItem
        End If
                
        ApplyBracketChanges True
        m_bSelectedFromPalette = False
    End If
    
    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub Form_Activate()

    Const METHOD = "Form_Activate"
    On Error GoTo ErrorHandler
    DEBUG_MSG "Entering " & METHOD
    
    LocalizeForm
    
    ctlSO.Holder = PLATE_ContourSmartItemHolder
    
    ctlCustomPropCtl.ShowReinforcementTab ' Allow reinforcement tab when appropriate
    ctlCustomPropCtl.Width = 5775
    ctlCustomPropCtl.Height = 2590
    m_bFormActivated = True
    
    DEBUG_MSG "Exiting " & METHOD
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Sub Form_Initialize()

    Const METHOD = "Form_Initialize"
    DEBUG_SOURCE "Custom Bracket GUI Form"
    DEBUG_MSG "Entering " & METHOD
    DEBUG_DEEP_BEGIN
    
    m_bPaletteStarted = False
    m_bHaveClicked = False
    m_bCreateMode = True
    m_bFormActivated = False
    m_bParameterGridHasSO = False
    m_bProcessingFinish = False
    m_bSelectedFromPalette = False
    m_bProcessingInternalCompute = False
    m_bMultiEditMode = False
    Set m_oSupports = New IMSElements.Elements
    
    DEBUG_MSG "Exiting " & METHOD
    Exit Sub
End Sub

Private Sub Form_Load()
    Const METHOD = "Form_Load"
    DEBUG_MSG "Entering " & METHOD
    On Error GoTo ErrorHandler
    
    m_bFormFirstLoading = False
    m_bNotifyRibbonBar = True
    
    DEBUG_MSG METHOD & ": setting XML on SOCtl"
    ctlSO.ConfigurationXMLForSmartItemResults = BPS_PALETTE_XML_RELATIVE_PATH
    
    imvReinforce.SetComboDropDownWidth 200
    imvReinforce.Holder = PLATE_BktReinforcementTypeHolder
    chkvwByRule.Holder = PLATE_BktOverrideTypeHolder
    With cbvOffset(0)
        .UseRecentValues = True
        .LastXValue = 10
        .Holder = PLATE_BracketOffset1PropertyHolder
        .Enabled = True
        .AllowUnListedValues = True
    End With
    With cbvOffset(1)
        .UseRecentValues = True
        .LastXValue = 10
        .Holder = PLATE_BracketOffset2PropertyHolder
        .Enabled = True
        .AllowUnListedValues = True
    End With

    ' wait to initialize palette here in case browser is closed using 'X'
    ' which unloads the form and terminates the catalog palette control
    ' so that it will need to be re-initialized
    ' NOTE: Since they can switch back and forth, always display assuming
    ' caching/performance issues are worked out for initial display
    Call InitializePalette
    m_bPaletteStarted = True

    If Not m_oSymbol Is Nothing Then _
        ICustomBracketGUI_Symbol = m_oSymbol

    If m_PlaneMethod = ByElements And m_bMultiEditMode = False Then
        'Create Value manager Key
        Dim oTrader As New Trader
        Dim oValueMgr As IJValueMgr
        Set oValueMgr = oTrader.Service(TKValueMgr, "")
        If Not oValueMgr Is Nothing Then
           If oValueMgr.ValidKey("TB_LoadPoint_Override") = False Then oValueMgr.Add "TB_LoadPoint_Override", True
        End If
        Set oTrader = Nothing
        Set oValueMgr = Nothing
        Frame(ccfUVOffsets).Visible = True
        grdTBOverride.PopulateGrid
    Else
        Frame(ccfUVOffsets).Visible = False
    End If
        
    Dim lTop As Long
    Dim lLeft As Long
    Dim lHeight As Long
    Dim lWidth As Long
    
    GetPreferencesWithHeightWidth MODULE, lTop, lLeft, lHeight, lWidth
    DEBUG_MSG METHOD & ": Preferences top=" & lTop & "; left=" & lLeft & "; height=" & lHeight & "; width=" & lWidth
    GetCustomBracketParameterPreferences
    AdjustCommandButtons
    
    SaveSizes ' need to save these now since prefs might resize the default form
    
    Me.Top = lTop
    Me.Left = lLeft
    If lHeight <> 0 And lWidth <> 0 Then
        Me.Height = lHeight
        Me.Width = lWidth
    End If
    
    SetTopMostWindow Me, True ' to keep the custom GUI from being hidden
    
    DEBUG_MSG "Exiting " & METHOD
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Sub Form_Resize()
    Const METHOD = "Form_Resize"
    On Error GoTo ErrorHandler
  
    'Don't adjust when form is first loading
    If m_bFormFirstLoading = False Then
        m_bFormFirstLoading = True
        Exit Sub
    End If
    
    ResizeControls
    AdjustCommandButtons

    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Sub Form_Terminate()

    Const METHOD = "Form_Terminate"
    DEBUG_MSG "Entering " & METHOD
    
    Set m_oEventHandler = Nothing
    Set m_oSymbol = Nothing
    m_oSupports.Clear
    Set m_oSupports = Nothing
    
    DEBUG_MSG "Exiting " & METHOD
    DEBUG_DEEP_END
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Const METHOD = "Form_Unload"
    DEBUG_MSG "Entering " & METHOD
    
    ctlSO.StopControl ' this releases the references so the control terminates correctly
    
    ' unloading terminates the palette, so mark it back so it will be restarted
    Set m_oSymbol = Nothing
    m_bPaletteStarted = False
    m_bFormActivated = False
    m_bParameterGridHasSO = False
    
    StorePreferences MODULE, Me.Top, Me.Left, Me.Height, Me.Width ' remember where the form was last displayed
    SaveCustomBracketParameterPreferences
    
    SetTopMostWindow Me, False
    
    DEBUG_MSG "Exiting " & METHOD
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Sub ICustomBracketGUI_ApplyChanges()
    m_bNotifyRibbonBar = False
    ApplyBracketChanges False
    m_bNotifyRibbonBar = True
End Sub

Private Property Let ICustomBracketGUI_EventHandler(ByVal pEventHandler As stDefinitionGUI.ICustomBracketGUIEventHandler)
    Const METHOD = "ICustomBracketGUI_EventHandler"
    DEBUG_MSG "Entering " & METHOD
    
    Set m_oEventHandler = Nothing
    
    If Not pEventHandler Is Nothing Then _
        Set m_oEventHandler = pEventHandler
    
    DEBUG_MSG "Exiting " & METHOD

End Property

Private Property Let ICustomBracketGUI_ExecuteMode(ByVal bCreateMode As Boolean)
    Const METHOD = "ICustomBracketGUI_ExecuteMode"
    DEBUG_MSG "Entering " & METHOD & ": bCreateMode=" & bCreateMode
    
    m_bCreateMode = bCreateMode
    
    DEBUG_MSG "Exiting " & METHOD
    Exit Property

End Property

Private Property Let ICustomBracketGUI_FinishEnabled(ByVal bFinishEnabled As Boolean)
    Const METHOD = "ICustomBracketGUI_FinishEnabled"
    DEBUG_MSG "Entering " & METHOD & ": bFinishEnabled=" & bFinishEnabled
    
    cmdButton(ButtonIndex.FinishBtn).Enabled = bFinishEnabled
    
    DEBUG_MSG "Exiting " & METHOD
    Exit Property
End Property

Private Property Let ICustomBracketGUI_FlipEnabled(ByVal bFlipEnabled As Boolean)
    cmdFlip.Enabled = bFlipEnabled
End Property

Private Sub ICustomBracketGUI_HideBrowser()
    'Nothing to do for this method in the form
End Sub

Private Property Let ICustomBracketGUI_OverrideRuleStatus(ByVal bChecked As Boolean)
    'The ribbon bar is notifying us, so just set the value to synch with ribbon bar, but
    'don't raise an event to the ribbon bar
    m_bNotifyRibbonBar = False
    chkByRule.Value = IIf(bChecked, vbChecked, vbUnchecked)
    
    If bChecked = True Then
        ShowSmartOccurrenceView True
    Else
        ShowSmartOccurrenceView False
    End If
    m_bNotifyRibbonBar = True

End Property

Private Property Let ICustomBracketGUI_PlaneMethod(ByVal ePlaneDefinitionMethod As stDefinitionGUI.stBktPlaneDefinitionMethod)
    Const METHOD = "ICustomBracketGUI_PlaneMethod"
    DEBUG_MSG "Entering " & METHOD & ": ePlaneDefinitionMethod=" & ePlaneDefinitionMethod
    
    'Don't enable override in create mode until bracket is computed
    'If m_bFormActivated = True Then
        chkOverrideUVOffsets.Enabled = (ePlaneDefinitionMethod = ByElements) 'And (m_bCreateMode = False)
        Frame(ccfUVOffsets).Visible = (ePlaneDefinitionMethod = ByElements)
        If ePlaneDefinitionMethod = ByElements And m_bMultiEditMode = False Then
           'Create Value manager Key
           Dim oTrader As New Trader
           Dim oValueMgr As IJValueMgr
           Set oValueMgr = oTrader.Service(TKValueMgr, "")
           If Not oValueMgr Is Nothing Then
              If oValueMgr.ValidKey("TB_LoadPoint_Override") = False Then oValueMgr.Add "TB_LoadPoint_Override", True
           End If
           Set oTrader = Nothing
           Set oValueMgr = Nothing
           Frame(ccfUVOffsets).Visible = True
           grdTBOverride.PopulateGrid
        End If
  '  End If
    m_PlaneMethod = ePlaneDefinitionMethod

    DEBUG_MSG "Exiting " & METHOD
End Property

Private Sub ICustomBracketGUI_PopulateOffsets()
    Const METHOD = "ICustomBracketGUI_PopulateOffsets"
    DEBUG_MSG "Entering " & METHOD
    On Error GoTo ErrorHandler
    If m_bFormActivated = True And m_PlaneMethod = ByElements Then
       grdTBOverride.PopulateGrid
    End If
    
    DEBUG_MSG "Exiting " & METHOD
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub ICustomBracketGUI_Reset()
    Const METHOD = "ICustomBracketGUI_Reset"
    DEBUG_MSG "Entering " & METHOD
    
    Set m_oSymbol = Nothing
    ctlCatalogPalette.CleanUp
    
    ' All we need to do is turn off this, and old SO
    ' will not be used as ApplyParameters will not be
    ' done again until a new SO has been set and this
    ' flag has been re-activated.
    m_bParameterGridHasSO = False
    If chkByRule.Value = vbChecked Then
        ClearSmartOccurrenceControl
        ctlSO.Visible = False ' reset to hide QADIView until have SO
    Else
        'in order to use preset values, the grid must be
        'properly reset
        ClearParameterGrid
    End If

    m_bProcessingFinish = False ' Reset is last step before actual Finish, so can clear this flag now
    
    DEBUG_MSG "Exiting " & METHOD
End Sub

Private Sub ICustomBracketGUI_SetUpForNextBracket()
    Const METHOD = "ICustomBracketGUI_SetUpForNextBracket"
    DEBUG_MSG "Entering " & METHOD
          
    cmdButton(ButtonIndex.FinishBtn).Enabled = False
    
    DEBUG_MSG "Exiting " & METHOD

End Sub

Private Sub ICustomBracketGUI_ShowBrowser()
    'Nothing to do for this method in the form
End Sub

Private Sub ICustomBracketGUI_SupportMultiEdit(ByVal bAllSameOverride As Boolean, ByVal bAllSamePlaneMethod As Boolean, ByVal bAllSameSymbol As Boolean)
    Const METHOD = "ICustomBracketGUI_SupportMultiEdit"
    DEBUG_MSG "Entering " & METHOD & ": bAllSameOverride=" & bAllSameOverride & "; bAllSamePlaneMethod=" & bAllSamePlaneMethod & "; bAllSameSymbol=" & bAllSameSymbol

    m_bMultiEditMode = True
    m_bAllSameOverride = bAllSameOverride
    m_bAllSameSymbol = bAllSameSymbol
    Frame(ccfUVOffsets).Visible = False ' offsets not currently support in multi-edit mode but may be at some point

    If m_bAllSameOverride = False Or (chkByRule.Value = vbUnchecked And bAllSameSymbol = False) Then
        ShowSmartOccurrenceView False ' by default show the palette for selecting a symbol in mixed mode
    ElseIf chkByRule.Value = vbChecked Then
        ShowSmartOccurrenceView True
    End If

    ' since this is only used to show override U/V offsets grid, just set so won't be shown if mixed
    If bAllSamePlaneMethod = False Then
        m_PlaneMethod = Coincident
    End If
    
    DEBUG_MSG "Exiting " & METHOD
End Sub

Private Property Let ICustomBracketGUI_Symbol(ByVal pSymbol As Object)
    Const METHOD = "ICustomBracketGUI_Symbol"
    DEBUG_MSG "Entering " & METHOD
    On Error GoTo ErrorHandler
    
    If m_bProcessingFinish = True Then Exit Property
    
    Dim oSmartItem As IJSmartItem
               
    m_bNotifyRibbonBar = False
    Set oSmartItem = pSymbol
    If Not oSmartItem Is Nothing Then
        On Error Resume Next
        If m_bPaletteStarted = True And chkByRule.Value = vbUnchecked Then
            ctlCatalogPalette.SelectedCatalogObject = oSmartItem
        Else
            ' for pre-selected with palette, setting the palette symbol will
            ' trigger updating the display, so only do here if symbol
            ' not being set on the palette
            UpdateSmartItemDisplay oSmartItem
        End If
        Set oSmartItem = Nothing
    End If
    
    chkOverrideUVOffsets.Enabled = (m_PlaneMethod = ByElements) ' Why is this done here?
    m_bNotifyRibbonBar = True
    
    DEBUG_MSG "Exiting " & METHOD
    Exit Property
    
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD

End Property

Private Sub ICustomBracketGUI_UpdateSelectedSupports(ByVal supportIndex As Integer, ByVal pSupport As Object)
    Const METHOD = "ICustomBracketGUI_UpdateSelectedSupports"
    DEBUG_MSG "Entering " & METHOD
    
    Dim oOldSupport As Object
    
    ' turn off error handling so checking for old support that may not be there does not throw an error
    On Error Resume Next
    If Not pSupport Is Nothing Then
        Set oOldSupport = m_oSupports.Item(supportIndex)
        If Not oOldSupport Is Nothing Then _
            m_oSupports.Remove supportIndex
            
        m_oSupports.Add pSupport, supportIndex
    Else
        m_oSupports.Remove supportIndex
    End If
    On Error GoTo ErrorHandler
    
    ' handle any repurcussions of adjusting supports here
    
    DEBUG_MSG "Exiting " & METHOD
    Exit Sub
    
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Property Let ICustomBracketGUI_UVOffsetsOverridden(ByVal bAreUVOffsetsOverridden As Boolean)
    Const METHOD = "ICustomBracketGUI_UVOffsetsOverridden"
    DEBUG_MSG "Entering " & METHOD
    
    'The ribbon bar is notifying us, so just set the value to synch with ribbon bar, but
    'don't raise an event to the ribbon bar
    m_bNotifyRibbonBar = False
    chkOverrideUVOffsets.Value = IIf(bAreUVOffsetsOverridden = True, vbChecked, vbUnchecked)
    m_bNotifyRibbonBar = True

    DEBUG_MSG "Exiting " & METHOD
End Property

Private Sub LocalizeForm()
    Const METHOD = "LocalizeForm"
    DEBUG_MSG "Entering " & METHOD
    On Error GoTo ErrorHandler
    
    Dim oLocalizer As IJLocalizer
    
    Set oLocalizer = GetLocalizer("\Resource\")
    
    cmdFlip.Picture = oLocalizer.GetIcon(CLng(IDI_FLIP))
'    Picture_No(0).Picture = oLocalizer.GetIcon(CLng(IDI_INTL_NO))
'    Picture_No(1).Picture = oLocalizer.GetIcon(CLng(IDI_INTL_NO))
    Me.Icon = oLocalizer.GetIcon(CLng(IDI_CUSTOM_GUI))
    m_sPaletteMsg1 = oLocalizer.GetString(IDS_NO_PART_FROM_PALETTE, "No part returned from Palette")
    m_sPaletteMsg2 = oLocalizer.GetString(IDS_NO_INFO_FROM_PALETTE, "Nothing returned from Palette for selected catalog item")
    m_sParameterMsg = oLocalizer.GetString(IDS_OVERRIDE_INVALID, "Override Parameter data is invalid")
    Me.Caption = oLocalizer.GetString(IDS_FORM_TITLE_BPS, "Define Bracket Plate System Properties")
    
    cmdButton(ButtonIndex.ApplyBtn).Caption = oLocalizer.GetString(IDS_APPLY, "Apply")
    cmdButton(ButtonIndex.ApplyBtn).ToolTipText = oLocalizer.GetString(IDS_APPLY_TIP, "Apply Changes")
    cmdButton(ButtonIndex.CloseBtn).Caption = oLocalizer.GetString(IDS_CLOSE, "Close")
    cmdButton(ButtonIndex.CloseBtn).ToolTipText = oLocalizer.GetString(IDS_CLOSE_TIP, "Skip Changes and Close")
    cmdButton(ButtonIndex.FinishBtn).Caption = oLocalizer.GetString(IDS_FINISH, "Finish")
    cmdButton(ButtonIndex.FinishBtn).ToolTipText = oLocalizer.GetString(IDS_FINISH_TIP, "Finish Changes and Commit to Database")
    
    chkOverrideUVOffsets.Caption = oLocalizer.GetString(IDS_OVERRIDE_UV, "Override UV Points")
    chkOverrideUVOffsets.ToolTipText = chkOverrideUVOffsets.Caption
    lblToggle.Caption = oLocalizer.GetString(IDS_TOGGLE_DIR, "Toggle Dir")
    chkByRule.Caption = oLocalizer.GetString(IDS_BY_RULE, "By Rule")
    chkByRule.ToolTipText = oLocalizer.GetString(IDS_RULE_TIP, "Check for rule based bracket selection")
'    lbl(2).Caption = oLocalizer.GetString(IDS_SYMBOL, "Symbol")
'    lbl(3).Caption = oLocalizer.GetString(IDS_FLANGE, "Flange")
'    lbl(4).Caption = oLocalizer.GetString(IDS_BK_STIFF, "BkStiff")
    cmdFlip.ToolTipText = oLocalizer.GetString(IDS_TOGGLE_DIR, "Toggle Dir")
    Label(0).Caption = oLocalizer.GetString(IDS_OFFSETS, "Offsets")
    Label(1).Caption = oLocalizer.GetString(IDS_VECTOR_LENGTH, "Vector Lengths")
    Label(2).Caption = oLocalizer.GetString(IDS_PREVIEW, "Preview"): Label(2).ZOrder 0
    Label(3).Caption = oLocalizer.GetString(IDS_RULE_SETTINGS, "Rule Settings")
    Label(4).Caption = oLocalizer.GetString(IDS_BRACKET_PROPERTIES, "Bracket Properties"): Label(4).ZOrder 0
'    Label(5).Caption = oLocalizer.GetString(IDS_PARAMETERS, "Parameters"): Label(5).ZOrder 0
    Set oLocalizer = Nothing
    
    DEBUG_MSG "Exiting " & METHOD
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Sub UpdateSmartItemDisplay(oSmartItem As IJSmartItem)
    Const METHOD = "UpdateSmartItemDisplay"
    On Error Resume Next
    DEBUG_MSG "Entering " & METHOD
    DEBUG_MSG METHOD & ": Updating to Smart Item " & oSmartItem.Name
    
    If Not TypeOf oSmartItem.SymbolDefinitionDef Is IJDSymbolDefinition Then
        'This occurs if unable to link to symbol file
        MsgBox m_sPaletteMsg2
        Exit Sub
    ElseIf m_bProcessingInternalCompute = True Then
        'if processing internal compute, no need to process the update right now
        Exit Sub
    End If
    
    Set m_oSymbol = oSmartItem
    
    'Get property holder information to populate the parameters
    'The holder supports both IJPropertyHolder and IJSmartItemPropertyHolder
    'Need to get a smart item (oSmartItem) to pass to UpdateGridWithRefDataAttributes
    'in order to populate the array m_arrOverrideRuleIndex in the control so that
    'scrolling works correctly.
    
    Dim colSOs As Collection
    
    If FoundSmartOccurrencesFromPropertyHolder(colSOs) = True Then
    
        'If the SO has previously been computed (or on Modify), synchronize
        'with the ribbon bar to match behavior of selecting from RB directly.
        If m_bSelectedFromPalette = True And m_bProcessingInternalCompute = False And m_bNotifyRibbonBar = True Then
            m_bProcessingInternalCompute = True
            m_oEventHandler.SymbolSelected m_oSymbol
            m_bProcessingInternalCompute = False
            
            'need to get the new SO now
            If FoundSmartOccurrencesFromPropertyHolder(colSOs) = False Then
                'since there was an SO before, this should not fail but if it
                'does assume other inputs (such as plane or support) removed
                'and just skip remaining code that is for when SO exists
                'and reset parameters to catalog definition only or clear SO
                Set colSOs = Nothing
                If chkByRule.Value = vbChecked Then
                    ClearSmartOccurrenceControl
                Else
                    Set grdSOSymbolParamGrid.CatalogDefinition = oSmartItem
                    grdSOSymbolParamGrid.UpdateGridWithRefDataAttributes oSmartItem
                End If
            End If
        End If
        
        If Not colSOs Is Nothing Then
            If chkByRule.Value = vbChecked Then
                ctlSO.Visible = True
                grdSOSymbolParamGrid.Visible = False
                
                ' do not want to update if this was triggered from the embedded palette
'                If m_bSelectedFromPalette = False Then
                    ctlSO.SmartOccurrence = colSOs
            Else
                grdSOSymbolParamGrid.Visible = True
                Set grdSOSymbolParamGrid.SymbolOcc = colSOs

                'update to set read/write columns
                grdSOSymbolParamGrid.UpdateGridWithRefDataAttributes oSmartItem
            End If
            m_bParameterGridHasSO = True
            Set colSOs = Nothing
    
            If m_PlaneMethod = ByElements And m_bMultiEditMode = False Then _
                grdTBOverride.PopulateGrid
        
        End If
    ElseIf chkByRule.Value = vbUnchecked Then ' pre-selected but no SO yet computed
    
        'get the symbol definition from the catalog item
        DEBUG_MSG METHOD & "Setting catalog definition on parameter grid control"
        Set grdSOSymbolParamGrid.CatalogDefinition = oSmartItem
        grdSOSymbolParamGrid.UpdateGridWithRefDataAttributes oSmartItem

        ' if pre-selected but no SO yet, do initial compute as soon as item is changed (is this a performance issue?)
'        If cmdButton(ButtonIndex.FinishBtn).Enabled = True Then
'            m_oEventHandler.SymbolSelected m_oSymbol
'            If FoundSmartOccurrencesFromPropertyHolder(colSOs) = True Then
'
'                Set grdSOSymbolParamGrid.SymbolOcc = colSOs
'                m_bParameterGridHasSO = True
'                Set colSOs = Nothing
'
'                'update to set read/write columns
'                grdSOSymbolParamGrid.UpdateGridWithRefDataAttributes oSmartItem
'
'                If m_PlaneMethod = ByElements Then _
'                    grdTBOverride.PopulateGrid
'            End If
'        End If
    End If

    DEBUG_MSG "Exiting " & METHOD
    Exit Sub
End Sub

'Private Function SymbolSupportsReinforcement(oSymbol As Object) As Boolean
'
'    Dim oBracketUtils As IJPlateAttributes
'    SymbolSupportsReinforcement = False
'    On Error Resume Next
'    Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
'
'    If oBracketUtils.IsReinforcementValid(oSymbol, "BucklingStiffener") = False And _
'       oBracketUtils.IsReinforcementValid(oSymbol, "EdgeReinforcement") = False Then
'       SymbolSupportsReinforcement = False
'    Else
'       SymbolSupportsReinforcement = True
'    End If
'    Set oBracketUtils = Nothing
'    Exit Function
'
'End Function

'****************************************************************************************************
' Routine: InitializePalette
'
' Abstract: Set up tabs for Palette browser based on paths to MF catalog brackets
'
' Description:
'   Parameters:   None
'****************************************************************************************************
Private Sub InitializePalette()
    Const METHOD = "InitializePalette"
    DEBUG_MSG "Entering " & METHOD
    On Error GoTo ErrorHandler
    
    With ctlCatalogPalette
        .DefaultIcon = eNoPic ' this must be done before the palette is started
        .ConfigurationXML = BPS_PALETTE_XML_RELATIVE_PATH
        .Start
        .RaiseEventOnSCO = True
    End With
    
    With ctlCatalogPreviewCtl
        .SetBroadCaster ctlCatalogPalette
        .HorizontalSBVisibleIfNeeded True
        .VerticalSBVisibleIfNeeded True
    End With
    
    DEBUG_MSG "Exiting " & METHOD
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

'****************************************************************************************************
' Routine: ResizeControls
'
' Abstract: When this Formis resized, all constituent controls need to resize themselves to fit
'
' Description:
'   Parameters:   None.  All controls are checked
'****************************************************************************************************
Private Sub ResizeControls()

    On Error GoTo ErrorHandler
    Dim i As Integer
    Dim ctl As Control
    Dim x_scale As Single
    Dim y_scale As Single
    
    ' Don't bother if we are minimized.
    If WindowState = vbMinimized Then Exit Sub
    
    ' Get the form's current scale factors.
    x_scale = ScaleWidth / m_FormWid
    y_scale = ScaleHeight / m_FormHgt
    
    ' Position the controls.
    i = 1
    For Each ctl In Controls
        With m_ControlPositions(i)
            If TypeOf ctl Is Line Then
                ctl.X1 = x_scale * .Left
                ctl.Y1 = y_scale * .Top
                ctl.X2 = ctl.X1 + x_scale * .Width
                ctl.Y2 = ctl.Y1 + y_scale * .Height
            Else
                ' don't scale size of command buttons, these are a std size
                ' also, do not scale anything in Frame(ccfStepControls)
                ' don't resize labels, since they are set to autoresize
                If (Not TypeOf ctl Is CommandButton) And (DoNotResizeControl(ctl) = False) Then
                   ctl.Left = x_scale * .Left
                   ctl.Top = y_scale * .Top
                   ctl.Width = x_scale * .Width
                   If Not (TypeOf ctl Is ComboBox) Then
                      ' Cannot change height of ComboBoxes.
                      ctl.Height = y_scale * .Height
                   End If
                End If
            End If
        End With
        i = i + 1
    Next ctl
     'Re-Center labels (they automatically resized).
    For i = 0 To 4
       Label_Change (i)
    Next i
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError "frmBracket", "ResizeControls"
End Sub

'****************************************************************************************************
' Routine: SaveSizes
'
' Abstract: Save the initial sizes of all controls for reference when Form is resized
'
' Description:
'   Parameters:   None
'****************************************************************************************************
Private Sub SaveSizes()
   Dim i As Integer
   Dim ctl As Control

   On Error GoTo ErrorHandler
   ' Save the controls' positions and sizes.
   ReDim m_ControlPositions(1 To Controls.count)
   i = 1
   For Each ctl In Controls
       With m_ControlPositions(i)
           If TypeOf ctl Is Line Then
               .Left = ctl.X1
               .Top = ctl.Y1
               .Width = ctl.X2 - ctl.X1
               .Height = ctl.Y2 - ctl.Y1
           Else
               .Left = ctl.Left
               .Top = ctl.Top
               .Width = ctl.Width
               .Height = ctl.Height
               On Error Resume Next
               .FontSize = ctl.Font.Size
               On Error GoTo 0
           End If
       End With
       i = i + 1
   Next ctl

   ' Save the form's size.
   m_FormWid = ScaleWidth
   m_FormHgt = ScaleHeight
   Exit Sub
ErrorHandler:
    ReportUnanticipatedError "frmBracket", "SaveSizes"
End Sub

Private Sub SetTopMostWindow(oForm As Form, bMakeTopmost As Boolean)
    On Error Resume Next

    Dim lParm As Long
    
    lParm = IIf(bMakeTopmost, HWND_TOPMOST, HWND_NOTOPMOST)

    'NOTE: In order to ensure that error messages can still be seen,
    'the flags for SWP_NOACTIVATE and SWP_SHOWWINDOW were removed.
    'This may mean there are cases where the custom GUI form could
    'be hidden, but if a modal error message is hidden it effectively
    'locks up the Marine application.
    SetWindowPos oForm.hwnd, _
                 lParm, _
                 0, _
                 0, _
                 0, _
                 0, _
                 (SWP_NOMOVE Or SWP_NOSIZE)

   Exit Sub
End Sub

Private Sub ApplyBracketChanges(Optional bComputeNow As Boolean = False)
    Const METHOD = "ApplyBracketChanges"
    DEBUG_MSG "Entering " & METHOD
    Dim colSOs As Collection
    
    If Not m_oSymbol Is Nothing Then
        If m_PlaneMethod = ByElements And m_bMultiEditMode = False And grdTBOverride.HasDataChanged = True Then
            grdTBOverride.WriteToValueManager
        End If
        
        If m_bParameterGridHasSO Then
            ' skip if doing compute, as event handler accept is supposed to apply changes itself
            If chkByRule.Value = vbChecked And bComputeNow = False Then
                ctlSO.ApplyChanges
            Else
                grdSOSymbolParamGrid.ApplyParametersToSymbol
            End If
            
            If bComputeNow = True Then
                m_oEventHandler.Accept
                
                If chkByRule.Value = vbChecked Then
                    ctlSO.UpdateSelectedSmartItem m_oSymbol, True
                End If
            End If
        ElseIf chkByRule.Value = vbUnchecked Then ' only process preset for pre-selected
            Dim oSmartItem As IJSmartItem
            Set oSmartItem = m_oSymbol
            grdSOSymbolParamGrid.WriteToValueMgr oSmartItem.Name
            Set oSmartItem = Nothing
    
            DEBUG_MSG METHOD & ": Setting symbol selected (first compute)"
            m_bProcessingInternalCompute = True
            m_oEventHandler.SymbolSelected m_oSymbol
            m_bProcessingInternalCompute = False
    
            If FoundSmartOccurrencesFromPropertyHolder(colSOs) = True Then
                Set grdSOSymbolParamGrid.SymbolOcc = colSOs
                m_bParameterGridHasSO = True
                Set colSOs = Nothing
                grdSOSymbolParamGrid.UpdateGridWithPresetParameters m_oSymbol
                grdSOSymbolParamGrid.ApplyParametersToSymbol
                If bComputeNow = True Then _
                    m_oEventHandler.Accept
            End If
        Else
            If bComputeNow = True Then m_oEventHandler.Accept
        End If
    ElseIf bComputeNow = True Then
        m_oEventHandler.Accept
    End If

    DEBUG_MSG "Exiting " & METHOD
    Exit Sub
    
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Sub ClearParameterGrid()
    Const METHOD = "ClearParameterGrid"
    DEBUG_MSG "Entering " & METHOD
    On Error Resume Next
    
    Dim dummyColl As Collection
    Set dummyColl = New Collection
    Set grdSOSymbolParamGrid.SymbolOcc = dummyColl
    Set dummyColl = Nothing
    
    DEBUG_MSG "Exiting " & METHOD
End Sub

Private Sub ClearSmartOccurrenceControl()
    Const METHOD = "ClearSmartOcurrenceControl"
    DEBUG_MSG "Entering " & METHOD
    On Error Resume Next
    
    Dim dummyColl As Collection
    Set dummyColl = New Collection
    ctlSO.SmartOccurrence = dummyColl
    Set dummyColl = Nothing
    
    DEBUG_MSG "Exiting " & METHOD
End Sub

Private Sub ShowSmartOccurrenceView(bShow As Boolean)
    Const METHOD = "ShowSmartOccurrenceView"
    DEBUG_MSG "Entering " & METHOD & ": bShow=" & bShow

    ctlCatalogPalette.Visible = Not bShow
    ctlCatalogPreviewCtl.Visible = Not bShow And (m_bHaveClicked = True)
    grdSOSymbolParamGrid.Visible = Not bShow
    Frame(ccfSmartOccurrence).Visible = bShow
    Frame(ccfPalette).Visible = Not bShow
    If bShow And (m_bParameterGridHasSO = True Or m_bMultiEditMode = True) Then
        ctlSO.Visible = True
    Else
        ctlSO.Visible = False
    End If

    DEBUG_MSG "Exiting " & METHOD
End Sub
Private Function DoNotResizeControl(oControl As Control) As Boolean
   If oControl Is Frame(ccfStepControls) Or _
      oControl Is imvReinforce Or _
      oControl Is lblReinforceByRule Or _
      oControl Is chkvwByRule Or _
      oControl Is chkByRule Or _
      oControl Is cmdFlip Or _
      oControl Is Label(0) Or _
      oControl Is Label(1) Or _
      oControl Is Label(2) Or _
      oControl Is Label(3) Or _
      oControl Is Label(4) Then
      DoNotResizeControl = True
   Else
      DoNotResizeControl = False
   End If
   Exit Function
   
End Function
Private Sub Label_Change(Index As Integer)

   Select Case Index
          Case 0
               Label(Index).Left = (Frame(ccfUVOffsets).Width / 2) - (Label(Index).Width / 2)
          Case 1
               Label(Index).Left = (Frame(ccfVectorLengths).Width / 2) - (Label(Index).Width / 2)
          Case 2
               Label(Index).Left = (CInt(Frame(ccfPalette).Width * 0.75)) - (Label(Index).Width / 2)
          Case 3
               Label(Index).Left = (Frame(ccfStepControls).Width / 2) - (Label(Index).Width / 2)
          Case 4
               Label(Index).Left = (Frame(ccfProperties).Width / 2) - (Label(Index).Width / 2)
   End Select
End Sub
Private Sub AdjustCommandButtons()
    cmdButton(0).Top = Frame(0).Top + Frame(0).Height - cmdButton(0).Height - 400
    cmdButton(1).Top = cmdButton(0).Top
    cmdButton(2).Top = cmdButton(0).Top
    cmdButton(0).ZOrder 0
    cmdButton(1).ZOrder 0
    cmdButton(2).ZOrder 0
    Exit Sub
End Sub
