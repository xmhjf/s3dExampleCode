Attribute VB_Name = "CustomGUIInclude"
'*** Public Enumerations ***

Public Enum ButtonClicked
    eSFormOK = 0
    esApply = 1
    esCancel = 2
    esRBNotify = 3
End Enum

Public Enum cboValues
    ParamInput = 0
End Enum

Public Enum cboViewValues
       Thickness = 0
       Grade = 1
       plateType = 2
       Support1Offset = 3
       Support2Offset = 4
End Enum

Public Enum cmdBtnValues
       LatestLists = 0
       General = 1
       Chamfered = 2
       Flange = 3
       Rib = 4
       DH = 5
       OKBtn = 6
       CancelBtn = 7
       ApplyBtn = 8
End Enum

Public Enum txtBoxViewValues
       Support1Offset = 0
       Support2Offset = 1
End Enum

Public Enum imgValues
       ThicknessDir = 0
End Enum

Public Enum chkBoxValues
       OffsetSelected = 0
End Enum

Public Enum cboImageValues
       ThicknessDirection = 0
End Enum

Public Enum txtBoxValue
       CurrentUnits = 0
       SymbolName = 1
       FlangeName = 2
End Enum

Public Enum eCatalogNode
       eBracketNode = 1
       eFlangeNode = 2
End Enum

Public Enum ecChkBoxValues
       SymSelected = 0
       FlgSelected = 1
End Enum

Public Enum eCommandMode
        CreateMode = 0
        ModifyMode = 1
End Enum

Public Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Public Const TKCustomBracketGUI = "CustomBracketGUI"
Public Const SESSIONMGR_CustomBracketParameters As String = "CustomBracketGUI.SmartItemParameters"
Public Const VALUEMGR_SOParameterGridSmartItemParameters As String = "SOParameterGrid.SmartItemParameters"

Private m_FormWid As Single
Private m_FormHgt As Single
Private m_ControlPositions() As ControlPositionType

'****************************************************************************************************
' Routine: HasValueChanged
'
' Abstract: Looks at previous values in a combo box to see whether the value is already there
'
' Description:
'   Parameters:   strValue   Potential value to be added
'                 oCombo     Combo box to be checked
'******************************************************************************************************

Public Function HasValueChanged(strValue As String, _
                                 oCombo As ComboBox) _
                                 As Boolean

    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim strCompValue As String
    
    HasValueChanged = True
    
    For i = 0 To oCombo.ListCount - 1
        If strValue = oCombo.List(i) Then
            HasValueChanged = False
            Exit For
        End If
    Next i
    
    Exit Function
ErrorHandler:
    
End Function
Public Sub PopulateSymColl(ByVal oSymCollection As Collection, _
                           ByVal oCatNodeCollection As Collection, _
                           ByRef oOutSymCollection As Collection, _
                           ByRef oOutCatNodeCollection As Collection)
   
   On Error GoTo ErrorHandler
   
   Set oOutSymCollection = Nothing
   Set oOutSymCollection = New Collection
   Set oOutCatNodeCollection = Nothing
   Set oOutCatNodeCollection = New Collection
   If oSymCollection Is Nothing Then Exit Sub
   Dim i As Integer
   For i = 1 To oSymCollection.count
       oOutSymCollection.Add oSymCollection.Item(i)
   Next

   If oCatNodeCollection Is Nothing Then Exit Sub

   For i = 1 To oCatNodeCollection.count
       oOutCatNodeCollection.Add oCatNodeCollection.Item(i)
   Next
   Exit Sub
ErrorHandler:
    MsgBox "Unable to Populate Symbol Collection"
End Sub
'****************************************************************************************************
' Routine: StorePreferences
'
' Abstract: Save user preferences so that they may be instituted if command is re-started
'
' Description:
'   Parameters:   None
'****************************************************************************************************
Public Sub StorePreferences(sFormName As String, lTop As Long, lLeft As Long, Optional lHeight As Long = 0, Optional lWidth As Long = 0)
    
    On Error GoTo ErrorHandler
    
    Dim oPreferences As IJPreferences
    Dim oTrader As Trader
    Set oTrader = New Trader
    Set oPreferences = oTrader.Service("Preferences", "")
    oPreferences.SetLongValue sFormName & "_Top", lTop
    oPreferences.SetLongValue sFormName & "_Left", lLeft
    If lHeight <> 0 Then _
        oPreferences.SetLongValue sFormName & "_Height", lHeight
    If lWidth <> 0 Then _
        oPreferences.SetLongValue sFormName & "_Width", lWidth
        
    Set oPreferences = Nothing
    Set oTrader = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "Unable to Store Preferences"
End Sub
'****************************************************************************************************
' Routine: GetPreferences
'
' Abstract: Read user's preferences and respond
'
' Description:
'   Parameters:   None
'****************************************************************************************************
Public Sub GetPreferences(sFormName As String, lTop As Long, lLeft As Long)

   On Error GoTo ErrorHandler
   
   Dim oPreferences As IJPreferences
   Dim i As Long
   Dim j As Long
   lTop = 0
   lLeft = 0
   Dim oTrader As Trader
   Set oTrader = New Trader
   Set oPreferences = oTrader.Service("Preferences", "")
   i = oPreferences.GetLongValue(sFormName & "_Top", 0)
   j = oPreferences.GetLongValue(sFormName & "_Left", 0)
   If i <> 0 And j <> 0 Then
      lTop = i
      lLeft = j
   End If
   Set oPreferences = Nothing
   Set oTrader = Nothing
   
   Exit Sub
ErrorHandler:
    MsgBox "Unable to Get Preferences"
End Sub

Public Sub GetPreferencesWithHeightWidth(sFormName As String, lTop As Long, lLeft As Long, lHeight As Long, lWidth As Long)

   On Error GoTo ErrorHandler

    Dim oPreferences As IJPreferences
    Dim oTrader As Trader
    Dim valueTop As Long
    Dim valueLeft As Long
    Dim valueHeight As Long
    Dim valueWidth As Long
    
    lTop = 0
    lLeft = 0
    
    Set oTrader = New Trader
    Set oPreferences = oTrader.Service("Preferences", "")
    valueTop = oPreferences.GetLongValue(sFormName & "_Top", 0)
    valueLeft = oPreferences.GetLongValue(sFormName & "_Left", 0)
    If valueTop <> 0 And valueLeft <> 0 Then
       lTop = valueTop
       lLeft = valueLeft
    End If
    
    lHeight = oPreferences.GetLongValue(sFormName & "_Height", 0)
    lWidth = oPreferences.GetLongValue(sFormName & "_Width", 0)
    
    Set oPreferences = Nothing
    Set oTrader = Nothing
    
    Exit Sub
ErrorHandler:
    MsgBox "Unable to Get Preferences"
End Sub
Public Sub StoreObjectBeforeCommit(ByRef m_oPointerStore As IJElements, ByVal oStructObj As Object, ByVal strKeyName As String)
  On Error GoTo ErrorHandler
  
  If m_oPointerStore Is Nothing Then Set m_oPointerStore = New IMSElements.Elements
  If Not GetStoredObject(m_oPointerStore, strKeyName) Is Nothing Then m_oPointerStore.Remove strKeyName
  m_oPointerStore.Add oStructObj, strKeyName
  Exit Sub
ErrorHandler:
    ReportAndRaiseUnanticipatedError "CustomTBGUI", "Unable to StoreObjectBeforeCommit"
End Sub
Public Function GetStoredObject(ByRef m_oPointerStore As IJElements, ByVal strKeyName As String) As Object
  On Error GoTo ErrorHandler
  
  If m_oPointerStore Is Nothing Then ' nothing stored
    Set GetStoredObject = Nothing
  Else ' attempt to retrieve stored object
    On Error Resume Next
    Set GetStoredObject = m_oPointerStore.Item(strKeyName)
  End If
  Exit Function
ErrorHandler:
    ReportAndRaiseUnanticipatedError "CustomTBGUI", "Unable to Get StoreObjectBeforeCommit"
End Function
Public Sub RemoveStoredObject(ByRef m_oPointerStore As IJElements, ByVal strKeyName As String)
   On Error GoTo ErrorHandler
   If Not GetStoredObject(m_oPointerStore, strKeyName) Is Nothing Then m_oPointerStore.Remove (strKeyName)
   Exit Sub
ErrorHandler:
    ReportAndRaiseUnanticipatedError "CustomTBGUI", "RemoveStoredObject"
End Sub

Public Function FoundSmartOccurrencesFromPropertyHolder(ByRef oSmartOcurrenceCollection As Collection) As Boolean
    Dim oHolder As IJSmartItemPropertyHolder
    Dim oTrader As Trader
    Dim bFoundSOs As Boolean
    bFoundSOs = False
    
    On Error Resume Next
    Set oTrader = New Trader
    Set oHolder = oTrader.Service(PLATE_ContourSmartItemHolder, "")
    Set oTrader = Nothing
    
    If Not oHolder Is Nothing Then
        oHolder.GetAllSmartOccurrences oSmartOcurrenceCollection
        Set oHolder = Nothing
        
        If Not oSmartOcurrenceCollection Is Nothing Then
        
            If oSmartOcurrenceCollection.count > 0 Then
                bFoundSOs = True
            End If
        End If
    End If

    FoundSmartOccurrencesFromPropertyHolder = bFoundSOs
End Function

Public Function GetCustomBracketParameterPreferences() As String

    On Error GoTo ErrorHandler

    Dim oPreferences As IJPreferences
    Dim oTrader As Trader
    Dim oValueMgr As IJValueMgr
    Dim sValue As String
    
    Set oTrader = New Trader
    Set oPreferences = oTrader.Service("Preferences", "")
    sValue = oPreferences.GetStringValue(SESSIONMGR_CustomBracketParameters, "")
    Set oValueMgr = oTrader.Service("ValueMgr", "")
    If Not oValueMgr Is Nothing And sValue <> "" Then
        If oValueMgr.ValidKey(VALUEMGR_SOParameterGridSmartItemParameters) = True Then
            oValueMgr.Remove VALUEMGR_SOParameterGridSmartItemParameters
        End If
        
        oValueMgr.Add VALUEMGR_SOParameterGridSmartItemParameters, sValue
        Set oValueMgr = Nothing
    End If
    
    Set oPreferences = Nothing
    Set oTrader = Nothing
    
    GetCustomBracketParameterPreferences = sValue
    Exit Function
ErrorHandler:
    MsgBox "Unable to Get Preferences"
End Function

Public Sub SaveCustomBracketParameterPreferences()
    Dim sValue As String
    Dim oPreferences As IJPreferences
    Dim oTrader As Trader
    Dim oValueMgr As IJValueMgr
    
    sValue = GetSelectionFromValueManager(VALUEMGR_SOParameterGridSmartItemParameters)
    If sValue <> "" Then
        Set oTrader = New Trader
        Set oPreferences = oTrader.Service("Preferences", "")
        oPreferences.SetStringValue SESSIONMGR_CustomBracketParameters, sValue
        
        Set oValueMgr = oTrader.Service("ValueMgr", "")
        If Not oValueMgr Is Nothing Then oValueMgr.Remove VALUEMGR_SOParameterGridSmartItemParameters
        
        Set oPreferences = Nothing
        Set oTrader = Nothing
    End If
End Sub

Public Function GetSelectionFromValueManager(In_sKeyName As String) As String
    Const METHOD = "GetSelectionFromValueManager"
    On Error GoTo ErrorHandler
    
    Dim oValueMgr As IMSICDPInterfacesLib.IJValueMgr
    Dim oTrader As New Trader
    Set oValueMgr = oTrader.Service("ValueMgr", "")
    
    GetSelectionFromValueManager = ""
    If Not oValueMgr Is Nothing Then
        Dim i As Integer
        Dim sKeyName As String
        Dim bFlag As Boolean
        bFlag = oValueMgr.ValidKey(In_sKeyName)

        If bFlag = True Then
            For i = 1 To oValueMgr.count
                sKeyName = oValueMgr.KeyName(i)

                If UCase(sKeyName) = UCase(In_sKeyName) Then
                    GetSelectionFromValueManager = oValueMgr.Item(In_sKeyName)
                    Exit For
                End If
            Next
        End If
    End If
    
CleanUp:
    Set oTrader = Nothing
    Set oValueMgr = Nothing
    Exit Function
ErrorHandler:
    ReportUnanticipatedError "CustomBracketGUI", METHOD
    GoTo CleanUp
End Function
