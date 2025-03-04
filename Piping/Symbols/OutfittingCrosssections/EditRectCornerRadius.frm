VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form EditRectCornerRadius 
   Caption         =   "EditRectCornerRadius"
   ClientHeight    =   6540
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   5676
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   5676
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   612
      Left            =   3120
      TabIndex        =   5
      Top             =   5520
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   612
      Left            =   600
      TabIndex        =   4
      Top             =   5520
      Width           =   1332
   End
   Begin VB.Frame Frame2 
      Caption         =   "Representation"
      Height          =   1572
      Left            =   600
      TabIndex        =   2
      Top             =   3480
      Width           =   3732
      Begin VB.ListBox Flatoval 
         Height          =   624
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   2892
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input representations"
      Height          =   2772
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3732
      Begin MSFlexGridLib.MSFlexGrid ParameterGrid 
         Height          =   1932
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   3252
         _ExtentX        =   5736
         _ExtentY        =   3408
         _Version        =   327680
         Rows            =   4
         Cols            =   4
         FixedCols       =   0
         AllowUserResizing=   3
      End
   End
End
Attribute VB_Name = "EditRectCornerRadius"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Win32 API
Private Declare Function SetParent& Lib "user32" (ByVal hWndChild As Long, _
                                                  ByVal hWndNewParent As Long)


Private m_SymbolOcc As IJDSymbol
Private m_SymbolDef As IJDSymbolDefinition
Private m_TransactionMgr As IMSTransactionManager.IJTransactionMgr
Private m_RefsColl As Object
Private m_ValuesColl As Object
Private m_ActiveConnection As WorkingSetLibrary.IJDConnection
Private m_SymbolFactory As DSymbolEntitiesFactory
Private m_bClickOnOK As Boolean

Dim bInputsChange As Boolean
Dim Rep As IJDRepresentation
Dim argument As IJDArgument
Dim PC As IJDParameterContent
Dim index As Long
Dim argIndex As Long
Dim nbArgs As Long
Dim nbInputs As Long
Dim nbReps As Long
Dim inputDesc As IJDInput
Dim bProcess As Boolean
Dim bFirstTime As Boolean
Dim bInitialize As Boolean
Dim IEnumJDArg As IEnumJDArgument
Dim IJDEditJDArg As IJDEditJDArgument
Dim firstRepIndex As Integer
Dim IJDInputs As IJDInputs
Dim found As Long
Dim DInput As DInput


Private Sub Form_Load()
        
    Set m_SymbolFactory = New DSymbolEntitiesFactory
    GetActiveConnection
    m_bClickOnOK = False
    
    ' Sets the edit form as a chlid of the mainframe
    Dim mf As Form
    Dim oTrader As New Trader
    Set mf = oTrader.Service("MainFrame", "")
    SetParent Me.hWnd, mf.hWnd
    Set oTrader = Nothing
    Set mf = Nothing
        
    'Initialization boolean
    bInitialize = True
    bInputsChange = False
    bProcess = False
    
    ' Initialization of the parameter & graphic input grids
    ParameterGrid.ColWidth(0) = 500
    ParameterGrid.ColWidth(1) = 1500
    ParameterGrid.ColWidth(2) = 600
    'ParameterGrid.ColWidth(3) = 800
    ParameterGrid.ColAlignment(0) = 7
    ParameterGrid.ColAlignment(1) = 7
    ParameterGrid.ColAlignment(2) = 7
    'ParameterGrid.ColAlignment(3) = 7
    ParameterGrid.Col = 0
    ParameterGrid.Row = 0
    ParameterGrid.Text = "Index"
    ParameterGrid.Col = 1
    ParameterGrid.Text = "Input"
    ParameterGrid.Col = 2
    ParameterGrid.Text = "Value"
    'ParameterGrid.Col = 3
    'ParameterGrid.Text = "Value"
        
    ' Get the Symbol Definition
    ' Rebyvalmove after cycle 1
    Set m_SymbolDef = m_SymbolOcc.IJDSymbolDefinition(2)
    Set IJDInputs = m_SymbolDef
    nbReps = m_SymbolDef.IJDRepresentations.RepresentationCount
    
    ' **** Init the parameter&graphic input grids ****
    ParameterGrid.Rows = 1
    nbInputs = IJDInputs.InputCount
    ' Loop on the inputs
    For index = 1 To nbInputs
        Set DInput = IJDInputs.GetInputAtIndex(index)
        
        If DInput.Properties = igINPUT_IS_A_PARAMETER Then
            ' it's a parameter
            ParameterGrid.Rows = ParameterGrid.Rows + 1
            ParameterGrid.Row = ParameterGrid.Rows - 1
            
            ' Set the index
            ParameterGrid.Col = 0
            ParameterGrid.Text = index
            'Set the name
            ParameterGrid.Col = 1
            ParameterGrid.Text = DInput.Name
            ' Set BYREF flag
            'ParameterGrid.Col = 2
            'ParameterGrid.Text = "N"
        End If
        Set DInput = Nothing
    Next index
    Set IJDInputs = Nothing
    
    ' **** Loads the parameters passed by value ****

    Set IEnumJDArg = m_SymbolOcc.IJDValuesArg.GetValues
    
    IEnumJDArg.Reset
    Do
        IEnumJDArg.Next 1, argument, found
        If found = 0 Then
            Set argument = Nothing
            Exit Do
        End If
        
        Set PC = argument.entity
        ParameterGrid.Row = argument.index
        ' Set the input value
        If PC.Type = igValue Then
            ParameterGrid.Col = 2
            ParameterGrid.Text = Str(PC.UomValue)
        End If
        Set PC = Nothing
    Loop
    
    Set IEnumJDArg = Nothing
    
    ' **** Loads the parameters grid passed by reference ****

    Set m_RefsColl = m_SymbolOcc.IJDReferencesArg.GetReferences
    
    If Not m_RefsColl Is Nothing Then
        Set IEnumJDArg = m_RefsColl
        IEnumJDArg.Reset
        Do
            IEnumJDArg.Next 1, argument, found
            If found = 0 Then Exit Do
            
            Set PC = argument.entity
            argIndex = argument.index
            ParameterGrid.Row = argIndex
            Set inputDesc = m_SymbolDef.IJDInputs.GetInputAtIndex(argIndex)
            
            ' Set the BYREF flag
            'ParameterGrid.Col = 2
            'ParameterGrid.Text = "Y"
            
            ' Set the input value
            If PC.Type = igValue Then
                ParameterGrid.Col = 2
                ParameterGrid.Text = Str(PC.UomValue)
            End If
            Set PC = Nothing
        Loop
        
        Set IEnumJDArg = Nothing
    End If
    
    ' *** Feed the representation ComBox ****
        
'    For index = 1 To nbReps
'        Set Rep = m_SymbolDef.IJDRepresentations.GetRepresentationAtIndex(index)
'        RepresentationList.AddItem Rep.Name
'        Set Rep = Nothing
'    Next index
'    ' Select the current representation name
'    Set Rep = m_SymbolDef.IJDRepresentations.GetRepresentationById(m_SymbolOcc.IJDRepresentationArg.RepresentationId)
'    RepresentationList.Selected(m_SymbolOcc.IJDRepresentationArg.RepresentationId - 1) = True
'    Set Rep = Nothing
'    firstRepIndex = m_SymbolOcc.IJDRepresentationArg.RepresentationId - 1
'    bInitialize = False
End Sub


Private Sub cmdCancel_Click()
    m_TransactionMgr.Abort
    Set m_SymbolOcc = Nothing
    Set m_SymbolDef = Nothing
       
    Unload Me
End Sub


Private Function CheckParams() As Boolean
    Dim index As Long
    
    CheckParams = True
    ParameterGrid.Col = 2
    For index = 1 To (ParameterGrid.Rows - 1)
        ParameterGrid.Row = index
        If StrComp(ParameterGrid.Text, "") = 0 Or Not IsNumeric(ParameterGrid.Text) Then
            CheckParams = False
            Exit Function
        End If
    Next index
    ' Scale should be non null
    ParameterGrid.Row = 1
    If (ParameterGrid.Text <= 0) Then
        MsgBox "Width must be greater than 0", vbOKOnly, "Error"
        CheckParams = False
    End If
End Function

Sub GetActiveConnection()
    Dim oTrader As New Trader
    Dim oWorkingSet As WorkingSet
    
    Set oWorkingSet = oTrader.Service("WorkingSet", "")
    Set m_ActiveConnection = oWorkingSet.ActiveConnection
    
    Set oTrader = Nothing
    Set oWorkingSet = Nothing
End Sub

Sub CreateAndConnectRefColl()

    Set m_RefsColl = Nothing
    Set m_RefsColl = m_SymbolFactory.CreateEntity(ReferencesCollection, m_ActiveConnection.ResourceManager)
    Set IEnumJDArg = m_RefsColl
    m_SymbolOcc.IJDReferencesArg.SetReferences IEnumJDArg
    Set IEnumJDArg = Nothing
    
End Sub



Private Sub cmdOK_Click()
    
    Dim IJDEditJDArgForValues As IJDEditJDArgument
    Dim IJDEditJDArgForRefs As IJDEditJDArgument
    Dim ArgumentEnumForValues As New DEnumArgument
    Dim curIndex As Long
    Dim entity As Object
    
    ' Check whether the input grid values are valid
    If Not CheckParams Then
        Exit Sub
    End If
    
    m_bClickOnOK = True
    
'    ' Update the repId
'    If RepresentationList.ListIndex <> firstRepIndex Then
'        Set Rep = m_SymbolDef.IJDRepresentations.GetRepresentationByName(RepresentationList.Text)
'        m_SymbolOcc.IJDRepresentationArg.RepresentationId = Rep.RepresentationId
'        Set Rep = Nothing
'    End If
                
    Set argument = New DArgument
    
    ' Update the repId
'    Set Rep = m_SymbolDef.IJDRepresentations.GetRepresentationByName(RepresentationList.Text)
'    m_SymbolOcc.IJDRepresentationArg.RepresentationId = Rep.RepresentationId
'    Set Rep = Nothing
        
    Set IJDEditJDArgForValues = ArgumentEnumForValues
    
    'Update the parameters
'    If (ParameterGrid.Rows > 1) Then
    
        ' Set the Parameters
'        For index = 1 To (ParameterGrid.Rows - 1)
'            ParameterGrid.Row = index

            'ParameterGrid.Col = 2
            
            'If StrComp(ParameterGrid.Text, "Y") = 0 Then
                '** Passed by reference **
                'If m_RefsColl Is Nothing Then CreateAndConnectRefColl
                'Set IJDEditJDArgForRefs = Nothing
               ' Set IJDEditJDArgForRefs = m_RefsColl
                
                
'                ParameterGrid.Col = 0
'                curIndex = ParameterGrid.Text
'
'                Set entity = IJDEditJDArgForRefs.GetEntityAtIndex(curIndex)
'
'                If entity Is Nothing Then
                    ' No existing persistent parameter passed by ref
'                    ParameterGrid.Col = 0
'                    argument.index = ParameterGrid.Text
                    
                    'Instanciate a persistent and filled PC
'                    ParameterGrid.Col = 2
'                    Set PC = m_SymbolFactory.CreateEntity(ParameterContent, m_ActiveConnection.ResourceManager)
'                    PC.Type = igValue
'                    PC.UomValue = Val(ParameterGrid.Text)
'                    argument.entity = PC
'                    IJDEditJDArgForRefs.SetArg argument
'                    Set PC = Nothing
'                Else
                    ' Already defined persistent PC passed by ref
'                    Set PC = entity
'                    ParameterGrid.Col = 2
'                    If PC.Type = igValue Then PC.UomValue = ParameterGrid.Text
'                End If
'                Set entity = Nothing
'
'            Else
                ' ** Passed By Value **
                ' Feed the PC
'                Set PC = New DParameterContent
'                ParameterGrid.Col = 2
'                PC.UomValue = Val(ParameterGrid.Text)
'                PC.Type = igValue
'
                ' Feed the Argument
'                ParameterGrid.Col = 0
'                argument.index = ParameterGrid.Text
'                argument.entity = PC
                ' Add the argument to the arg collection
                IJDEditJDArgForValues.SetArg argument
'                Set PC = Nothing
'            End If
'
'        Next index
'        m_SymbolOcc.IJDValuesArg.SetValues ArgumentEnumForValues.IEnumJDArgument
'    End If
    
    ' Update the graphical outputs

'    firstRepIndex = RepresentationList.ListIndex
    'm_SymbolOcc.IJDInputsArg.Update
    m_TransactionMgr.Compute
    
'    GraphicSelectHelper1.RaiseEventClickOnOK
    
    bInputsChange = False
    Set IJDEditJDArgForRefs = Nothing
    Set ArgumentEnumForValues = Nothing
    Set IJDEditJDArgForValues = Nothing
    Set m_SymbolOcc = Nothing
    Set m_SymbolDef = Nothing
    Set argument = Nothing
    
    Unload Me

End Sub



Private Sub Form_Unload(Cancel As Integer)

If Not m_bClickOnOK Then
'    GraphicSelectHelper1.RaiseEventClickOnCancel
End If

    Set m_ActiveConnection = Nothing
    Set m_SymbolOcc = Nothing
    Set m_SymbolDef = Nothing
    Set m_SymbolFactory = Nothing
End Sub

'Private Sub GraphicSelectHelper1_ClickOnOK()

'End Sub

Private Sub ParameterGrid_EnterCell()
    bProcess = True
    bFirstTime = True
End Sub

Private Sub ParameterGrid_KeyPress(KeyAscii As Integer)
    
If ParameterGrid.Col = 2 Then

    ' Values column
    bInputsChange = True

    If bFirstTime Then
        ParameterGrid.Text = ""
    End If
    If KeyAscii = vbKeyBack Then
        If StrComp(ParameterGrid.Text, "") <> 0 Then
            ParameterGrid.Text = Left(ParameterGrid.Text, Len(ParameterGrid.Text) - 1)
        End If
    ElseIf KeyAscii = vbKeyReturn Then
        bProcess = False
        bFirstTime = True
    Else
        ParameterGrid.Text = ParameterGrid.Text + Chr(KeyAscii)
    End If
    bFirstTime = False
    
'ElseIf ParameterGrid.Col = 1 Then

    ' BYREF flag column
'    Dim letter As String
'    letter = UCase(Chr(KeyAscii))

'    If letter = "N" And ParameterGrid.Text = "Y" Then
        ' When a BYREF flag changes from Y to N, remove the corresponding reference
        ' from the reference collection
        ' Get the index of the current row
'          End If
'   If Not m_RefsColl Is Nothing Then
'            ParameterGrid.Col = 0
'            argIndex = ParameterGrid.Text
'            Set IJDEditJDArg = m_RefsColl
'            IJDEditJDArg.RemoveAtIndex argIndex
'            Set IJDEditJDArg = Nothing
'       End If
    
    ' Update the cell contain
'    ParameterGrid.Col = 2
'    If letter = "Y" Or letter = "N" Then ParameterGrid.Text = letter
    
End If

End Sub

Private Sub ParameterGrid_LeaveCell()
    bProcess = False
End Sub

Public Property Set SymbolOccurrence(ByRef SymbolOcc As Object)
    Set m_SymbolOcc = SymbolOcc
End Property

Public Property Set TransactionMgr(ByRef TransactionMgr As Object)
    Set m_TransactionMgr = TransactionMgr
End Property







