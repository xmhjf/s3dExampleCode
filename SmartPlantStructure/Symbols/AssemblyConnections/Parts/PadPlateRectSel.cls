VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PadPlateRectSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'---------------------------------------------------------------------------
'    Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'
'
'   Selector:
'   ---------
'   It selects the list of avaible Items or SmartClasses in the context of
'   the smart occurrence.
'
'
'History
'    AS      April 21, 2006      Creation
'---------------------------------------------------------------------------------------


Option Explicit

Const m_SelectorProgid As String = "SPSPartMacros.PadPlateRectSel"
Const m_SelectorName As String = "SPSPartMacros.PadPlateRectSel"
Const m_PadPlateRectNode As String = "PadPlateRectangular"
Const m_PadPlateRectGenItem As String = "GenericPadPlateRect_1"

Implements IJDUserSymbolServices


Public Sub SelectorInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
    'MsgBox m_SelectorProgid + "SelectorInputs"
  pIH.SetInput "MemberPart"
  pIH.SetInput "Surface"
  
  Exit Sub
ErrorHandler:
  pIH.ReportError
End Sub
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
  On Error GoTo ErrorHandler
  ' TODO - Add your question below
  
  Exit Sub
ErrorHandler:
  pQH.ReportError
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
  On Error GoTo ErrorHandler
  
    Dim oStructPlate As IJStructPlate
    Dim oOcc As IJSmartOccurrence
    Dim oSC As IJSmartClass
    Dim oOccParent As IJSmartOccurrence
    Dim oIJAttribs As IJDAttributes
    Dim oRefColl As IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim colPorts As IJElements
    Dim colPlateSizeParams As CollectionProxy
    Dim pDesignParent As IJDesignParent
    Dim pIJDesignChild As IJDesignChild
    Dim oStructConn As IJAppConnection
          
    'Pad attributes
    Dim oAttributes As IJDAttributes
    Dim colPlatePartDim As CollectionProxy
    Dim lSizingRule As Long
    Dim dClearance As Double
    Dim dPadDepth As Double
    Dim dPadWidth As Double
                                                                
    'member Part Xsection
    Dim dWidth As Double
    Dim dDepth As Double
    Dim dSectionDepth As Double
    Dim dSectionWidth As Double

    'Plate RectangularNode and its items
    Dim oPadPlateRectNode As IJRefDataNode2
    Dim oPadItemNode As IJRefDataNode2
    Dim oPadItems As IJDObjectCollection
    Dim index  As Integer
    Dim sDisplayName As String
    Dim bFoundItem As Boolean
    Dim oCatalogPOM As IJDPOM
    Dim oNamingContext As IJDNamingContextObject
    Dim oItem As Object
    
    'Initialize the plate and its parent
    Set oStructPlate = pSL.SmartOccurrence
    Set pIJDesignChild = pSL.SmartOccurrence
    Set pDesignParent = pIJDesignChild.GetParent
    
    If TypeOf pDesignParent Is IJStructAssemblyConnection Then
        Set oOccParent = pDesignParent
    Else
        Exit Sub
    End If
    
    Set oIJAttribs = oOccParent   '.ItemObject
    Set colPlateSizeParams = oIJAttribs.CollectionOfAttributes("IJUAGenSurfAsmConnPad")
    lSizingRule = colPlateSizeParams.Item("SizingRule").Value
    dClearance = colPlateSizeParams.Item("Offset").Value
    
    Set oStructConn = pDesignParent
    oStructConn.enumPorts colPorts

    'Get the inputs- supported port and the supg surface
    If colPorts.Count = 1 Then
        'Check if a valid member part is selcted
        Set oSuppedPort = colPorts.Item(1)
    ElseIf colPorts.Count = 2 Then ' a plane is also input
        Set oSuppedPort = GetEndPort(colPorts)
    End If
    If oSuppedPort Is Nothing Then
        'We should never reach here as the AC should check the inputs before.
        Exit Sub
        
    End If
   
    Set oSuppedPart = oSuppedPort.Part
    
    'select the correct item based on x-section size if sizing rule is Auto
    ' else juast select genericRectPlate
    If lSizingRule = 1 Then 'ByRule
        'Get cross-section size
        Call GetSectionWidthAndDepth(oSuppedPart, dSectionWidth, dSectionDepth)
        
        'Add clearance
        dWidth = dSectionWidth + 2# * dClearance
        dDepth = dSectionDepth + 2# * dClearance
        
       'Get all plates under the rect pad node
        Set oCatalogPOM = GetCatalogResourceManager()
        Set oNamingContext = New NamingContextObject
        On Error Resume Next
        Set oSC = oNamingContext.ObjectMoniker(oCatalogPOM, m_PadPlateRectNode)
        If oSC Is Nothing Then
            'if here and no smart class, error out
        End If
        On Error GoTo ErrorHandler
    
        Set oPadItems = oSC.Children
        'MsgBox "In RectPlate" + Str(oPadItems.Count)
        
        bFoundItem = False
        
        For Each oPadItemNode In oPadItems
            sDisplayName = oPadItemNode.DisplayText
            'TODO: No need to check this
            If sDisplayName <> m_PadPlateRectGenItem Then
                'Get the size of the pad
                Set oAttributes = oPadItemNode
                Set colPlatePartDim = oAttributes.CollectionOfAttributes("IJUASPSPlatePartDim")
                dPadWidth = colPlatePartDim.Item("Width").Value
                dPadDepth = colPlatePartDim.Item("Length").Value
            
                If dPadWidth > dWidth And dPadDepth > dDepth Then
                    'Select the first item which has dimension bigger than required.
                    pSL.Add sDisplayName
                    bFoundItem = True
                    Exit For
                End If
                
            End If
        Next
        Set oPadItems = Nothing
        If Not bFoundItem Then
            pSL.Add m_PadPlateRectGenItem
        End If
    
    Else
        pSL.Add m_PadPlateRectGenItem
    End If

    'MsgBox "out of RectPadPlate SelectorLogic"

  Exit Sub
ErrorHandler:
  pSL.ReportError
End Sub
  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************

' ** End CM **





' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
  Dim pDFact As New DefinitionFactory
  pDFact.InitAbstractSelector pSelector
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pSelector
  SelectorInputs pIH
  Dim pQH As IJDQuestionsHelper
  Set pQH = New QuestionHelper
  pQH.Selector = pSelector
  SelectorQuestions pQH
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMSelector(pRep As IJDRepresentation)
  Dim pSL As IJDSelectorLogic
  Set pSL = New SelectorLogic
  pSL.Representation = pRep
  SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************


