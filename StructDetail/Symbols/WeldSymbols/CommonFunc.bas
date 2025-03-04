Attribute VB_Name = "CommonFuncs"
 '/*******************************************************************
'Copyright (C) 1998, Intergraph Corporation.  All rights reserved.
'
'
'Project: S:\StructDetail\Middle\Symbols\WeldSymbols\WeldSymbols.vbp
'
'File: S:\StructDetail\Middle\Symbols\WeldSymbols\CommonFuncs.bas
'
'Revision:
'     02/07/01 GDreybus
'
' Description:
' The functions and subs in this file are commonly used by slot
' symbols.
'
'*******************************************************************/

Option Explicit
Private Const MODULE = "S:\StructDetail\Middle\Symbols\WeldSymbols\CommonFunc.bas"

Sub SetWeldCommonInputs(pDef As IJDSymbolDefinition)
  
  Dim pInput As IJDInput
  Set pInput = New DInput

  pInput.Name = "RefPart"
  pInput.Description = "Reference Part"
  pInput.Index = igRefPartIndex
  pDef.IJDInputs.Add pInput
  pInput.Reset
    
  pInput.Name = "NonRefPart"
  pInput.Description = "Non Reference Part"
  pInput.Index = igNonRefPartIndex
  pDef.IJDInputs.Add pInput
  pInput.Reset

End Sub


'********************************************************************
' ' Routine: LogError
'
' Description:  default Error Reporter
'********************************************************************
Public Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
     
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
End Function


Public Sub UpdateControlFlags(oObject As Object, _
                              eFlag As IMSEntitySupport.ControlFlagConstant, _
                              Optional bSetFlag As Boolean = True)
 Const sMETHOD = "UpdateControlFlags"
    
    Dim oFlags As IJControlFlags

 On Error GoTo ErrorHandler
  
    If TypeOf oObject Is IJControlFlags Then
        Set oFlags = oObject
        oFlags.ControlFlags(eFlag) = IIf(bSetFlag, eFlag, 0)
        Set oFlags = Nothing
    End If
    
    Exit Sub
    
ErrorHandler:
   Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

Public Sub GetWeldingSymbolInputsDescription( _
                     ByRef arrayInputs() As WELDING_SYMBOL_INPUT_INFO, _
                     ByVal strWeldType As String, _
                     ByRef nInputCount As Long)
   '
   ' !!! Change array size if adding new attributes
   '
   ReDim arrayInputs(1 To 34)
   Dim nInputIndex As Long
   
   nInputCount = 0
   
   ' 1
   nInputIndex = 1
   arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_SYMBOL
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = 0
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 2
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_SYMBOL
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = 0
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 3
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_GROOVE
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = 0
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 4
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_GROOVE
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = 0
   arrayInputs(nInputIndex).dUomValue = 0

   ' 5
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_GROOVE_SIZE
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 6
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_GROOVE_SIZE
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
   arrayInputs(nInputIndex).dUomValue = 0

   ' 7
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_ACTUAL_THROAT_THICKNESS
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 8
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_ACTUAL_THROAT_THICKNESS
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
   arrayInputs(nInputIndex).dUomValue = 0

   ' 9
   If Not strWeldType = WELD_TYPE_BUTT Then
      'Nominal throat thickness is not used for butt
      nInputIndex = nInputIndex + 1
      arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_NOMINAL_THROAT_THICKNESS
      arrayInputs(nInputIndex).eType = igValue
      arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
      arrayInputs(nInputIndex).dUomValue = 0
   
   ' 10
      nInputIndex = nInputIndex + 1
      arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_NOMINAL_THROAT_THICKNESS
      arrayInputs(nInputIndex).eType = igValue
      arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
      arrayInputs(nInputIndex).dUomValue = 0
      
   End If
   
   ' 11
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = FIELD_WELD
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = 0
   arrayInputs(nInputIndex).dUomValue = 0

   ' 12
   If Not strWeldType = WELD_TYPE_BUTT Then
      ' All around is not used for butt
      nInputIndex = nInputIndex + 1
      arrayInputs(nInputIndex).sInputName = ALL_AROUND
      arrayInputs(nInputIndex).eType = igValue
      arrayInputs(nInputIndex).nUomType = 0
      arrayInputs(nInputIndex).dUomValue = 0
      
   End If
   
   ' 13
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_SUPPLEMENTARY_SYMBOL
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = 0
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 14
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_SUPPLEMENTARY_SYMBOL
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = 0
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 15
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = TAIL_NOTES
   arrayInputs(nInputIndex).eType = igString

   ' 16
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = TAIL_NOTE_IS_REFERENCE
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = 0
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 17
   If Not strWeldType = WELD_TYPE_BUTT Then
      ' Length, Pitch are not used for Butt
      nInputIndex = nInputIndex + 1
      arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_LENGTH
      arrayInputs(nInputIndex).eType = igValue
      arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
      arrayInputs(nInputIndex).dUomValue = 0
   
   ' 18
      nInputIndex = nInputIndex + 1
      arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_LENGTH
      arrayInputs(nInputIndex).eType = igValue
      arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
      arrayInputs(nInputIndex).dUomValue = 0
   
   ' 19
      nInputIndex = nInputIndex + 1
      arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_PITCH
      arrayInputs(nInputIndex).eType = igValue
      arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
      arrayInputs(nInputIndex).dUomValue = 0
   
   ' 20
      nInputIndex = nInputIndex + 1
      arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_PITCH
      arrayInputs(nInputIndex).eType = igValue
      arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
      arrayInputs(nInputIndex).dUomValue = 0
      
   End If
   
'   Diameter is not used for all
   
'   ' 21
'   nInputIndex = nInputIndex + 1
'   arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_DIAMETER
'   arrayInputs(nInputIndex).eType = igValue
'   arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
'   arrayInputs(nInputIndex).dUomValue = 0
'
'   ' 22
'   nInputIndex = nInputIndex + 1
'   arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_DIAMETER
'   arrayInputs(nInputIndex).eType = igValue
'   arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
'   arrayInputs(nInputIndex).dUomValue = 0
   
   
   ' 23
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_CONTOUR
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = 0
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 24
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_CONTOUR
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = 0
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 25
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_FINISH_METHOD
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = 0
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 26
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_FINISH_METHOD
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = 0
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 27
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_ROOT_OPENING
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
   arrayInputs(nInputIndex).dUomValue = 0
      
   ' 28
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_ROOT_OPENING
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 29
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_GROOVE_ANGLE
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = UNIT_ANGLE
   arrayInputs(nInputIndex).dUomValue = 0
   
   ' 30
   nInputIndex = nInputIndex + 1
   arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_GROOVE_ANGLE
   arrayInputs(nInputIndex).eType = igValue
   arrayInputs(nInputIndex).nUomType = UNIT_ANGLE
   arrayInputs(nInputIndex).dUomValue = 0
   
'   Number of welds is not used for all
   ' 31
'   nInputIndex = nInputIndex + 1
'   arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_NUMBER_OF_WELDS
'   arrayInputs(nInputIndex).eType = igValue
'   arrayInputs(nInputIndex).nUomType = 0
'   arrayInputs(nInputIndex).dUomValue = 0
'
'   ' 32
'   nInputIndex = nInputIndex + 1
'   arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_NUMBER_OF_WELDS
'   arrayInputs(nInputIndex).eType = igValue
'   arrayInputs(nInputIndex).nUomType = 0
'   arrayInputs(nInputIndex).dUomValue = 0
      
   ' 33
   If Not strWeldType = WELD_TYPE_BUTT Then
      nInputIndex = nInputIndex + 1
      arrayInputs(nInputIndex).sInputName = PRIMARY_SIDE_ACTUAL_LEG_LENGTH
      arrayInputs(nInputIndex).eType = igValue
      arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
      arrayInputs(nInputIndex).dUomValue = 0
      
   ' 34
      nInputIndex = nInputIndex + 1
      arrayInputs(nInputIndex).sInputName = SECONDARY_SIDE_ACTUAL_LEG_LENGTH
      arrayInputs(nInputIndex).eType = igValue
      arrayInputs(nInputIndex).nUomType = UNIT_DISTANCE
      arrayInputs(nInputIndex).dUomValue = 0
   End If
   
   nInputCount = nInputIndex
End Sub
'
' This method reads attributes from physical connection symobl inputs and
' store them on IJWeldingSymbol interface
'
Public Sub CopyAttributesToIJWeldingSymbol( _
         ByVal sWeldType As String, _
         ByVal oSymbol As IJDSymbol)
   On Error GoTo ErrorHandler
   
   Dim arrayWeldingSymbolInputs() As WELDING_SYMBOL_INPUT_INFO
   Dim nWeldingInputCount As Long
    
   ' Get welding symbol inputs
   GetWeldingSymbolInputsDescription _
                               arrayWeldingSymbolInputs, _
                               sWeldType, _
                               nWeldingInputCount
                             
   ' Get attributes stored on IJWeldingSymbol interface
   GetAttributesOnIJWeldingSymbol _
                               oSymbol, _
                               nWeldingInputCount, _
                               arrayWeldingSymbolInputs
                             
   Dim oSymbolDef As IJDSymbolDefinition
   Dim oSymbolInputs As IJDInputs
   Dim oEnumArg As IEnumJDArgument

   Set oSymbolDef = oSymbol.IJDSymbolDefinition(1)
   Set oSymbolInputs = oSymbolDef.IJDInputs
   Set oSymbolDef = Nothing
   Set oEnumArg = oSymbol.IJDValuesArg.GetValues(igINPUT_ARGUMENTS_MERGE)
   
   Dim oArgument As IJDArgument
   Dim bFound As Long
   Dim oSymbolInput As IJDInput
   Dim nAttributeIndex   As Integer
   Dim oParameterContent As IJDParameterContent

   Do
      oEnumArg.Next 1, oArgument, bFound
      If bFound = 0 Then
          Exit Do
      End If

      Set oSymbolInput = oSymbolInputs.Item(oArgument.Index)
      If igINPUT_IS_A_PARAMETER = oSymbolInput.Properties Then
         Set oParameterContent = oArgument.Entity
         For nAttributeIndex = 1 To nWeldingInputCount
            If oSymbolInput.Name = arrayWeldingSymbolInputs(nAttributeIndex).sInputName Then
               If oParameterContent.Type = igString Then
                  If Not arrayWeldingSymbolInputs(nAttributeIndex).sValue = oParameterContent.String Then
                     arrayWeldingSymbolInputs(nAttributeIndex).bNeedToSave = True
                     arrayWeldingSymbolInputs(nAttributeIndex).sValue = oParameterContent.String
                  Else
                     arrayWeldingSymbolInputs(nAttributeIndex).bNeedToSave = False
                  End If
               Else
                  If Not arrayWeldingSymbolInputs(nAttributeIndex).dUomValue = oParameterContent.UomValue Then
                     arrayWeldingSymbolInputs(nAttributeIndex).bNeedToSave = True
                     arrayWeldingSymbolInputs(nAttributeIndex).dUomValue = oParameterContent.UomValue
                  Else
                     arrayWeldingSymbolInputs(nAttributeIndex).bNeedToSave = False
                  End If
               End If
            End If
         Next
         Set oParameterContent = Nothing
      End If
      Set oSymbolInput = Nothing
   Loop
   Set oSymbolInputs = Nothing
   Set oEnumArg = Nothing
   
   Dim vInterfaceType As Variant
   Dim oAttributes As IJDAttributes
   Dim oAttributesCol As IJDAttributesCol
   Dim oAttribute As IJDAttribute
   Dim oAttributeInfo As IJDAttributeInfo
   
   Set oAttributes = oSymbol
   
   ' Update IJWeldingSymbol attributes
   For Each vInterfaceType In oAttributes
      Set oAttributesCol = oAttributes.CollectionOfAttributes(vInterfaceType)
      If oAttributesCol.Count > 0 Then
         For Each oAttribute In oAttributesCol
            Set oAttributeInfo = oAttribute.AttributeInfo
            If LCase(oAttributeInfo.InterfaceName) = LCase(INTERFACE_NAME_IJWELDINGSYMBOL) Then
               For nAttributeIndex = 1 To nWeldingInputCount
                  If UCase(oAttributeInfo.Name) = UCase(arrayWeldingSymbolInputs(nAttributeIndex).sInputName) Then
                     If arrayWeldingSymbolInputs(nAttributeIndex).bNeedToSave = True Then
                        If igString = oAttributeInfo.Type Then
                           oAttribute.Value = arrayWeldingSymbolInputs(nAttributeIndex).sValue
                        Else
                           oAttribute.Value = arrayWeldingSymbolInputs(nAttributeIndex).dUomValue
                        End If
                     End If
                     Exit For
                  End If
               Next
            End If
            Set oAttributeInfo = Nothing
            Set oAttribute = Nothing
         Next
      End If
      Set oAttributesCol = Nothing
   Next
   Set oAttributes = Nothing
   
   Exit Sub
ErrorHandler:
   Err.Raise Err.Number, MODULE & "WeldSymbols::CopyAttributesToIJWeldingSymbol"
End Sub
'
' This method reads IJWeldingSymbol attributes
'
Public Sub GetAttributesOnIJWeldingSymbol( _
                     ByVal oPCObject As IJSmartOccurrence, _
                     ByVal nAttributeCount As Integer, _
                     ByRef arrayWeldingSymbolInputs() As WELDING_SYMBOL_INPUT_INFO)
                                               
   On Error GoTo ErrorHandler

   Dim oAttributes             As IMSAttributes.IJDAttributes
   Dim vInterfaceType          As Variant
   Dim nAttributeSet           As Integer

   nAttributeSet = 0
   Set oAttributes = oPCObject
   For Each vInterfaceType In oAttributes
      Dim oAttributesCol          As IJDAttributesCol

      Set oAttributesCol = oAttributes.CollectionOfAttributes(vInterfaceType)
      If oAttributesCol.Count > 0 Then
         Dim oAttribute       As IJDAttribute
         Dim oAttributeInfo   As IJDAttributeInfo
         Dim oInterface       As IJDInterfaceInfo

         For Each oAttribute In oAttributesCol
            Dim nAttributeIndex As Integer
            
            Set oAttributeInfo = oAttribute.AttributeInfo
            If LCase(oAttributeInfo.InterfaceName) = LCase(INTERFACE_NAME_IJWELDINGSYMBOL) Then
               For nAttributeIndex = 1 To nAttributeCount
                  If LCase(oAttributeInfo.Name) = LCase(arrayWeldingSymbolInputs(nAttributeIndex).sInputName) Then
                     If igString = oAttributeInfo.Type Then
                        arrayWeldingSymbolInputs(nAttributeIndex).sValue = oAttribute.Value
                     Else
                        arrayWeldingSymbolInputs(nAttributeIndex).dUomValue = oAttribute.Value
                     End If
                     Exit For
                  End If
               Next
            End If
            Set oAttribute = Nothing
            Set oAttributeInfo = Nothing
         Next
      End If
      Set oAttributesCol = Nothing
   Next
   Set oAttributes = Nothing
   
   Exit Sub
   
ErrorHandler:
   Err.Raise Err.Number, MODULE & "WeldSymbols::GetAttributesOnIJWeldingSymbol"
End Sub

