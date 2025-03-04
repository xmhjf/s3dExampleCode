Attribute VB_Name = "WeldKeys"

Public Enum InputIndex
    igRefPartIndex = 1
    igNonRefPartIndex = 2
    igClearanceIndex = 3
End Enum

Public Type WELDING_SYMBOL_INPUT_INFO
   eType       As IMSParameterContentTypes
   sInputName  As String
   nUomType    As Long
   dUomValue   As Double
   sValue      As String
   bNeedToSave As Boolean
End Type

Public Const WELD_TYPE_TEE As String = "Tee"
Public Const WELD_TYPE_BUTT As String = "Butt"
Public Const WELD_TYPE_LAP As String = "Lap"
Public Const INTERFACE_NAME_IJWELDINGSYMBOL As String = "IJWeldingSymbol"

