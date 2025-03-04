Attribute VB_Name = "CommonFuncs"
'/*******************************************************************
'Copyright (C) 2005, Intergraph Corporation.  All rights reserved.
'
'
'Project: S:\StructDetail\Middle\Symbols\EdgeTreatmentSymbols\FreeEdgeTreatmentSym.vbp
'
'File: S:\StructDetail\Middle\Symbols\EdgeTreatmentSymbols\CommonFuncs.bas
'
'Revision:
'     08/24/05 Venu Kasarla.
'
'
' Description:
' The functions and subs in this file are commonly used by Edge Treatment
' symbols.
'
'*******************************************************************/

Option Explicit
Private Const MODULE = "S:\StructDetail\Middle\Symbols\EdgeTreatmentSymbols\CommonFunc.bas"


Sub SetCommonInputs(pDef As IJDSymbolDefinition)
  
  Dim pInput As IJDInput
  Set pInput = New DInput

  pInput.Name = "RefPart"
  pInput.Description = "Reference Part"
  pInput.Index = igRefPartIndex
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

