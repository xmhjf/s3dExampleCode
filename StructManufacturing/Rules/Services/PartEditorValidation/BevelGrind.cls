VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BevelGrind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJMfgEditCustomValidation

Private Declare Function GetTempPath Lib "kernel32" _
         Alias "GetTempPathA" (ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

Private m_oSettingsXML As DOMDocument
Private m_oRadDoc As RAD2D.Document

Private Const MAXIMUM_THICKNESS_TOLERANCE = 0.001 '1mm in meters
Private Const VALIDATE_CUSTOM_BEVEL_GRIND_ATTRIB = "ValidateBevelGrindAttribCustom"

Private Const MODULE As String = "CustomValidation.BevelGrind"

'********************************************************************
' Routine: IJMfgEditCustomValidation_RunCustomValidate
' Abstract: Implements the required interface for the rule.
' NOTE: the oPart is the currently active manufactured part. If the user loaded from an XML file then this object
'     : will be nothing. Your validation rule should attempt to proceed to validate the geometry and/or attributes
'     : on the currently active RAD Document and Active Sheet, or other sheets (in the case of profiles as plates),
'     : as best as it can without the manufacturing part. ONLY FAIL if there is no other option!
'     : Actually your validation rule can and should be able to finish no matter what, you can report fatal errors
'     : in the error collection and hault the persistence process.
'     : The user cannot persist a part that was loaded from XML anyways!
'     : So you don't need to worry about giving an error in that case either! Just do your best to survive!
'     : If it's not a fatal error, a warning might be more appropriate.
'********************************************************************
Private Function IJMfgEditCustomValidation_RunCustomValidate(ByVal oRadDoc As Object, ByVal sXMLSettingsDoc As String, _
ByVal sXMLSettingsPath As String, Optional ByVal oPart As Object) As Object
    On Error GoTo ErrorHandler
    Const METHOD As String = "IJMfgEditCustomValidation_RunCustomValidate"
    Dim oErrorCollection    As New Collection
    Dim oRad2dDoc           As RAD2D.Document
    If oRadDoc Is Nothing Or sXMLSettingsDoc = "" Then GoTo CleanUp
    Set m_oRadDoc = oRadDoc
    Set m_oSettingsXML = New DOMDocument
    m_oSettingsXML.resolveExternals = True
    m_oSettingsXML.validateOnParse = False
    If m_oSettingsXML.loadXML(sXMLSettingsDoc) = True Then
        Set m_oSMS_ViewerDOM = m_oSettingsXML
        ValidateBevelGrindAttrib oErrorCollection
    End If
    Set IJMfgEditCustomValidation_RunCustomValidate = oErrorCollection
CleanUp:
    Set m_oSettingsXML = Nothing
    Set m_oSMS_ViewerDOM = Nothing
    Set oErrorCollection = Nothing
    Set oRad2dDoc = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
' Routine: ValidateBevelGrindAttrib
' Abstract: Validates the Bevel & Grind attributes
'********************************************************************
Private Sub ValidateBevelGrindAttrib(oErrorCollection As Collection)
    On Error GoTo ErrorHandler
    Const METHOD As String = "ValidateBevelGrindAttrib"
    Dim oObject             As Object
    Dim oSubObject          As Object
    Dim oAttribute          As RAD2D.Attribute
    Dim oAttributeSet       As RAD2D.AttributeSet
    Dim bFoundCommonSeam    As Boolean
    
    If oErrorCollection Is Nothing Then Set oErrorCollection = New Collection
    
    On Error Resume Next
    For Each oObject In m_oRadDoc.ActiveSheet.DrawingObjects
        bFoundCommonSeam = False
        If Not oObject Is Nothing Then
            ValidateBevelGrindThicknessMgr oObject, oErrorCollection
            If TypeOf oObject Is RAD2D.Group Then
                If oObject.HasNested Then
                    For Each oSubObject In oObject
                        bFoundCommonSeam = False
                        If Not oSubObject Is Nothing Then
                            ValidateBevelGrindThicknessMgr oSubObject, oErrorCollection
                        End If
                    Next oSubObject
                End If
            End If
        End If
    Next oObject
    On Error GoTo ErrorHandler
CleanUp:
    Set oAttributeSet = Nothing
    Set oSubObject = Nothing
    Set oObject = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Sub

'********************************************************************
' Routine: ValidateBevelGrindThicknessMgr
' Abstract: Acts as a wrapper for calling ValidateBevelGrindThickness & ValidateBevelGrindCommonCutThickness
'********************************************************************
Private Sub ValidateBevelGrindThicknessMgr(oObject As Object, oErrorCollection As Collection)
    On Error GoTo ErrorHandler
    Const METHOD As String = "ValidateBevelGrindThicknessMgr"
    Dim oAttribute      As RAD2D.Attribute
    Dim oAttributeSet   As RAD2D.AttributeSet
    If oObject Is Nothing Then GoTo CleanUp
    For Each oAttributeSet In oObject.AttributeSets
        If Not oAttributeSet Is Nothing Then
            If oAttributeSet.SetName = "Contour" Then
                'Check the bevel or grind thickness
                ValidateBevelGrindThickness oObject, oErrorCollection, oAttributeSet.SetName
            ElseIf oAttributeSet.SetName = "Marking" Then
                'Check the common cutting bevel or grind thickness
                If ValidateBevelGrindCommonCutThickness(oObject, oAttributeSet, oErrorCollection, _
                oAttributeSet.SetName) = True Then
                    Exit For
                End If
            End If
        End If
    Next oAttributeSet
CleanUp:
    Set oAttribute = Nothing
    Set oAttributeSet = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Sub

'********************************************************************
' Routine: ValidateBevelGrindCommonCutThickness
' Abstract: Acts as a wrapper for calling ValidateBevelGrindThickness, but only does so if the
'           object being validated is a strmfg_common_seam_mark!
'********************************************************************
Private Function ValidateBevelGrindCommonCutThickness(oObject As Object, oAttributeSet As RAD2D.AttributeSet, _
oErrorCollection As Collection, Optional sPurpose As String = "Marking") As Boolean
    On Error GoTo ErrorHandler
    Const METHOD As String = "ValidateBevelGrindCommonCutThickness"
    Dim oAttribute          As RAD2D.Attribute
    Dim bFoundCommonSeam    As Boolean
    ValidateBevelGrindCommonCutThickness = False
    bFoundCommonSeam = False
    If oObject Is Nothing Or oAttributeSet Is Nothing Then GoTo CleanUp
    'Check the bevel or grind thickness for common cutting!
    'We need to first make sure it's a common cutting marking object.
    'Otherwise there is no point in running a bevel/grind validation against it!
    'Marking type must be set to "strmfg_common_seam", and that's the ONLY WAY
    'that we can have bevel/grind information on a marking line!
    For Each oAttribute In oAttributeSet
        If Not oAttribute Is Nothing Then
            If oAttribute.Name = "SMS_MARKING" & "||" & "TYPE" Then
                If oAttribute.Value = "strmfg_common_seam_mark" Then
                    bFoundCommonSeam = True
                    Exit For
                Else 'Get out of the loop either way, we found the attribute we were looking for
                    'It wasn't set to the value we needed to run bevel/grind validation against it!
                    Exit For
                End If
            End If
        End If
    Next oAttribute
    If bFoundCommonSeam = True Then
        'Check the Bevel or grind thickness
        ValidateBevelGrindThickness oObject, oErrorCollection, sPurpose
        ValidateBevelGrindCommonCutThickness = True
    End If
CleanUp:
    Set oAttribute = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
' Routine: ValidateBevelGrindThickness
' Abstract: The total thickness for all sections of a bevel or grind
'           to determine if the total bevel or grind thickness exceeds
'           the total thickness for the part.
'********************************************************************
Private Sub ValidateBevelGrindThickness(oObject As Object, oErrorCollection As Collection, _
Optional sPurpose As String = "Contour")
    On Error GoTo ErrorHandler
    Const METHOD As String = "ValidateBevelGrindThickness"
    
    If oObject Is Nothing Then Exit Sub
    If oErrorCollection Is Nothing Then Set oErrorCollection = New Collection
    
    Dim oAttribute  As RAD2D.Attribute
    Dim oError      As STRMFGPartEditorError.CPartEditErrorObj
    Dim eErrLvl     As ErrorLevel
    Dim dPartTotal  As Double
    Dim dCutTotal   As Double
    Dim sError      As String
    Dim sSolution   As String
    Dim sTemp       As String
    Dim sValTemp    As String
    Dim sValue      As String
    Dim bHasCut     As Boolean
    
    bHasCut = False
    sValue = RadGetAttributeValue(oObject, "TYPE", False, m_oSettingsXML, , , , "SMS_BEVEL")
    If sValue <> "" And sValue <> "(null)" Then bHasCut = True
    sValue = ""
    sValue = RadGetAttributeValue(oObject, "TYPE", False, m_oSettingsXML, , , , "SMS_GRIND")
    If sValue <> "" And sValue <> "(null)" And UCase(sValue) <> "GRIND_RADIUS" Then bHasCut = True
    
    If Not bHasCut Then GoTo CleanUp
    
    dCutTotal = 0
    dPartTotal = 0
    sTemp = "THICKNESS"
    Set oAttribute = GetAttribute(m_oRadDoc.ActiveSheet, "SMS_PLATE" & "||" & "THICKNESS")
    If oAttribute Is Nothing Then
        sTemp = "FACE_THICKNESS"
        Set oAttribute = GetAttribute(oObject, "SMS_PROFILE_2D" & "||" & "FACE_THICKNESS")
        If oAttribute Is Nothing Then GoTo CleanUp 'No thickness against which to validate, just exit!
    End If
    
    dPartTotal = RadGetAttributeValue(oAttribute, sTemp, True, m_oSettingsXML)
    
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_BEVEL" & "||" & "CHAMFER_DEPTH_M", dCutTotal)
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_GRIND" & "||" & "CHAMFER_DEPTH_M", dCutTotal)
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_BEVEL" & "||" & "CHAMFER_DEPTH_UM", dCutTotal)
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_GRIND" & "||" & "CHAMFER_DEPTH_UM", dCutTotal)
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_BEVEL" & "||" & "DEPTH_A", dCutTotal)
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_GRIND" & "||" & "DEPTH_A", dCutTotal)
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_BEVEL" & "||" & "DEPTH_B", dCutTotal)
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_GRIND" & "||" & "DEPTH_B", dCutTotal)
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_BEVEL" & "||" & "DEPTH_N", dCutTotal)
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_GRIND" & "||" & "DEPTH_N", dCutTotal)
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_BEVEL" & "||" & "DEPTH_D", dCutTotal)
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_GRIND" & "||" & "DEPTH_D", dCutTotal)
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_BEVEL" & "||" & "DEPTH_E", dCutTotal)
    dCutTotal = AccumulateBevelGrindThickness(oObject, "SMS_GRIND" & "||" & "DEPTH_E", dCutTotal)
    
    If Abs(dCutTotal - dPartTotal) > MAXIMUM_THICKNESS_TOLERANCE Then
        'This is an error, we need to report it!
        sError = "Bevel or Grind thickness does not match the total thickness for the plate or face."
        sSolution = "Edit the value in the attribute editor."
        'No localization needed, if you need to do some localization, then just replace your localized string
        'to the above variables.
        Set oError = New STRMFGPartEditorError.CPartEditErrorObj
        eErrLvl = GetValidationErrorLevel(VALIDATE_CUSTOM_BEVEL_GRIND_ATTRIB)
        If eErrLvl < ERRORLEVEL_ErrInformation Then
            eErrLvl = ERRORLEVEL_ErrInformation
        ElseIf eErrLvl > ERRORLEVEL_ErrCritical Then
            eErrLvl = ERRORLEVEL_ErrCritical
        End If
        If sPurpose = "Contour" Then
            If oError.SetData(oObject, "Bevel-Grind Attributes", sError, sSolution, eErrLvl, _
            VALIDATE_CUSTOM_BEVEL_GRIND_ATTRIB, "") Then
                oErrorCollection.Add oError
            End If
        ElseIf sPurpose = "Marking" Then
            If oError.SetData(oObject, "Common Cut Bevel-Grind Attributes", sError, sSolution, eErrLvl, _
            VALIDATE_CUSTOM_BEVEL_GRIND_ATTRIB, "") Then
                oErrorCollection.Add oError
            End If
        End If
    End If
CleanUp:
    Set oAttribute = Nothing
    Set oError = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Sub

'********************************************************************
' Routine: AccumulateBevelGrindThickness
' Abstract: Will Add up all the Bevel & Grind thicknesses that are found for a given RAD object.
'********************************************************************
Private Function AccumulateBevelGrindThickness(oObj As Object, sPropName As String, dTotal As Double) As Double
    On Error GoTo ErrorHandler
    Const METHOD As String = "AccumulateBevelGrindThickness"
    Dim sVal            As String
    Dim oUnitsOfMeasure As IJUomVBInterface
    Set oUnitsOfMeasure = New UnitsOfMeasureServicesLib.UomVBInterface
    
    If Not m_oSettingsXML Is Nothing And Not oUnitsOfMeasure Is Nothing And Not oObj Is Nothing Then
        sVal = RadGetAttributeValue(oObj, sPropName, True, m_oSettingsXML)
        If IsNumeric(sVal) Then sVal = sVal & " m"
        If sVal = "" Then sVal = "0"
        dTotal = dTotal + oUnitsOfMeasure.ParseUnit(UNIT_DISTANCE, sVal)
    End If
CleanUp:
    AccumulateBevelGrindThickness = dTotal
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function
