VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SmallContourSeg"
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

Private Const PI As Double = 3.14159265358979
Private m_strXMLData As String
Private Const MINIMUM_SEG_LENGTH = 0.003 '3mm in meters

Private Const MODULE As String = "CustomValidation.SmallContourSeg"

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
    Dim oErrorCollection As New Collection
    Dim oRad2dDoc As RAD2D.Document
    If oRadDoc Is Nothing Or sXMLSettingsDoc = "" Then GoTo CleanUp
    
    ValidateAgainstSmallContourSegmentsMgr oErrorCollection, oRadDoc, sXMLSettingsDoc
    Set IJMfgEditCustomValidation_RunCustomValidate = oErrorCollection
CleanUp:
    Set oErrorCollection = Nothing
    Set oRad2dDoc = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
' Routine: ValidateAgainstSmallContourSegments
' Abstract: Handles the validation against small contour segments.
'********************************************************************
Private Sub ValidateAgainstSmallContourSegmentsMgr(oErrorCollection As Collection, oRadDoc As RAD2D.Document, _
sXMLSettingsDoc As String)
    On Error GoTo ErrorHandler
    Const METHOD As String = "ValidateAgainstSmallContourSegments"
    
    Dim oAttribSets As RAD2D.AttributeSets
    Dim oSheet As RAD2D.Sheet
    Dim oOuterCollection As Collection
    Dim oInnerCollection As Collection
    Dim oErrColObjs As Collection
    Dim oErrObj As STRMFGPartEditorError.CPartEditErrorObj
    
    If oErrorCollection Is Nothing Then GoTo CleanUp
    Set oErrColObjs = New Collection
    
    Set oOuterCollection = GetOuterContoursCollection(oRadDoc.ActiveSheet)
    Set oInnerCollection = GetInnerContoursCollection(oRadDoc.ActiveSheet)
    ValidateAgainstSmallContourSegments oOuterCollection, oErrorCollection, oErrColObjs, oRadDoc
    ValidateAgainstSmallContourSegments oInnerCollection, oErrorCollection, oErrColObjs, oRadDoc
    For Each oErrObj In oErrorCollection
        If Not oErrObj Is Nothing Then
            On Error Resume Next
            'Try to grab group objects rather than geometry primitives for
            'error reporting where there are several layers of groups
            If TypeOf oErrObj.RadObject.Parent.Parent Is RAD2D.Group Then
                oErrObj.RadObject = oErrObj.RadObject.Parent
            End If
            On Error GoTo ErrorHandler
        End If
    Next oErrObj
CleanUp:
    'oErrColObjs is only needed while running validation
    'to ensure we don't report objects more then once.
    Set oAttribSets = Nothing
    Set oSheet = Nothing
    Set oOuterCollection = Nothing
    Set oInnerCollection = Nothing
    Set oErrColObjs = Nothing
    Set oErrObj = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Sub

Private Sub ValidateAgainstSmallContourSegments(oOutContourCol As Collection, oErrorCollection As Collection, _
oErrColObjs As Collection, oRadDoc As RAD2D.Document)
    Const METHOD As String = "ValidateAgainstSmallContourSegments"
    On Error GoTo ErrorHandler
    
    Dim oObject                 As Object
    Dim oOutGroup               As Group
    Dim oInGroup                As Group
    Dim oValidateCollection     As Collection
    Dim oTempColl               As Collection
    Dim oErrObj                 As STRMFGPartEditorError.CPartEditErrorObj
    
    If oErrorCollection Is Nothing Then GoTo CleanUp
    
    If oRadDoc Is Nothing Then GoTo CleanUp
    Set oErrColObjs = New Collection
    
    FlattenGroups oOutContourCol, oOutGroup, oValidateCollection, oRadDoc
    FindSmallContourSegmentErrors oValidateCollection, oErrorCollection, oErrColObjs
    UnFlattenGroups oValidateCollection, oOutGroup, oRadDoc
CleanUp:
    Set oObject = Nothing
    Set oOutGroup = Nothing
    Set oInGroup = Nothing
    Set oValidateCollection = Nothing
    Set oTempColl = Nothing
    'Set oInContourCol = Nothing
    Set oOutContourCol = Nothing
    Set oErrObj = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Sub

Private Sub FindSmallContourSegmentErrors(oContourCollection As Collection, oErrorCollection As Collection, _
oErrColObjs As Collection)
    Const METHOD As String = "FindSmallContourSegmentErrors"
    On Error GoTo ErrorHandler
    Dim oObject         As Object
    Dim dTotalObjLength As Double
    Dim sError          As String
    Dim sSolution       As String
    Dim bErrorFound     As Boolean
    
    If oContourCollection Is Nothing Or oErrorCollection Is Nothing Then GoTo CleanUp
    If oContourCollection.Count <= 1 Then GoTo CleanUp
    
    For Each oObject In oContourCollection
        bErrorFound = False
        If Not oObject Is Nothing Then
            dTotalObjLength = GetTotalLengthOfObject(oObject)
            If dTotalObjLength < MINIMUM_SEG_LENGTH Then
                bErrorFound = True
            End If
        End If
        If bErrorFound = True Then
            sError = "Contour segment is to small."
            sSolution = "Adjust the start and end points of the segment to meet the minimum length requirement."
            ReportRADError oObject, oErrorCollection, oErrColObjs, sError, sSolution, "Small Contour Segments", _
                            ERRORLEVEL_ErrCritical, "ValidateSmallContourSeg"
        End If
    Next oObject
CleanUp:
    Set oObject = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Sub
