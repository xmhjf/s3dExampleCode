VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlternatePublisher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
'Copyright © 2008, Intergraph Corporation
'
'Class:     AlternatePublisher
'
'Author:    Min Wang
'
'Abstract:  This class implements IJReportPublisher. It is called when the regular
'           catalog report does not exist, and the definition of display cannot be
'           known.
'
'Description:
'
'History:   First delivery October, 2008
'   10/15/2008  mw  DI152393 Deliver an Alternative Publisher to display the file
'******************************************************************************
Option Explicit

Private Const MODULE = "SP3DReportPublisher.AlternatePublisher:"

Implements IJReportPublisher

'To Hold properties
Private m_sReportOutput As String
Private m_oReportOutput As Object
Private m_oDisplayParameters As IJDisplayParameters
Private m_oCacheController As IJCacheController

'******************************************************************************
'Class Constructor/Destructor
'******************************************************************************
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    
    On Error GoTo ErrorHandler

    
    Exit Sub

ErrorHandler:
    

End Sub

Private Sub Class_Terminate()
    Set m_oReportOutput = Nothing
    Set m_oDisplayParameters = Nothing
End Sub


Private Property Set IJReportPublisher_CacheController(ByVal RHS As SP3DReportsRuntimeInterfaces.IJCacheController)
    Set m_oCacheController = RHS
End Property

Private Property Set IJReportPublisher_DisplayParameters(ByVal RHS As SP3DReportsObjectsInterfaces.IJDisplayParameters)
    Set m_oDisplayParameters = RHS
End Property

Private Property Get IJReportPublisher_DisplayParameters() As SP3DReportsObjectsInterfaces.IJDisplayParameters
    Set IJReportPublisher_DisplayParameters = m_oDisplayParameters
End Property

'Command may be "DisplayLocal", "Print"
'Argument may be "True" or a printer name
Private Sub IJReportPublisher_Execute(ByVal Command As String, ByVal Argument As String)
    Const METHOD = "IJReportPublisher_Execute"
    Dim oDisplayTool As IJReportPublisher
    Dim sProgId As String
    Dim sExtenson As String
    Dim iPos As Long
    
    On Error GoTo ErrorHandler
    
    If m_sReportOutput = "" Then
        Err.Description = "No file to display."
        GoTo ErrorHandler
    End If
        
    iPos = InStr(1, m_sReportOutput, ".")
    If iPos > 0 Then
        sExtenson = Mid(m_sReportOutput, iPos + 1)
        
        If StrComp(sExtenson, "xls", vbTextCompare) = 0 _
            Or StrComp(sExtenson, "xlsx", vbTextCompare) = 0 Then   'Excel
            
            sProgId = "SP3DExcelFormat.Publisher"
            
            Set oDisplayTool = SP3DCreateObject(sProgId)
            Set oDisplayTool.CacheController = m_oCacheController
            oDisplayTool.ReportOutput = m_sReportOutput

            oDisplayTool.Execute Command, Argument
            
        Else
            'other type of file
            'user may add code to handle
            Err.Description = "Do not know how to display: " & m_sReportOutput
            GoTo ErrorHandler
        End If
        
    Else
        'other type of file?
        Err.Description = "Do not know how to display: " & m_sReportOutput
        GoTo ErrorHandler
    End If
    GoTo ShutDown
    
ErrorHandler:
    Log.AddFromErr Err, Err.Description, METHOD, MODULE
    
ShutDown:
    Set oDisplayTool = Nothing
End Sub

Private Property Let IJReportPublisher_ReportOutput(ByVal RHS As String)
    m_sReportOutput = RHS
End Property

Private Property Get IJReportPublisher_ReportOutput() As String
    On Error Resume Next
    IJReportPublisher_ReportOutput = m_sReportOutput
End Property

Private Property Set IJReportPublisher_ReportOutputAsObject(ByVal RHS As Object)
    Set m_oReportOutput = RHS
End Property

Private Property Get IJReportPublisher_ReportOutputAsObject() As Variant
    Set IJReportPublisher_ReportOutputAsObject = m_oReportOutput
End Property
