VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemNamingRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************************************
'  Copyright (C) 2001, Intergraph Corporation.  All rights reserved.
'
'  Project: NameRules
'
'  Abstract: The file contains a sample implementation of a naming rule
'
'  Author: Kamrooz K. Hormoozi oss/kaho
'  History
'   May 4 2002           oss\kaho   Created.
'*****************************************************************

Implements IJNameRule
                                                        
Private Const Module = "KoreanNamingRule"
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Private m_oErrors As IJEditErrors       ' To collect and propagate the errors.
Private m_oError As IJEditError         ' The error to raise.
Private Const TKWorkingSet = "WorkingSet"
Private Const TKApplicationContext = "ApplicationContext"
Private Const E_ACCESSDENIED As Long = 70
Private Sub Class_Initialize()
   Set m_oErrors = CreateObject(ERRORPROGID)
End Sub

Private Sub Class_Terminate()
   Set m_oErrors = Nothing
   Set m_oError = Nothing
End Sub

'********************************************************************
' Description:
' Creates a name for the object passed in. The name is based on the  the syatem name of the parts
' contain in the object
' It is assumed that all Naming Parents and the Object implement IJNamedItem.
' The Naming Parents are added in AddNamingParents() of the same interface.
' Both these methods are called from naming rule semantic.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName."
    On Error GoTo ErrorHandler
    

    Dim oEntityName As IJNamedItem
    Dim oPartName As IJNamedItem
    Dim strPartname As String

    If oObject Is Nothing Then
        Err.Raise E_ACCESSDENIED, Module, METHOD
    End If
    
    If elements.Count > 0 Then
       Dim oPart As Object
       ' The Parent System of a child of the content of Assembly
       Set oPart = elements.Item(1)
       Set oPartName = oPart
       ' The Named of the Parent System
       strPartname = oPartName.Name
    End If
    
    ' The Assembly
    Set oEntityName = oObject
    'Assign the name to Assembly
    oEntityName.Name = strPartname
   
   

Exit Sub
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' All the Naming Parents that need to participate in an objects naming are added here to the
' IJElements collection. The parent added here is the syatem which is parent system to contents of
' the assembly. Both these methods are called from
' naming rule semantic.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

      Dim oAssembly As IJAssembly
      Set oAssembly = oEntity
      Dim oSystemChild As IJSystemChild
      Dim oObject As Object
       ' Register list of children in selection list
      Dim oChildren As IJDTargetObjectCol
      Set oChildren = oAssembly.GetChildren
    
      Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

      Dim Index As Long
      For Index = 1 To oChildren.Count
        ' Get next child
        Dim oChild As GSCADAsmHlpers.IJAssemblyChild
        Set oChild = oChildren.Item(Index)
        
        'Get real parent to be used later on for naming of Assembly
        If TypeOf oChild Is IJSystemChild Then
            If TypeOf oChild Is IJSystem Then
                      Set oSystemChild = oChild
                      Set oObject = oSystemChild.GetParent
                      Dim SysChild As Boolean
                      SysChild = False
                      While Not SysChild
                            Set oSystemChild = oObject
                            Set oObject = oSystemChild.GetParent
                            If TypeOf oObject Is IJConfigProjectRoot Then
                               Set oObject = oSystemChild
                               SysChild = True
                            End If
                      Wend
            ElseIf Not TypeOf oChild Is IJSystem Then
                      Set oSystemChild = oChild
                      Set oObject = oSystemChild.GetParent
                      
            End If
            Call IJNameRule_GetNamingParents.Add(oObject)
            Exit For
            
        End If
    
      Next Index

Exit Function
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
    
End Function

