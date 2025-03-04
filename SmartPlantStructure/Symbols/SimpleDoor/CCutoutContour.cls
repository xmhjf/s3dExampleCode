VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCutoutContour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CCutoutContour.cls
'   Author:          Babu Govindarajan
'   Creation Date:  Friday, September 19 - 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IA2,
'   PDMS Category Name=IFAH). The symbol consists of five oupputs: Three physical,
'   and two insulation aspect outputs.
'
'   Note: This symbol expects the user to place an Olet component on the pipe and this (Annubar) should be placed
'   at the connect point2  of the Olet.The parameters values are to be taken from connect point2 of the Olet.'
'   The port for the symbol is fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   05.Sep.2003     Babu Govindarajan   TR-CP 49538  Created new class file CCutoutContour.cls to add operation aspect to the Symbol.
'   05.Apr.2006     Regis Dechance                     Adaptation for Doors, Windows, Louvers
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CutoutContour:" 'Used for error messages
Private Sub Class_Initialize()
'''
End Sub
'Private m_oCodeListMetadata     As IJDCodeListMetaData
'

Private Sub Class_Terminate()
'    Set m_oCodeListMetadata = Nothing
End Sub

Sub Run(ByVal pOutputCollection As IJDOutputCollection, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "Run:"
    On Error GoTo ErrorHandler
    
    Call RunForDoorsWindowsOperationAspect(pOutputCollection, arrayOfInputs(), arrayOfOutputs(), True)
        
    Exit Sub
        
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
 
End Sub
    
