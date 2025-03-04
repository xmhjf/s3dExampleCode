VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeldCompYes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Module:       WeldCompYes
'
' Description:  Apply Weld Compensation with Default Parameters
'
' Comments:
'' 8 Feb, 2008    Creation
'
'*******************************************************************************
' Logic for computing default parameters:
'
'   Defaults:
'
'   Len < 15000     p1 = 30     p2 = 0.5 * p1 * Ht/Len  p3 = -p2
'   Len > 15000     p1 = 45     p2 = 0.5 * p1 * Ht/Len  p3 = -p2
'                   Param1 = Param 2 = 0.5
'
'   Note a,
'
'       If we have a top flange and a bottom flange,
'           apply compensation on the flange with larger height
'
'   Note b,
'
'       If we have a "clear" inflexion along the landing curve,
'           set Param1 = Param2 = tangency point
'
'-------------------------------------------------------------------------------

Option Explicit

Implements IJDMfgSystemMarkRule

Private m_oProfilePart As IJProfilePart

Private m_oWC As IJMfgWeldCompensation
Private m_oWCFactory As IJMfgWeldCompensationFactory

Private m_oProfileHelper As Object
Private m_oProfileSupport As IJProfilePartSupport
Private m_oPartSupport As IJPartSupport

Const MODULE = "MfgMemberProcess.WeldCompYes"


Private Function IJDMfgSystemMarkRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
'
End Function

Private Function IJDMfgSystemMarkRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "IJDMfgSystemMarkRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
        

CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2038, , "RULES")
End Function
