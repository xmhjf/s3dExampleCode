VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IsPlateIdentical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2009, Intergraph Corp.  All rights reserved.
'
' Project: MfgCustomFilter
' Module: CMfgIsIdentical
'
' Description:  To customize the filter criteria during the manufacture of plates.
'  History:
'       Suma Mallena         May 06. 2009   created
'
'*******************************************************************************

Option Explicit

'******************************************************************************
' Routine: IsPlatePartIdentical
'
' Description:  Through this function, user can send the filter criteria for the plates to be manufactured.
'
'******************************************************************************

Public Function IsPlatePartIdentical(elem As Object) As Integer
    IsPlatePartIdentical = 0
    
    ' your code goes here
End Function


 
