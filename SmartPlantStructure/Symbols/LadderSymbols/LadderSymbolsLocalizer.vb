'=====================================================================================================
'
'Copyright 2010 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'Ladder symbols localizer
'
'File
'  LadderSymbolsLocalizer.vb
'
'=====================================================================================================
Imports Ingr.SP3D.Common.Middle.Services

Namespace Ingr.SP3D.Content.Structure

    NotInheritable Class LadderSymbolsLocalizer
        Public Shared Function GetString(ByVal iID As Integer, ByVal defMsgStr As String) As String
            Return CmnLocalizer.GetString(iID, defMsgStr, LadderSymbolsResourceIDs.DEFAULT_RESOURCE, LadderSymbolsResourceIDs.DEFAULT_ASSEMBLY)
        End Function
    End Class

End Namespace
