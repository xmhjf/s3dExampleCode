//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Symbols Constants for TemplateSet on Plate.
//                 
//
//      Author:  
//
//      History:
//      August 5th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.ProfileEdge
{
    /// <summary>
    /// Symbols Constants for TemplateSet on Plate.
    /// </summary>
    public static class SymbolsConstants
    {

        #region Profile Definition, SelectorRule, and ParameterRule

        internal const string ProcessSelectorName = "IAStrMfgTemplateSelEdge_SelRuleProc";
        internal const string ProcessSelectorUserName = "IAStrMfgTemplateSelEdge_SelRuleProc";

        internal const string MarkingSelectorName = "IAStrMfgTemplateSelEdge_SelRuleMark";
        internal const string MarkingSelectorUserName = "IAStrMfgTemplateSelectorPlate_SelRuleMrk";
        
        internal const string ProcessDefaultParamName = "IAStrMfgTemplateSelEdge_DefProcParEdge";
        internal const string ProcessDefaultParamUserName = "IAStrMfgTemplateSelEdge_DefProcParEdge";

        internal const string MarkingDefaultParamName = "IAStrMfgTemplateSelEdge_DefMarkParEdge";
        internal const string MarkingDefaultParamUserName = "IAStrMfgTemplateSelEdge_DefMarkParEdge";

        // TemplateSet Process Selector Constants

        internal const string ProcessDefault = "Default_TemplateProcessEdge";
                    
        // TemplateSet Marking Selector Constants

        internal const string MarkingDefault = "Default_TemplateMarkingEdge";  

        //TemplateSet Process Interface Name
        internal const string ProcessParamInterface = "IJUAMfgTemplateProcessEdge";

        //TemplateSet Process Constants
        internal const string ProcessType = "Type";
        internal const string ProcessSide = "Side";
        internal const string ProcessOffset = "Offset";
        internal const string ProcessExtension = "Extension";
        internal const string ProcessMinHeight = "MinHeight";
        internal const string ProcessMaxHeight = "MaxHeight";
        internal const string ProcessService = "TemplateService";
        internal const string ProcessUsrDefinedValue = "UserDefinedValues";
        internal const string ProcessTemplateName = "TemplateNaming";
                
        //TemplateSet Marking Interface Name
        internal const string MarkingParamInterface = "IJUAMfgTemplateMarkingEdge";

        //TemplateSet Marking Constants
        internal const string MarkingBaseCtlLineMark = "BaseCtrlLineMark";
        internal const string MarkingSeamMark = "SeamMarks";
        internal const string MarkingShipDirectionMark = "ShipDirectionMark";
        internal const string MarkingCustomMark = "CustomMark";
        
        // Representaion 
        internal const string ProcessAspect = "Process";
        internal const string ProcessAspectDesc = "Process";

        internal const string MarkingAspect = "Marking";
        internal const string MarkingAspectDesc = "Marking";

        #endregion
    }
}

