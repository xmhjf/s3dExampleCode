using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.Tube
{
    /// <summary>
    /// Symbols Constants for TemplateSet on Tube.
    /// </summary>
    public static class SymbolsConstants
    {
        internal const string ProcessSelectorName = "IAStrMfgTemplateSelTube_SelRuleProc";
        internal const string ProcessSelectorUserName = "IAStrMfgTemplateSelTube_SelRuleProc";
        
        internal const string ProcessDefaultParamName = "IAStrMfgTemplateSelTube_DefaultProcPar";
        internal const string ProcessDefaultParamUserName = "IAStrMfgTemplateSelTube_DefaultProcPar";

        internal const string ProcessLongDistBCLParamName = "IAStrMfgTemplateSelTube_LongDstBCLProcPar";
        internal const string ProcessLongDistBCLParamUserName = "IAStrMfgTemplateSelTube_LongDstBCLProcPar";

        internal const string ProcessShortDistBCLParamName = "IAStrMfgTemplateSelTube_ShrtDstBCLProcPar";
        internal const string ProcessShortDistBCLParamUserName = "IAStrMfgTemplateSelTube_ShrtDstBCLProcPar";

        internal const string MarkingDefaultParamName = "IAStrMfgTemplateSelTube_DefaultMarkPar";
        internal const string MarkingDefaultParamUserName = "IAStrMfgTemplateSelTube_DefaultMarkPar";

        internal const string MarkingSelectorName = "IAStrMfgTemplateSelTube_SelRuleMark";
        internal const string MarkingSelectorUserName = "IAStrMfgTemplateSelTube_SelRuleMark";

        //TemplateSet Process Interface Name
        internal const string ProcessParamInterface = "IJUAMfgTemplateProcessTube";

        //TemplateSet Marking Interface Name
        internal const string MarkingParamInterface = "IJUAMfgTemplateMarkingTube";

        // TemplateSet Process Selector Constants
        internal const string ProcessDefault = "Default_TemplateProcessTube";
        internal const string ProcessShortDistBCL = "ShortDistBCL_TemplateProcessTube";
        internal const string ProcessLongDistBCL = "LongDistBCL_TemplateProcessTube";

        //TemplateSet Process Constants
        internal const string ProcessMinHeight = "MinHeight";
        internal const string ProcessMaxHeight = "MaxHeight";
        internal const string ProcessExtension = "Extension";
        internal const string ProcessSide = "Side";
        internal const string ProcessSideEnd = "SideEnd";
        internal const string ProcessTemplateService = "TemplateService";
        internal const string ProcessTemplateNaming = "TemplateNaming";
        internal const string ProcessType = "Type";
        internal const string ProcessUserDefinedValues = "UserDefinedValues";
        internal const string ProcessBevel = "Bevel";
        internal const string ProcessBasePlane = "BasePlane";

        // TemplateSet Marking Selector Constants
        internal const string MarkingDefault = "Default_TemplateMarkingTube"; 

        //TemplateSet Marking Constants
        internal const string MarkingSeamMarks = "SeamMarks";
        internal const string MarkingBaseCtlLineMark = "BaseCtrlLineMark";
        internal const string MarkingShipDirectionMark = "ShipDirectionMark";
        internal const string MarkingCustomMarks = "CustomMark";
        internal const string MarkingFrameMarks = "FrameMarks";
        internal const string MarkingFittingMarks = "FittingMark";
        internal const string MarkingQuarterLineMarks = "QuarterLineMarks";

        // Representaion 
        internal const string ProcessAspect = "Process";
        internal const string ProcessAspectDesc = "Process";

        internal const string MarkingAspect = "Marking";
        internal const string MarkingAspectDesc = "Marking";
    }
}
