//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Symbols Constants for TemplateSet on Profile Face.
//                 
//
//      Author:  
//
//      History:
//      August 15th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.ProfileFace
{
    /// <summary>
    /// Symbols Constants for TemplateSet on Profile Face.
    /// </summary>
    public static class SymbolsConstants
    {

        #region Profile Definition, SelectorRule, and ParameterRule

        internal const string ProcessSelectorName = "IAStrMfgTemplateSelectorProfile_SelRuleP";
        internal const string ProcessSelectorUserName = "IAStrMfgTemplateSelectorProfile_SelRuleP";

        internal const string ProcessDefaultParamName = "IAStrMfgTemplateSelectorProfile_PrcDefPR";
        internal const string ProcessDefaultParamUserName = "IAStrMfgTemplateSelectorProfile_PrcDefPR";

        internal const string ProcessFrameParamName = "IAStrMfgTemplateSelectorProfile_PrcFramPR";
        internal const string ProcessFrameParamUserName = "IAStrMfgTemplateSelectorProfile_PrcFramPR";

        internal const string ProcessEvenParamName = "IAStrMfgTemplateSelectorProfile_PrcEvenPR";
        internal const string ProcessEvenParamUserName = "IAStrMfgTemplateSelectorProfile_PrcEvenPR";

        internal const string ProcessPerpendicularParamName = "IAStrMfgTemplateSelectorProfile_PrcPerpPR";
        internal const string ProcessPerpendicularParamUserName = "IAStrMfgTemplateSelectorProfile_PrcPerpPR";

        internal const string ProcessUserDefinedParamName = "IAStrMfgTemplateSelectorProfile_PrcUsrDPR";
        internal const string ProcessUserDefinedParamUserName = "IAStrMfgTemplateSelectorProfile_PrcUsrDPR";

        internal const string ProcessAftForwardParamName = "IAStrMfgTemplateSelectorProfile_PrcAFPR";
        internal const string ProcessAftForwardParamUserName = "IAStrMfgTemplateSelectorProfile_PrcAFPR";

        internal const string MarkingSelectorName = "IAStrMfgTemplateSelectorProfile_SelRuleM";
        internal const string MarkingSelectorUserName = "IAStrMfgTemplateSelectorProfile_SelRuleM";

        internal const string MarkingDefaultParamName = "IAStrMfgTemplateSelectorProfile_MrkDefPR";
        internal const string MarkingDefaultParamUserName = "IAStrMfgTemplateSelectorProfile_MrkDefPR";


        // TemplateSet Process Selector Constants
        internal const string ProcessDefault = "Default_TemplateProcessProfile";
        internal const string ProcessFrame = "Frame_TemplateProcessProfile";               
        internal const string ProcessPerpendicular = "Perpendicular_TemplateProcessProfile";
        internal const string ProcessUserDefined = "UserDefined_TemplateProcessProfile";
        internal const string ProcessEven = "Even_TemplateProcessProfile"; 
        internal const string ProcessAftForward = "AftForward_TemplateProcessProfile";
        

        // TemplateSet Marking Selector Constants
        internal const string MarkingDefault = "Default_TemplateMarkingProfile"; 

        //TemplateSet Process Interface Name
        internal const string ProcessParamInterface = "IJUAMfgTemplateProcessProfile";

        //TemplateSet Process Constants
        internal const string ProcessMinHeight = "MinHeight";
        internal const string ProcessMaxHeight = "MaxHeight";
        internal const string ProcessExtension = "Extension";
        internal const string ProcessSide = "Side";
        internal const string ProcessType = "Type";
        internal const string ProcessOrientation = "Orientation";
        internal const string ProcessPosition = "Position";
        internal const string ProcessBasePlane = "BasePlane";
        internal const string ProcessDirection = "Direction";
        internal const string ProcessTemplateService = "TemplateService";
        internal const string ProcessUserDefinedValues = "UserDefinedValues";
        internal const string ProcessTemplateNaming = "TemplateNaming";


        //TemplateSet Marking Interface Name
        internal const string MarkingParamInterface = "IJUAMfgTemplateMarkingProfile";

        //TemplateSet Marking Constants
        internal const string MarkingSideMark = "SideMark";
        internal const string MarkingSeamMarks = "SeamMarks";
        internal const string MarkingBaseCtlLineMark = "BaseCtlLineMark";
        internal const string MarkingSightLineMark = "SightLineMark";
        internal const string MarkingShipDirectionMark = "ShipDirectionMark";
        internal const string MarkingLabelMark = "LabelMark";
        internal const string MarkingFrameMarks = "FrameMarks";
        internal const string MarkingKnuckleMarks = "KnuckleMarks";
        internal const string MarkingReferenceMarks = "ReferenceMarks";
        internal const string MarkingTemplateMarks = "TemplateMarks";
        internal const string MarkingCustomMarks = "CustomMarks";

        // Representaion 
        internal const string ProcessAspect = "Process";
        internal const string ProcessAspectDesc = "Process";

        internal const string MarkingAspect = "Marking";
        internal const string MarkingAspectDesc = "Marking";

        #endregion
    }
}

