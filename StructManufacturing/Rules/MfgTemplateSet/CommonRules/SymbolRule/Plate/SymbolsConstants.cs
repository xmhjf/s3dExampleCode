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

namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.Plate
{
    /// <summary>
    /// Symbols Constants for TemplateSet on Plate.
    /// </summary>
    public static class SymbolsConstants
    {

        #region Plate Definition, SelectorRule, and ParameterRule

        internal const string ProcessSelectorName = "IAStrMfgTemplateSelectorPlate_SelRulePrc";
        internal const string ProcessSelectorUserName = "IAStrMfgTemplateSelectorPlate_SelRulePrc";

        internal const string ProcessDefaultParamName = "IAStrMfgTemplateSelectorPlate_PrcDefPar";
        internal const string ProcessDefaultParamUserName = "IAStrMfgTemplateSelectorPlate_PrcDefPar";

        internal const string ProcessFrameParamName = "IAStrMfgTemplateSelectorPlate_PrcFramePar";
        internal const string ProcessFrameParamUserName = "IAStrMfgTemplateSelectorPlate_PrcFramePar";

        internal const string ProcessFrameEqualHeightParamName = "IAStrMfgTemplateSelectorPlate_PrcEqHtPar";
        internal const string ProcessFrameEqualHeightParamUserName = "IAStrMfgTemplateSelectorPlate_PrcEqHtPar";

        internal const string ProcessEvenParamName = "IAStrMfgTemplateSelectorPlate_PrcEvenPar";
        internal const string ProcessEvenParamUserName = "IAStrMfgTemplateSelectorPlate_PrcEvenPar";

        internal const string ProcessCenterLineParamName = "IAStrMfgTemplateSelectorPlate_PrcCLPar";
        internal const string ProcessCenterLineParamUserName = "IAStrMfgTemplateSelectorPlate_PrcCLPar";

        internal const string ProcessPerpendicularParamName = "IAStrMfgTemplateSelectorPlate_PrcPerpPar";
        internal const string ProcessPerpendicularParamUserName = "IAStrMfgTemplateSelectorPlate_PrcPerpPar";

        internal const string ProcessStemSternParamName = "IAStrMfgTemplateSelectorPlate_PrcStSnPar";
        internal const string ProcessStemSternParamUserName = "IAStrMfgTemplateSelectorPlate_PrcStSnPar";

        internal const string ProcessPerpendicularXYParamName = "IAStrMfgTemplateSelectorPlate_PrcPrpXYPar";
        internal const string ProcessPerpendicularXYParamUserName = "IAStrMfgTemplateSelectorPlate_PrcPrpXYPar";

        internal const string ProcessUserDefinedParamName = "IAStrMfgTemplateSelectorPlate_PrcUsrDfPar";
        internal const string ProcessUserDefinedParamUserName = "IAStrMfgTemplateSelectorPlate_PrcUsrDfPar";

        internal const string ProcessAftForwardParamName = "IAStrMfgTemplateSelectorPlate_PrcAftFwPar";
        internal const string ProcessAftForwardParamUserName = "IAStrMfgTemplateSelectorPlate_PrcAftFwPar";

        internal const string ProcessBoxParamName = "IAStrMfgTemplateSelectorPlate_PrcBoxPar";
        internal const string ProcessBoxParamUserName = "IAStrMfgTemplateSelectorPlate_PrcBoxPar";

        internal const string ProcessUserDefinedBoxParamName = "IAStrMfgTemplateSelectorPlate_PrcUBoxPar";
        internal const string ProcessUserDefinedBoxParamUserName = "IAStrMfgTemplateSelectorPlate_PrcUBoxPar";

        internal const string ProcessUserDefBoxEdgesParamName = "IAStrMfgTemplateSelectorPlate_PrcUBoxEPar";
        internal const string ProcessUserDefBoxEdgesParamUserName = "IAStrMfgTemplateSelectorPlate_PrcUBoxEPar";

        internal const string MarkingSelectorName = "IAStrMfgTemplateSelectorPlate_SelRuleMrk";
        internal const string MarkingSelectorUserName = "IAStrMfgTemplateSelectorPlate_SelRuleMrk";

        internal const string MarkingDefaultParamName = "IAStrMfgTemplateSelectorPlate_MrkDefPar";
        internal const string MarkingDefaultParamUserName = "IAStrMfgTemplateSelectorPlate_MrkDefPar";

        internal const string MarkingBoxParamName = "IAStrMfgTemplateSelectorPlate_MrkBoxPar";
        internal const string MarkingBoxParamUserName = "IAStrMfgTemplateSelectorPlate_MrkBoxPar";

        // TemplateSet Process Selector Constants
        internal const string ProcessDefault = "Default_TemplateProcessPlate";
        internal const string ProcessFrame = "Frame_TemplateProcessPlate";
        internal const string ProcessFrameEqualHeight = "FrameEqualHeight_TemplateProcessPlate";

        internal const string ProcessEven = "Even_TemplateProcessPlate";
        internal const string ProcessCenterLine = "CenterLine_TemplateProcessPlate";
        internal const string ProcessPerpendicular = "Perpendicular_TemplateProcessPlate";

        internal const string ProcessStemStern = "StemStern_TemplateProcessPlate";
        internal const string ProcessPerpendicularXY = "PerpendicularXY_TemplateProcessPlate";
        internal const string ProcessUserDefined = "UserDefined_TemplateProcessPlate";

        internal const string ProcessAftForward = "AftForward_TemplateProcessPlate";
        internal const string ProcessBox = "Box_TemplateProcessPlate";
        internal const string ProcessUserDefinedBox = "UserDefinedBox_TemplateProcessPlate";
        internal const string ProcessUserDefBoxEdges = "UserDefBoxEdges_TemplateProcessPlate"; 

        // TemplateSet Marking Selector Constants
        internal const string MarkingDefault = "Default_TemplateMarkingPlate";  
        internal const string MarkingBox = "Box_TemplateMarkingPlate";

        //TemplateSet Process Interface Name
        internal const string ProcessParamInterface = "IJUAMfgTemplateProcessPlate";

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
        internal const string MarkingParamInterface = "IJUAMfgTemplateMarkingPlate";

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

