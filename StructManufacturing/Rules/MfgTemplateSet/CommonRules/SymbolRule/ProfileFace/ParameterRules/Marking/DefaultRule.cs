//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   DefaultRule is a .NET marking parameter rule for TemplateSet on Profile Face
//                 
//
//      Author:  
//
//      History:
//      August 15th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.ProfileFace
{
    /// <summary>
    /// Marking Default Parameter Rule for TemplateSet on Profile Face.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    #region RuleInterface
    [RuleInterface(SymbolsConstants.MarkingDefaultParamName, SymbolsConstants.MarkingDefaultParamUserName)]
    #endregion

    public class MarkingParameterDefaultRule : ProfileParameterRule
    {  
        #region Parameters

        /// <summary>
        /// The side mark
        /// </summary>
        [ControlledParameter(1, SymbolsConstants.MarkingSideMark, SymbolsConstants.MarkingSideMark)]
        public ControlledParameterDouble sideMark;
        /// <summary>
        /// The seam mark
        /// </summary>
        [ControlledParameter(2, SymbolsConstants.MarkingSeamMarks, SymbolsConstants.MarkingSeamMarks)]
        public ControlledParameterDouble seamMark;
        /// <summary>
        /// The base control mark
        /// </summary>
        [ControlledParameter(3, SymbolsConstants.MarkingBaseCtlLineMark, SymbolsConstants.MarkingBaseCtlLineMark)]
        public ControlledParameterDouble baseControlMark;
        /// <summary>
        /// The sight line mark
        /// </summary>
        [ControlledParameter(4, SymbolsConstants.MarkingSightLineMark, SymbolsConstants.MarkingSightLineMark)]
        public ControlledParameterDouble sightLineMark;
        /// <summary>
        /// The ship direction mark
        /// </summary>
        [ControlledParameter(5, SymbolsConstants.MarkingShipDirectionMark, SymbolsConstants.MarkingShipDirectionMark)]
        public ControlledParameterDouble shipDirectionMark;
        /// <summary>
        /// The label mark
        /// </summary>
        [ControlledParameter(6, SymbolsConstants.MarkingLabelMark, SymbolsConstants.MarkingLabelMark)]
        public ControlledParameterDouble labelMark;
        /// <summary>
        /// The frame mark
        /// </summary>
        [ControlledParameter(7, SymbolsConstants.MarkingFrameMarks, SymbolsConstants.MarkingFrameMarks)]
        public ControlledParameterDouble frameMark;
        /// <summary>
        /// The knuckle mark
        /// </summary>
        [ControlledParameter(8, SymbolsConstants.MarkingKnuckleMarks, SymbolsConstants.MarkingKnuckleMarks)]
        public ControlledParameterDouble knuckleMark;
        /// <summary>
        /// The reference mark
        /// </summary>
        [ControlledParameter(9, SymbolsConstants.MarkingReferenceMarks, SymbolsConstants.MarkingReferenceMarks)]
        public ControlledParameterDouble referenceMark;
        /// <summary>
        /// The template mark
        /// </summary>
        [ControlledParameter(10, SymbolsConstants.MarkingTemplateMarks, SymbolsConstants.MarkingTemplateMarks)]
        public ControlledParameterDouble templateMark;
        /// <summary>
        /// The custom mark
        /// </summary>
        [ControlledParameter(11, SymbolsConstants.MarkingCustomMarks, SymbolsConstants.MarkingCustomMarks)]
        public ControlledParameterDouble customMark;

        #endregion Parameters

        #region overriden methods
        /// <summary>
        /// To be implemented by the user custom parameter rule to evaluate
        ///             the parameter rules.
        /// </summary>
        public override void Evaluate()
        {
            try
            {
                    this.sideMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingSideMark);
                    this.seamMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingSeamMarks);
                    this.baseControlMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingBaseCtlLineMark);

                    this.sightLineMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingSightLineMark);
                    this.shipDirectionMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingShipDirectionMark);
                    this.labelMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingLabelMark);

                    this.frameMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingFrameMarks);
                    this.knuckleMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingKnuckleMarks);
                    this.referenceMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingReferenceMarks);

                    this.templateMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingTemplateMarks);
                    this.customMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingCustomMarks); 

            }
            catch  
            {
                //To Do
            }
        }
        #endregion
    }
}
