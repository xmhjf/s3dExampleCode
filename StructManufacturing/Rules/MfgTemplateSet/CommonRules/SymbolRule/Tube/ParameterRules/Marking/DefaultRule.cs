//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   BoxRule is a .NET marking parameter rule for TemplateSet on Plate
//                 
//
//      Author:  
//
//      History:
//      August 5th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.Tube
{
    /// <summary>
    /// Marking Box Parameter Rule for TemplateSet on Tube.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    #region RuleInterface
    [RuleInterface(SymbolsConstants.MarkingDefaultParamName, SymbolsConstants.MarkingDefaultParamUserName)]
    #endregion

    public class MarkingParameterDefaultRule : ProfileParameterRule
    {
        #region Parameters

        /// <summary>
        /// The base control mark
        /// </summary>
        [ControlledParameter(1, SymbolsConstants.MarkingBaseCtlLineMark, SymbolsConstants.MarkingBaseCtlLineMark)]
        public ControlledParameterDouble baseControlMark;
        /// <summary>
        /// The fitting mark
        /// </summary>
        [ControlledParameter(2, SymbolsConstants.MarkingFittingMarks, SymbolsConstants.MarkingFittingMarks)]
        public ControlledParameterDouble fittingMark;
        /// <summary>
        /// The frame mark
        /// </summary>
        [ControlledParameter(3, SymbolsConstants.MarkingFrameMarks, SymbolsConstants.MarkingFrameMarks)]
        public ControlledParameterDouble frameMark;
        /// <summary>
        /// The quarter line mark
        /// </summary>
        [ControlledParameter(4, SymbolsConstants.MarkingQuarterLineMarks, SymbolsConstants.MarkingQuarterLineMarks)]
        public ControlledParameterDouble quarterLineMark;
        /// <summary>
        /// The seam mark
        /// </summary>
        [ControlledParameter(5, SymbolsConstants.MarkingSeamMarks, SymbolsConstants.MarkingSeamMarks)]
        public ControlledParameterDouble seamMark;
        /// <summary>
        /// The ship direction mark
        /// </summary>
        [ControlledParameter(6, SymbolsConstants.MarkingShipDirectionMark, SymbolsConstants.MarkingShipDirectionMark)]
        public ControlledParameterDouble shipDirectionMark;
        /// <summary>
        /// The custom mark
        /// </summary>
        [ControlledParameter(7, SymbolsConstants.MarkingCustomMarks, SymbolsConstants.MarkingCustomMarks)]
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

                this.seamMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingSeamMarks);
                this.baseControlMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingBaseCtlLineMark);
                this.shipDirectionMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingShipDirectionMark);

                this.customMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingCustomMarks);
                this.frameMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingFrameMarks);
                this.fittingMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingFittingMarks);

                this.quarterLineMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingQuarterLineMarks);

            }
            catch
            {
                //To Do
            }
        }
        #endregion
    }
}
