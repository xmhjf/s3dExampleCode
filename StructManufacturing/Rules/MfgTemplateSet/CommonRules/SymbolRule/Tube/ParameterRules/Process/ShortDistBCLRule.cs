//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   AftForwardRule is a .NET process parameter rule for TemplateSet on Plate.
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
    /// ShortDistBCLRule Rule for TemplateSet on Tube.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    #region RuleInterface
    [RuleInterface(SymbolsConstants.ProcessShortDistBCLParamName, SymbolsConstants.ProcessShortDistBCLParamName)]
    #endregion

    public class ProcessParameterShortDistBCLRuleRule : ProfileParameterRule
    {
        #region Parameters

        /// <summary>
        /// The minimum height
        /// </summary>
        [ControlledParameter(1, SymbolsConstants.ProcessMinHeight, SymbolsConstants.ProcessMinHeight)]
        public ControlledParameterDouble minHeight;
        /// <summary>
        /// The maximum height
        /// </summary>
        [ControlledParameter(2, SymbolsConstants.ProcessMaxHeight, SymbolsConstants.ProcessMaxHeight)]
        public ControlledParameterDouble maxHeight;
        /// <summary>
        /// The extension
        /// </summary>
        [ControlledParameter(3, SymbolsConstants.ProcessExtension, SymbolsConstants.ProcessExtension)]
        public ControlledParameterDouble extension;
        /// <summary>
        /// The side
        /// </summary>
        [ControlledParameter(4, SymbolsConstants.ProcessSide, SymbolsConstants.ProcessSide)]
        public ControlledParameterDouble side;
        /// <summary>
        /// The side end
        /// </summary>
        [ControlledParameter(5, SymbolsConstants.ProcessSideEnd, SymbolsConstants.ProcessSideEnd)]
        public ControlledParameterDouble sideEnd;
        /// <summary>
        /// The type
        /// </summary>
        [ControlledParameter(6, SymbolsConstants.ProcessType, SymbolsConstants.ProcessType)]
        public ControlledParameterDouble type;
        /// <summary>
        /// The base plane
        /// </summary>
        [ControlledParameter(7, SymbolsConstants.ProcessBasePlane, SymbolsConstants.ProcessBasePlane)]
        public ControlledParameterDouble basePlane;
        /// <summary>
        /// The template service
        /// </summary>
        [ControlledParameter(8, SymbolsConstants.ProcessTemplateService, SymbolsConstants.ProcessTemplateService)]
        public ControlledParameterDouble templateService;
        /// <summary>
        /// The user defined values
        /// </summary>
        [ControlledParameter(9, SymbolsConstants.ProcessUserDefinedValues, SymbolsConstants.ProcessUserDefinedValues)]
        public ControlledParameterDouble userDefinedValues;
        /// <summary>
        /// The template naming
        /// </summary>
        [ControlledParameter(10, SymbolsConstants.ProcessTemplateNaming, SymbolsConstants.ProcessTemplateNaming)]
        public ControlledParameterDouble templateNaming;
        /// <summary>
        /// The bevel
        /// </summary>
        [ControlledParameter(11, SymbolsConstants.ProcessBevel, SymbolsConstants.ProcessBevel)]
        public ControlledParameterDouble bevel;

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
                this.minHeight.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessMinHeight);
                this.maxHeight.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessMaxHeight);
                this.extension.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessExtension);

                this.type.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessType);
                this.side.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessSide);

                this.sideEnd.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessSideEnd);
                this.basePlane.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessBasePlane);
                this.templateService.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessTemplateService);

                this.userDefinedValues.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessUserDefinedValues);
                this.templateNaming.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessTemplateNaming);
                this.bevel.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessBevel);

            }
            catch
            {
                //To Do
            }

        }
        #endregion
    }
}
