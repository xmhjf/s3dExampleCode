﻿//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   FrameEqualHeightRule is a .NET process parameter rule for TemplateSet on Plate.
//                 
//
//      Author:  
//
//      History:
//      August 5th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.Plate
{
    /// <summary>
    /// Process FrameEqualHeight Parameter Rule for TemplateSet on Plate.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    #region RuleInterface
    [RuleInterface(SymbolsConstants.ProcessFrameEqualHeightParamName, SymbolsConstants.ProcessFrameEqualHeightParamUserName)]
    #endregion

    public class ProcessParameterFrameEqualHeightRule : PlateParameterRule
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
        /// The type
        /// </summary>
        [ControlledParameter(5, SymbolsConstants.ProcessType, SymbolsConstants.ProcessType)]
        public ControlledParameterDouble type;
        /// <summary>
        /// The orientation
        /// </summary>
        [ControlledParameter(6, SymbolsConstants.ProcessOrientation, SymbolsConstants.ProcessOrientation)]
        public ControlledParameterDouble orientation;
        /// <summary>
        /// The position
        /// </summary>
        [ControlledParameter(7, SymbolsConstants.ProcessPosition, SymbolsConstants.ProcessPosition)]
        public ControlledParameterDouble position;
        /// <summary>
        /// The base plane
        /// </summary>
        [ControlledParameter(8, SymbolsConstants.ProcessBasePlane, SymbolsConstants.ProcessBasePlane)]
        public ControlledParameterDouble basePlane;
        /// <summary>
        /// The direction
        /// </summary>
        [ControlledParameter(9, SymbolsConstants.ProcessDirection, SymbolsConstants.ProcessDirection)]
        public ControlledParameterDouble direction;
        /// <summary>
        /// The template service
        /// </summary>
        [ControlledParameter(10, SymbolsConstants.ProcessTemplateService, SymbolsConstants.ProcessTemplateService)]
        public ControlledParameterDouble templateService;
        /// <summary>
        /// The user defined values
        /// </summary>
        [ControlledParameter(11, SymbolsConstants.ProcessUserDefinedValues, SymbolsConstants.ProcessUserDefinedValues)]
        public ControlledParameterDouble userDefinedValues;
        /// <summary>
        /// The template naming
        /// </summary>
        [ControlledParameter(12, SymbolsConstants.ProcessTemplateNaming, SymbolsConstants.ProcessTemplateNaming)]
        public ControlledParameterDouble templateNaming;

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

                this.side.Value = base.ConcaveSide;
                this.type.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessType);
                this.orientation.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessOrientation);

                this.position.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessPosition);
                this.basePlane.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessBasePlane);
                this.direction.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessDirection);

                this.templateService.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessTemplateService);
                this.userDefinedValues.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessUserDefinedValues);
                this.templateNaming.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessTemplateNaming);

            }
            catch
            {
                //To Do
            }

        }
        #endregion
    }
}
