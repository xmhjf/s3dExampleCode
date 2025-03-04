//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   DefaultRule is a .NET marking parameter rule for TemplateSet on Plate
//                 
//
//      Author:  
//
//      History:
//      August 5th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.ProfileEdge
{
    /// <summary>
    /// Marking Default Parameter Rule for TemplateSet on ProfileEdge.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    #region RuleInterface
    [RuleInterface(SymbolsConstants.ProcessDefaultParamName, SymbolsConstants.ProcessDefaultParamUserName)]
    #endregion

    public class ProcessParameterDefaultRule : ProfileParameterRule
    {
        #region Parameters

        /// <summary>
        /// The type
        /// </summary>
        [ControlledParameter(1, SymbolsConstants.ProcessType, SymbolsConstants.ProcessType)]
        public ControlledParameterDouble type;
        /// <summary>
        /// The side
        /// </summary>
        [ControlledParameter(2, SymbolsConstants.ProcessSide, SymbolsConstants.ProcessSide)]
        public ControlledParameterDouble side;
        /// <summary>
        /// The offset
        /// </summary>
        [ControlledParameter(3, SymbolsConstants.ProcessOffset, SymbolsConstants.ProcessOffset)]
        public ControlledParameterDouble offset;
        /// <summary>
        /// The extension
        /// </summary>
        [ControlledParameter(4, SymbolsConstants.ProcessExtension, SymbolsConstants.ProcessExtension)]
        public ControlledParameterDouble extension;
        /// <summary>
        /// The minimum height
        /// </summary>
        [ControlledParameter(5, SymbolsConstants.ProcessMinHeight, SymbolsConstants.ProcessMinHeight)]
        public ControlledParameterDouble minHeight;
        /// <summary>
        /// The maximum height
        /// </summary>
        [ControlledParameter(6, SymbolsConstants.ProcessMaxHeight, SymbolsConstants.ProcessMaxHeight)]
        public ControlledParameterDouble maxHeight;
        /// <summary>
        /// The service
        /// </summary>
        [ControlledParameter(7, SymbolsConstants.ProcessService, SymbolsConstants.ProcessService)]
        public ControlledParameterDouble service;
        /// <summary>
        /// The user defined values
        /// </summary>
        [ControlledParameter(8, SymbolsConstants.ProcessUsrDefinedValue, SymbolsConstants.ProcessUsrDefinedValue)]
        public ControlledParameterDouble userDefinedValues;
        /// <summary>
        /// The template name
        /// </summary>
        [ControlledParameter(9, SymbolsConstants.ProcessTemplateName, SymbolsConstants.ProcessTemplateName)]
        public ControlledParameterDouble templateName;

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

                this.type.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessType);
                this.side.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessSide);
                this.offset.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessOffset);

                this.extension.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessExtension);
                this.minHeight.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessMinHeight);
                this.maxHeight.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessMaxHeight);
                this.service.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessService);

                this.userDefinedValues.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessUsrDefinedValue);
                this.templateName.Value = base.GetCatalogValue(SymbolsConstants.ProcessParamInterface, SymbolsConstants.ProcessTemplateName);

            }
            catch
            {
                //To Do
            }
        }
        #endregion
    }
}
