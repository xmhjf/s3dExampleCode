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
        /// The seam mark
        /// </summary>
        [ControlledParameter(2, SymbolsConstants.MarkingSeamMark, SymbolsConstants.MarkingSeamMark)]
        public ControlledParameterDouble seamMark;
        /// <summary>
        /// The ship dir mark
        /// </summary>
        [ControlledParameter(3, SymbolsConstants.MarkingShipDirectionMark, SymbolsConstants.MarkingShipDirectionMark)]
        public ControlledParameterDouble shipDirMark;
        /// <summary>
        /// The custom mark
        /// </summary>
        [ControlledParameter(4, SymbolsConstants.MarkingCustomMark, SymbolsConstants.MarkingCustomMark)]
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
                //IPart partItem = (IPart)base.Part;

                //if ((partItem.PartNumber == SymbolsConstants.MarkingDefault) ||
                //    (partItem.PartNumber == SymbolsConstants.MarkingBox))
                //{

                this.baseControlMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingBaseCtlLineMark);
                this.seamMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingSeamMark);
                this.shipDirMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingShipDirectionMark);
                this.customMark.Value = base.GetCatalogValue(SymbolsConstants.MarkingParamInterface, SymbolsConstants.MarkingCustomMark);
                
                //}
                //else if (partItem.PartNumber == "XXX_TemplateMarkingPlate") 
                //{
                //    // User can overwrite the values. 
                //}

            }
            catch
            {
                //To Do
            }
        }
        #endregion
    }
}
