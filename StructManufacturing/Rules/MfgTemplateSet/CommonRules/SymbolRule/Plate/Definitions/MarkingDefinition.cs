//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   MarkingDefinition is a .NET Symbol Definition for TemplateSet on Plate
//                 
//
//      Author:  
//
//      History:
//      August 5th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.Plate
{
    /// <summary>
    /// Marking Symbol Definition for TemplateSet on Plate.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    public class MarkingDefinition : ManufacturingSymbolDefinition
    {
        //Define 11 inputs
        #region "Definition of Inputs"
        /// <summary>
        /// Side Mark
        /// </summary>
        [InputDouble(1, SymbolsConstants.MarkingSideMark, SymbolsConstants.MarkingSideMark, 5200)] //Apply
        public InputDouble sideMark;
        /// <summary>
        /// Seam Marks
        /// </summary>
        [InputDouble(2, SymbolsConstants.MarkingSeamMarks, SymbolsConstants.MarkingSeamMarks, 5210)] //Apply
        public InputDouble seamMark;
        /// <summary>
        /// Base Control Line Mark 
        /// </summary>
        [InputDouble(3, SymbolsConstants.MarkingBaseCtlLineMark, SymbolsConstants.MarkingBaseCtlLineMark, 5220)] //Apply
        public InputDouble baseControlMark;
        /// <summary>
        /// Sight Line Mark
        /// </summary>
        [InputDouble(4, SymbolsConstants.MarkingSightLineMark, SymbolsConstants.MarkingSightLineMark, 5230)] //BaseSide
        public InputDouble sightLineMark;
        /// <summary>
        /// Ship Direction Mark
        /// </summary>
        [InputDouble(5, SymbolsConstants.MarkingShipDirectionMark, SymbolsConstants.MarkingShipDirectionMark, 5240)] //Apply
        public InputDouble shipDirectionMark;
        /// <summary>
        /// Label Mark
        /// </summary>
        [InputDouble(6, SymbolsConstants.MarkingLabelMark, SymbolsConstants.MarkingLabelMark, 5250)] //Apply
        public InputDouble labelMark;
        /// <summary>
        /// Frame Marks
        /// </summary>
        [InputDouble(7, SymbolsConstants.MarkingFrameMarks, SymbolsConstants.MarkingFrameMarks, 5260)] //Apply
        public InputDouble frameMark;
        /// <summary>
        /// Knuckle Marks
        /// </summary>
        [InputDouble(8, SymbolsConstants.MarkingKnuckleMarks, SymbolsConstants.MarkingKnuckleMarks, 5270)] //Apply
        public InputDouble knuckleMark;
        /// <summary>
        /// reference Marks
        /// </summary>
        [InputDouble(9, SymbolsConstants.MarkingReferenceMarks, SymbolsConstants.MarkingReferenceMarks, 5280)] //Apply
        public InputDouble referenceMark;
        /// <summary>
        /// Template Marks
        /// </summary>
        [InputDouble(10, SymbolsConstants.MarkingTemplateMarks, SymbolsConstants.MarkingTemplateMarks, 5290)] //Apply
        public InputDouble templateMark;
        /// <summary>
        /// Custom Marks
        /// </summary>
        [InputDouble(11, SymbolsConstants.MarkingCustomMarks, SymbolsConstants.MarkingCustomMarks, 52100)] //Apply
        public InputDouble customMark;
        
        #endregion

        //Define Aspect
        #region Define aspect
        /// <summary>
        /// The simple physical
        /// </summary>
        [Aspect(SymbolsConstants.MarkingAspect, SymbolsConstants.MarkingAspectDesc, AspectID.SimplePhysical)]
        public AspectDefinition simplePhysical;
        #endregion
    }
}