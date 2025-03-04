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

namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.Tube
{
    /// <summary>
    /// Marking Symbol Definition for TemplateSet on Plate.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    public class MarkingDefinition : ManufacturingSymbolDefinition
    {
        //Define 7 inputs
        #region "Definition of Inputs"
        /// <summary>
        /// Base Control Line Mark 
        /// </summary>
        [InputDouble(1, SymbolsConstants.MarkingBaseCtlLineMark, SymbolsConstants.MarkingBaseCtlLineMark, 7200)] //Apply
        public InputDouble baseControlMark;
        /// <summary>
        /// Fitting Marks
        /// </summary>
        [InputDouble(2, SymbolsConstants.MarkingFittingMarks, SymbolsConstants.MarkingFittingMarks, 7200)] //Apply
        public InputDouble fittingMark;
        /// <summary>
        /// Frame Marks
        /// </summary>
        [InputDouble(3, SymbolsConstants.MarkingFrameMarks, SymbolsConstants.MarkingFrameMarks, 7200)] //Apply
        public InputDouble frameMark;
        /// <summary>
        /// Quarter Line Marks
        /// </summary>
        [InputDouble(4, SymbolsConstants.MarkingQuarterLineMarks, SymbolsConstants.MarkingQuarterLineMarks, 7200)] //Apply
        public InputDouble quarterLineMark;
        /// <summary>
        /// Seam Marks
        /// </summary>
        [InputDouble(5, SymbolsConstants.MarkingSeamMarks, SymbolsConstants.MarkingSeamMarks, 7200)] //Apply
        public InputDouble seamMark;
        /// <summary>
        /// Ship Direction Mark
        /// </summary>
        [InputDouble(6, SymbolsConstants.MarkingShipDirectionMark, SymbolsConstants.MarkingShipDirectionMark, 7200)] //Apply
        public InputDouble shipDirectionMark;       
        /// <summary>
        /// Custom Marks
        /// </summary>
        [InputDouble(7, SymbolsConstants.MarkingCustomMarks, SymbolsConstants.MarkingCustomMarks, 7201)] //Ignore
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