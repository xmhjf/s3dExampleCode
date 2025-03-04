//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   MarkingSelectorRule is a .NET selection rule for TemplateSet on Plate.
//                 
//
//      Author:  
//
//      History:
//      August 5th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System.Collections;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using System.Collections.ObjectModel;


namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.Plate
{
    /// <summary>
    /// Marking Selector Rule for TemplateSet on Plate.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [RuleInterface(SymbolsConstants.MarkingSelectorName, SymbolsConstants.MarkingSelectorUserName)]
    public class MarkingSelectorRule : PlateSelectorRule
    {
        /// <summary>
        /// Gets the list of Predefined Set of TemplateSetMarking like a Default_TemplateMarkingPlate, Box_TemplateMarkingPlate,...
        /// </summary>
        /// <value>
        /// The selections.
        /// </value>
        /// <exception cref="System.NotImplementedException"></exception>
        public override ReadOnlyCollection<string> Selections
        {
            get
            {
                Collection<string> choices = new Collection<string>();

                try
                {

                    PlatePartBase platePart = base.PlatePart;
                    if (platePart != null)
                    {  
                        if (platePart is PlatePart)
                        {
                            PlatePart platePartObj = (PlatePart)platePart;
                            if (platePartObj.PartGeometryState == PartGeometryStateType.LightPart)
                                return new ReadOnlyCollection<string>(choices);
                        }

                        if (base.IsValidPart(SymbolsConstants.MarkingDefault) == true)
                            choices.Add(SymbolsConstants.MarkingDefault);
                        if (base.IsValidPart(SymbolsConstants.MarkingBox) == true)
                            choices.Add(SymbolsConstants.MarkingBox);                           

                    }

                }
                catch
                {
                    //To Do
                }
                return new ReadOnlyCollection<string>(choices);
            }
        }
    }
}
