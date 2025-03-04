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


namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.ProfileEdge
{
    /// <summary>
    /// Marking Selector Rule for TemplateSet on Plate.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [RuleInterface(SymbolsConstants.MarkingSelectorName, SymbolsConstants.MarkingSelectorUserName)]
    public class MarkingSelectorRule : ProfileSelectorRule
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
                    ProfilePart profilePart = base.ProfilePart;
                    if (profilePart != null)
                    {
                        if (profilePart is StiffenerPart)
                        {
                            StiffenerPart profilePartObj = (StiffenerPart)profilePart;
                            if (profilePartObj.PartGeometryState == PartGeometryStateType.LightPart)
                                return new ReadOnlyCollection<string>(choices);
                        }

                        if (base.IsValidPart(SymbolsConstants.MarkingDefault) == true)
                            choices.Add(SymbolsConstants.MarkingDefault);                      

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
