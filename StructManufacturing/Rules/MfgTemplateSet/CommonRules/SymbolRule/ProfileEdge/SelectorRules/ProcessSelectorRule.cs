//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   ProcessSelectorRule is a .NET selection rule for TemplateSet on Plate.
//                 
//
//      Author:  
//
//      History:
//      August 5th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using System.Collections.ObjectModel;


namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.ProfileEdge
{
    /// <summary>
    /// Process Selector Rule for TemplateSet on Plate.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [RuleInterface(SymbolsConstants.ProcessSelectorName, SymbolsConstants.ProcessSelectorUserName)]
    public class ProcessSelectorRule : ProfileSelectorRule
    {
        /// <summary>
        /// Gets the list of Predefined Set of TemplateSetProcess like a Default_TemplateProcessPlate, Frame_TemplateProcessPlate,...
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

                        if (base.IsValidPart(SymbolsConstants.ProcessDefault) == true)
                            choices.Add(SymbolsConstants.ProcessDefault);
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
