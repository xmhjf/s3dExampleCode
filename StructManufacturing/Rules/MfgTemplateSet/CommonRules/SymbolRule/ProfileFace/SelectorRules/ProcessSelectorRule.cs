//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   ProcessSelectorRule is a .NET selection rule for TemplateSet on Profile Face.
//                 
//
//      Author:  
//
//      History:
//      August 15th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System.Collections;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using System.Collections.ObjectModel;


namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.ProfileFace
{
    /// <summary>
    /// Process Selector Rule for TemplateSet on Profile Face.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [RuleInterface(SymbolsConstants.ProcessSelectorName, SymbolsConstants.ProcessSelectorUserName)]
    public class ProcessSelectorRule : ProfileSelectorRule
    {
        /// <summary>
        /// Gets the list of Predefined Set of TemplateSetProcess like a Default_TemplateProcessProfile, Frame_TemplateProcessProfile,...
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
                        if (base.IsValidPart(SymbolsConstants.ProcessFrame) == true)
                            choices.Add(SymbolsConstants.ProcessFrame);
                        if (base.IsValidPart(SymbolsConstants.ProcessPerpendicular) == true)
                            choices.Add(SymbolsConstants.ProcessPerpendicular);
                        if (base.IsValidPart(SymbolsConstants.ProcessUserDefined) == true)
                            choices.Add(SymbolsConstants.ProcessUserDefined);
                        if (base.IsValidPart(SymbolsConstants.ProcessEven) == true)
                            choices.Add(SymbolsConstants.ProcessEven); 
                        if (base.IsValidPart(SymbolsConstants.ProcessAftForward) == true)
                            choices.Add(SymbolsConstants.ProcessAftForward); 
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
