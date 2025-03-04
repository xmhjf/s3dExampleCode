//----------------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   GenericMarkSelectorRule is a .NET selection rule for Manufacturing Generic Marks.
//                 
//
//      Author:  
//
//      History:
//      November 5th, 2015   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System.Collections;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using System.Collections.ObjectModel;


namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Generic Mark Selector Rule.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [RuleInterface(SymbolsConstants.GenericMarkSelectorName, SymbolsConstants.GenericMarkSelectorUserName)]
    public class GenericMarkSelectorRule : ManufacturingSelectorRule
    {
        /// <summary>
        /// Gets the list of Predefined Set of Generic Marks like Bilge Keel Mark, Block Mark.
        /// </summary>
        /// <value>
        /// The selections.
        /// </value>
        public override ReadOnlyCollection<string> Selections
        {
            get
            {
                Collection<string> choices = new Collection<string>();
                    
                choices.Add(SymbolsConstants.BilgeKeelMark);
                choices.Add(SymbolsConstants.BlockMark);

                return new ReadOnlyCollection<string>(choices);
            }
        }
    }
}
