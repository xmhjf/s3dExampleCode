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

using System.Collections;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using System.Collections.ObjectModel;


namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.Plate
{
    /// <summary>
    /// Process Selector Rule for TemplateSet on Plate.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [RuleInterface(SymbolsConstants.ProcessSelectorName, SymbolsConstants.ProcessSelectorUserName)]
    public class ProcessSelectorRule : PlateSelectorRule
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
                    PlatePartBase platePart = base.PlatePart;
                    if (platePart != null)
                    {
                        if (platePart is PlatePart)
                        {
                            PlatePart platePartObj = (PlatePart)platePart;
                            if (platePartObj.PartGeometryState == PartGeometryStateType.LightPart)
                                return new ReadOnlyCollection<string>(choices);
                        }

                        if (base.IsValidPart(SymbolsConstants.ProcessDefault) == true)
                            choices.Add(SymbolsConstants.ProcessDefault);
                        if (base.IsValidPart(SymbolsConstants.ProcessFrame) == true)
                            choices.Add(SymbolsConstants.ProcessFrame);
                        if (base.IsValidPart(SymbolsConstants.ProcessFrameEqualHeight) == true)
                            choices.Add(SymbolsConstants.ProcessFrameEqualHeight);

                        if (base.IsValidPart(SymbolsConstants.ProcessEven) == true)
                            choices.Add(SymbolsConstants.ProcessEven);
                        if (base.IsValidPart(SymbolsConstants.ProcessCenterLine) == true)
                            choices.Add(SymbolsConstants.ProcessCenterLine);
                        if (base.IsValidPart(SymbolsConstants.ProcessPerpendicular) == true)
                            choices.Add(SymbolsConstants.ProcessPerpendicular);

                        if (base.IsValidPart(SymbolsConstants.ProcessStemStern) == true)
                            choices.Add(SymbolsConstants.ProcessStemStern);
                        if (base.IsValidPart(SymbolsConstants.ProcessPerpendicularXY) == true)
                            choices.Add(SymbolsConstants.ProcessPerpendicularXY);
                        if (base.IsValidPart(SymbolsConstants.ProcessUserDefined) == true)
                            choices.Add(SymbolsConstants.ProcessUserDefined);

                        if (base.IsValidPart(SymbolsConstants.ProcessAftForward) == true)
                            choices.Add(SymbolsConstants.ProcessAftForward);
                        if (base.IsValidPart(SymbolsConstants.ProcessBox) == true)
                            choices.Add(SymbolsConstants.ProcessBox);
                        if (base.IsValidPart(SymbolsConstants.ProcessUserDefinedBox) == true)
                            choices.Add(SymbolsConstants.ProcessUserDefinedBox);
                        if (base.IsValidPart(SymbolsConstants.ProcessUserDefBoxEdges) == true)
                            choices.Add(SymbolsConstants.ProcessUserDefBoxEdges);                    

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
