/***************************************************************************************************************************************/
// Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  RootTeeWeldSelectorRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMPhysConnRules.dll
//  Original Class Name: ‘RootTeeWeldSel’ in VB content
//
//Abstract
//	RootTeeWeldSelectorRule is a .NET selector rule which selects the list of available items in the context of the PhysicalConnections. 
//  This class subclasses from PhysicalConnectionSelectorRule.
/*****************************************************************************************************************************************/

using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Selector for PhysicalConnections, which selects the list of available items in the context of the ChamferTeeWelds.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get selector question. 
    [RuleInterface(DetailingCustomAssembliesConstants.IASMPhysConnRules_RootTeeWeldSel, DetailingCustomAssembliesConstants.IASMPhysConnRules_RootTeeWeldSel)]
    [DefaultLocalizer(ConnectionsResourceIdentifiers.ConnectionsResource, ConnectionsResourceIdentifiers.ConnectionsAssembly)]
    public class RootTeeWeldSelectorRule : PhysicalConnectionSelectorRule
    {
        //=================================================================================================================
        //DefinitionName/ProgID of this symbol is "Connections,Ingr.SP3D.Content.Structure.RootTeeWeldSelectorRule"
        //=================================================================================================================
        public override ReadOnlyCollection<string> Selections
        {
            get
            {
                Collection<string> choices = new Collection<string>();

                //Validating the inputs required to create the slot.
                //This is a virtual method and can be overriden to add specific validation, if specific input is missing.
                ValidateInputs();

                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created while validating the inputs
                    return new ReadOnlyCollection<string>(choices);
                }

                choices.Add(DetailingCustomAssembliesConstants.RootTeeWeldAutomaticSplitter);
                return new ReadOnlyCollection<string>(choices);
            }

        }
    }
}
