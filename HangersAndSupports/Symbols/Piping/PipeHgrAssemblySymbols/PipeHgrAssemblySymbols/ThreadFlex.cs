//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   ThreadFlex.cs
//   PipeHgrAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.ThreadFlex
//   Author       :  Rajeswari
//   Creation Date:  03-OCT-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   03-OCT-2013  Rajeswari CR-CP-241285--Convert HgrAssemblySymbols to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    public class ThreadFlex : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.ThreadFlex"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "Length", "Length", 0.5)]
        public InputDouble Length;
        [InputDouble(3, "RodSizeDia", "RodSizeDia", 0.1)]
        public InputDouble RodSizeDia;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("FlexRod", "Flexible Rod")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        public AspectDefinition symbolicAspect;
        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)PartInput.Value;
                Double length = Length.Value;
                Double rodSizeDiameter = RodSizeDia.Value / 2;

                // create hgrports as part of the output
                Port port1 = new Port(OccurrenceConnection, part, "RodBottom", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                symbolicAspect.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "RodTop", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                symbolicAspect.Outputs["Port2"] = port2;

                if (rodSizeDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrInvalidRodSizediameterGTZ, "Diameter should be greater than zero."));
                    return;
                }
                if (length == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrInvalidLengthNEZ, "Length can not be equal to zero."));
                    return;
                }
                // construct the rod
                Circle3d circle = new Circle3d(new Position(0, 0, 0), new Vector(0.0, 0, 1.0), rodSizeDiameter);
                Projection3d flexRod = new Projection3d(circle, new Vector(0, 0, 1), length, true);
                symbolicAspect.Outputs["FlexRod"] = flexRod;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of ThreadFlex.cs"));
                    return;
                }
            }
        }

        #endregion

    }
}
