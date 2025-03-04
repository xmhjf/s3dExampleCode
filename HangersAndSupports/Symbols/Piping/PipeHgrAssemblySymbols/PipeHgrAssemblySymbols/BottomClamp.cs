//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   BottomClamp.cs
//   PipeHgrAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.BottomClamp
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
using System.Collections.ObjectModel;

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
    public class BottomClamp : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.BottomClamp"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "PipeRadius", "PipeRadius", 0.5)]
        public InputDouble PipeRadius;
        [InputDouble(3, "Depth", "Thickness", 0.15)]
        public InputDouble Depth;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("LeftHalfClamp", "Left Half of the clamp")]
        [SymbolOutput("RightHalfClamp", "Right Half of the clamp")]
        [SymbolOutput("Bolt", "Bolt")]
        [SymbolOutput("Pin1", "Pin1")]
        [SymbolOutput("Pin2", "Pin2")]
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
                Double pipeRadius = PipeRadius.Value;
                Double width = Depth.Value;
                Double thickness = pipeRadius / 8;

                // Create hgrports as part of the output
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(width / 2, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                symbolicAspect.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Rod", new Position(width / 2, 0, 2 * pipeRadius), new Vector(1, 0, 0), new Vector(0, 0, 1));
                symbolicAspect.Outputs["Port2"] = port2;

                if (pipeRadius <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrInvalidPipeRadiusGTZ, "Pipe Radius should be greater than zero."));
                    return;
                }
                if (width == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrInvalidWidthNEZ, "Depth can not be equal to zero."));
                    return;
                }

                // create the arcs
                double yOffset = thickness / 2;
                Collection<ICurve> curveCollection = new Collection<ICurve>();

                Arc3d arc = new Arc3d(new Position(0, -yOffset, -pipeRadius), new Position(0, -pipeRadius, 0), new Position(0, -yOffset, pipeRadius));
                curveCollection.Add(arc);
                arc = new Arc3d(new Position(0, -(yOffset + thickness), pipeRadius + thickness), new Position(0, -(pipeRadius + thickness), 0), new Position(0, -(yOffset + thickness), -(pipeRadius + thickness)));
                curveCollection.Add(arc);
                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0, -(yOffset + thickness), -(pipeRadius + thickness)));
                pointCollection.Add(new Position(0, -(yOffset + thickness), -(1.5 * pipeRadius + thickness)));
                pointCollection.Add(new Position(0, -yOffset, -(1.5 * pipeRadius + thickness)));
                pointCollection.Add(new Position(0, -yOffset, -pipeRadius));
                curveCollection.Add(new LineString3d(pointCollection));
                pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0, -yOffset, pipeRadius));
                pointCollection.Add(new Position(0, -yOffset, (1.5 * pipeRadius + thickness)));
                pointCollection.Add(new Position(0, -(yOffset + thickness), (1.5 * pipeRadius + thickness)));
                pointCollection.Add(new Position(0, -(yOffset + thickness), (pipeRadius + thickness)));
                curveCollection.Add(new LineString3d(pointCollection));
                // construct the left half clamp
                Projection3d leftClamp = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), width, true);
                symbolicAspect.Outputs["LeftHalfClamp"] = leftClamp;

                curveCollection = new Collection<ICurve>();
                arc = new Arc3d(new Position(0, yOffset, -pipeRadius), new Position(0, pipeRadius, 0), new Position(0, yOffset, pipeRadius));
                curveCollection.Add(arc);
                arc = new Arc3d(new Position(0, (yOffset + thickness), pipeRadius + thickness), new Position(0, (pipeRadius + thickness), 0), new Position(0, (yOffset + thickness), -(pipeRadius + thickness)));
                curveCollection.Add(arc);
                pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0, (yOffset + thickness), -(pipeRadius + thickness)));
                pointCollection.Add(new Position(0, (yOffset + thickness), -(1.5 * pipeRadius + thickness)));
                pointCollection.Add(new Position(0, yOffset, -(1.5 * pipeRadius + thickness)));
                pointCollection.Add(new Position(0, yOffset, -pipeRadius));
                curveCollection.Add(new LineString3d(pointCollection));
                pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0, yOffset, pipeRadius));
                pointCollection.Add(new Position(0, yOffset, (1.5 * pipeRadius + thickness)));
                pointCollection.Add(new Position(0, (yOffset + thickness), (1.5 * pipeRadius + thickness)));
                pointCollection.Add(new Position(0, (yOffset + thickness), (pipeRadius + thickness)));
                curveCollection.Add(new LineString3d(pointCollection));
                // construct the right half clamp
                Projection3d rightClamp = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), width, true);
                symbolicAspect.Outputs["RightHalfClamp"] = rightClamp;

                // construct the bolt
                pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0.05 * width, -yOffset, pipeRadius));
                pointCollection.Add(new Position(0.05 * width, yOffset, pipeRadius));
                pointCollection.Add(new Position(0.05 * width, yOffset, 1.75 * pipeRadius));
                pointCollection.Add(new Position(0.05 * width, 2 * thickness, 1.75 * pipeRadius));
                pointCollection.Add(new Position(0.05 * width, 2 * thickness, 2 * pipeRadius));
                pointCollection.Add(new Position(0.05 * width, -2 * thickness, 2 * pipeRadius));
                pointCollection.Add(new Position(0.05 * width, -2 * thickness, 1.75 * pipeRadius));
                pointCollection.Add(new Position(0.05 * width, -yOffset, 1.75 * pipeRadius));
                pointCollection.Add(new Position(0.05 * width, -yOffset, pipeRadius));
                Projection3d bolt = new Projection3d(new LineString3d(pointCollection), new Vector(1, 0, 0), 0.9 * width, true);
                symbolicAspect.Outputs["Bolt"] = bolt;

                // create the pins
                Circle3d circle = new Circle3d(new Position(width / 2, -2 * thickness, thickness + 1.25 * pipeRadius), new Vector(0, 1, 0), pipeRadius / 8);
                Projection3d pin1 = new Projection3d(circle, new Vector(0, 1, 0), 4 * thickness, true);
                symbolicAspect.Outputs["Pin1"] = pin1;

                circle = new Circle3d(new Position(width / 2, -2 * thickness, -(thickness + 1.25 * pipeRadius)), new Vector(0, 1, 0), pipeRadius / 8);
                Projection3d pin2 = new Projection3d(circle, new Vector(0, 1, 0), 4 * thickness, true);
                symbolicAspect.Outputs["Pin2"] = pin2;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of BottomClamp.cs"));
                    return;
                }
            }
        }

        #endregion

    }
}
