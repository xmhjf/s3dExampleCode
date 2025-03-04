//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   BottomLug.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.BottomLug
//   Author       :  Sridhar
//   Creation Date:  29/10/2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   29/10/2012    Sridhar   CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

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
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class BottomLug : CustomSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.BottomLug"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Thickness", "Thickness", 0.1)]
        public InputDouble m_Thickness;
        [InputDouble(3, "PipeRadius", "PipeRadius", 0)]
        public InputDouble m_PipeRadius;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Represntation of the 3d flexible", AspectID.SimplePhysical)]
        [SymbolOutput("Lug", "Lug")]
        [SymbolOutput("FlatBar1", "Flat Bar1")]
        [SymbolOutput("FlatBar2", "Flat Bar2")]
        [SymbolOutput("TopFlatBar", "Top Flat Bar")]
        [SymbolOutput("Pin", "Pin")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Weldline", "Weld Line")]
        public AspectDefinition m_Symbolic;

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
                Part part = (Part)m_PartInput.Value;

                Double thickness = m_Thickness.Value;
                Double pipeRadius = m_PipeRadius.Value;
                //Assumptions
                Double height = 1.5 * pipeRadius;
                Double width = 0.9 * pipeRadius;
                Double radius = 0.35 * width;
                
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, thickness / 2, -pipeRadius), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Rod", new Position(0, thickness / 2, height + 1.5 * radius + thickness), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidThickness,"Thickness cannot be zero"));
                    return;
                }

                Line3d weldLine = new Line3d(new Position(-width / 2, thickness / 2, 0), new Position(width / 2, thickness / 2, 0));
                m_Symbolic.Outputs["Weldline"] = weldLine;
                Collection<Position> collectionPoints = new Collection<Position>();

                collectionPoints.Add(new Position(-radius, 0, height));
                collectionPoints.Add(new Position(-width / 2, 0, 0));
                collectionPoints.Add(new Position(width / 2, 0, 0));
                collectionPoints.Add(new Position(radius, 0, height));

                Collection<ICurve> collectionCurves = new Collection<ICurve>();
                collectionCurves.Add(new LineString3d(collectionPoints));
                collectionCurves.Add(new Arc3d(new Position(radius, 0, height), new Position(0, 0, height + radius), new Position(-radius, 0, height)));
                Projection3d lug = new Projection3d(new ComplexString3d(collectionCurves), new Vector(0, 1, 0), thickness, true);
                m_Symbolic.Outputs["Lug"] = lug;

                collectionPoints = new Collection<Position>();
                collectionPoints.Add(new Position(-0.5 * radius, 0, height - radius));
                collectionPoints.Add(new Position(0.5 * radius, 0, height - radius));
                collectionPoints.Add(new Position(0.5 * radius, 0, height + 1.5 * radius));
                collectionPoints.Add(new Position(-0.5 * radius, 0, height + 1.5 * radius));
                collectionPoints.Add(new Position(-0.5 * radius, 0, height - radius));

                Projection3d flatBar, pin;
                flatBar = new Projection3d(new LineString3d(collectionPoints), new Vector(0, -1, 0), thickness, true);
                m_Symbolic.Outputs["FlatBar1"] = flatBar;

                collectionPoints = new Collection<Position>();
                collectionPoints.Add(new Position(-0.5 * radius, thickness, height - radius));
                collectionPoints.Add(new Position(0.5 * radius, thickness, height - radius));
                collectionPoints.Add(new Position(0.5 * radius, thickness, height + 1.5 * radius));
                collectionPoints.Add(new Position(-0.5 * radius, thickness, height + 1.5 * radius));
                collectionPoints.Add(new Position(-0.5 * radius, thickness, height - radius));
                flatBar = new Projection3d(new LineString3d(collectionPoints), new Vector(0, 1, 0), thickness, true);
                m_Symbolic.Outputs["FlatBar2"] = flatBar;

                collectionPoints = new Collection<Position>();
                collectionPoints.Add(new Position(-0.5 * radius, 2 * thickness, height + 1.5 * radius));
                collectionPoints.Add(new Position(0.5 * radius, 2 * thickness, height + 1.5 * radius));
                collectionPoints.Add(new Position(0.5 * radius, -thickness, height + 1.5 * radius));
                collectionPoints.Add(new Position(-0.5 * radius, -thickness, height + 1.5 * radius));
                collectionPoints.Add(new Position(-0.5 * radius, 2 * thickness, height + 1.5 * radius));
                flatBar = new Projection3d(new LineString3d(collectionPoints), new Vector(0, 0, 1), thickness, true);
                m_Symbolic.Outputs["TopFlatBar"] = flatBar;

                pin = new Projection3d(new Circle3d(new Position(0, -2 * thickness, height), new Vector(0, 1, 0), radius / 3), new Vector(0, 1, 0), 5 * thickness, true);
                m_Symbolic.Outputs["Pin"] = pin;
            }
            catch 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of BottomLug.cs"));
                    return;
                }
            }
        }
        #endregion
    }
}