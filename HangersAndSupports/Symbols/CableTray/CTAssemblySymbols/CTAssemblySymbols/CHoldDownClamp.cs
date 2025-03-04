//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   CHoldDownClamp.cs
//    CTAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.CHoldDownClamp
//   Author       :  Vijay
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013     Vijay    CR-CP-222296- Converted CabletrayAssemblies to C# .Net       
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle;
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

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class CHoldDownClamp : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "CTAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.CHoldDownClamp"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        //Following inputs are added by the wizard. Do not remove the catalog part input or change it's index.
        //Indices are sequential and need to be preserved.

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "A", "A", 1)]
        public InputDouble A;
        [InputDouble(3, "B", "B", 1)]
        public InputDouble B;
        [InputDouble(4, "Depth", "Material Depth", 1)]
        public InputDouble Depth;
        [InputDouble(5, "Thickness", "Material Thickness", 0.2)]
        public InputDouble Thickness;
        [InputDouble(6, "BoltDiameter", "Bolt Diameter", 0.3)]
        public InputDouble BoltDiameter;
        [InputDouble(7, "BoltOffset", "Bolt Length", 1)]
        public InputDouble BoltOffset;
        [InputDouble(8, "BoltLength", "Bolt Offset", 0.5)]
        public InputDouble BoltLength;
        [InputDouble(9, "TrayDepth", "Cable Tray Depth", 6)]
        public InputDouble TrayDepth;
        [InputDouble(10, "TrayWidth", "Cable Tray Width", 2)]
        public InputDouble TrayWidth;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("SimplePhysical", "SimplePhysical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Structure", "Rod Port")]
        [SymbolOutput("Route", "Tray Port")]
        [SymbolOutput("Bolt", "Bolt")]
        [SymbolOutput("Clamp1", "Clamp1")]
        [SymbolOutput("Clamp2", "Clamp2")]
        [SymbolOutput("Clamp3", "Clamp3")]
        [SymbolOutput("Clamp4", "Clamp4")]
        [SymbolOutput("Clamp5", "Clamp5")]
        [SymbolOutput("Clamp6", "Clamp6")]
        [SymbolOutput("Clamp7", "Clamp7")]
        [SymbolOutput("Clamp8", "Clamp8")]
        [SymbolOutput("Clamp9", "Clamp9")]
        [SymbolOutput("Clamp10", "Clamp10")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construct Outputs"

        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;

                Double a = A.Value;
                Double b = B.Value;
                Double depth = Depth.Value;
                Double thickness = Thickness.Value;
                Double boltDiameter = BoltDiameter.Value;
                Double boltLength = BoltOffset.Value;
                Double trayW = TrayWidth.Value;
                Double trayH = TrayDepth.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                if (boltLength == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HgrCTAssemblySymbolsLocalizer.GetString(HgrCTAssemblySymbolsResourceIDs.ErrBoltLengthNEZ, "Bolt length value cannot be zero."));
                    return;
                }
                //The origin of local coordinate is located at the origin of TrayPort

                // First, generate trace curve

                Collection<Position> pointCollectionTrace = new Collection<Position>();
                pointCollectionTrace.Add(new Position(trayW / 2.0 + b, -trayH / 2.0, depth / 2.0));
                pointCollectionTrace.Add(new Position(trayW / 2.0, -trayH / 2.0, depth / 2.0));
                pointCollectionTrace.Add(new Position(trayW / 2.0, trayH / 2.0, depth / 2.0));
                pointCollectionTrace.Add(new Position(trayW / 2.0 - a + thickness, trayH / 2.0, depth / 2.0));

                //cross section
                Collection<Position> pointCollectionSweep = new Collection<Position>();
                pointCollectionSweep.Add(new Position(trayW / 2.0 + b, -trayH / 2.0, depth / 2.0));
                pointCollectionSweep.Add(new Position(trayW / 2.0 + b, -trayH / 2.0, -depth / 2.0));
                pointCollectionSweep.Add(new Position(trayW / 2.0 + b, -trayH / 2.0 + thickness, -depth / 2.0));
                pointCollectionSweep.Add(new Position(trayW / 2.0 + b, -trayH / 2.0 + thickness, depth / 2.0));
                pointCollectionSweep.Add(new Position(trayW / 2.0 + b, -trayH / 2.0, depth / 2.0));

                Collection<Surface3d> zigPlateSurfaces = Surface3d.GetSweepSurfacesFromCurve(new LineString3d(pointCollectionTrace), new LineString3d(pointCollectionSweep), (SurfaceSweepOptions)1);

                int i = 0;
                //Add Base to output collection
                foreach (Surface3d item in zigPlateSurfaces)
                {
                    ++i;
                    m_Symbolic.Outputs.Add("Clamp" + i, item);
                }
                
                Projection3d projection = new Projection3d(new Circle3d(new Position((trayW / 2.0 + trayW / 2.0 + b) / 2.0, -trayH / 2.0 + thickness + boltDiameter, 0), new Vector(0, -1, 0), boltDiameter / 2.0), new Vector(0, -1, 0), boltLength, true);
                m_Symbolic.Outputs["Bolt"] = projection;


                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(trayW / 2.0 + b, -trayH / 2.0, 0), new Vector(0, 0, 1), new Vector(0, 1, 0));
                m_Symbolic.Outputs["Structure"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(0, 0, 1), new Vector(0, 1, 0));
                m_Symbolic.Outputs["Route"] = port2;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HgrCTAssemblySymbolsLocalizer.GetString(HgrCTAssemblySymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of CHoldDownClamp."));
                    return;
                }
            }
        }

        #endregion

    }
}


