//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   CClipHoldClamp.cs
//    CTAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.CClipHoldClamp
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
    public class CClipHoldClamp : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "CTAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.CClipHoldClamp"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        //Following inputs are added by the wizard. Do not remove the catalog part input or change it's index.
        //Indices are sequential and need to be preserved.

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Width", "Clamp Width", 2)]
        public InputDouble Width;
        [InputDouble(3, "Depth", "Clamp Depth", 1)]
        public InputDouble Depth;
        [InputDouble(4, "Thickness", "Material Thickness", 0.2)]
        public InputDouble Thickness;
        [InputDouble(5, "RodDiameter", "Rod Diameter", 0.3)]
        public InputDouble RodDiameter;
        [InputDouble(6, "RodOffset", "Rod Offset", 0.5)]
        public InputDouble RodOffset;
        [InputDouble(7, "RodDepth", "Rod Depth", 1)]
        public InputDouble RodDepth;
        [InputDouble(8, "TrayThickness", "Tray Thickness", 0.2)]
        public InputDouble TrayThickness;
        [InputDouble(9, "TrayWidth", "Tray Width", 4)]
        public InputDouble TrayWidth;
        [InputDouble(10, "TrayDepth", "Tray Depth", 2)]
        public InputDouble TrayDepth;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("SimplePhysical", "SimplePhysical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Structure", "Rod Port")]
        [SymbolOutput("Route", "Tray Port")]
        [SymbolOutput("Clamps1", "Clamps1")]
        [SymbolOutput("Clamps2", "Clamps2")]
        [SymbolOutput("Clamps3", "Clamps3")]
        [SymbolOutput("Clamps4", "Clamps4")]
        [SymbolOutput("Clamps5", "Clamps5")]
        [SymbolOutput("Clamps6", "Clamps6")]
        [SymbolOutput("Clamps7", "Clamps7")]
        [SymbolOutput("Clamps8", "Clamps8")]
        [SymbolOutput("Clamps9", "Clamps9")]
        [SymbolOutput("Clamps10", "Clamps10")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construct Outputs"      
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;

                Double width = Width.Value;
                Double depth = Depth.Value;
                Double thickness = Thickness.Value;               
                Double rodOffset = RodOffset.Value;
                Double rodDepth = RodDepth.Value;
                Double trayThickness = TrayThickness.Value;
                Double trayW = TrayWidth.Value;
                Double trayH = TrayDepth.Value;

                trayThickness = thickness;      //TrayThickness not available, so set it to Thickness
                Double w = 0.5 * depth;         //W not available as occ attribute

                //=================================================
                //Construction of Physical Aspect 
                //=================================================               

                //'''''''''''  Generate top clamp '''''''''''''''''''''
                Collection<Position> pointCollectionTrace = new Collection<Position>();
                pointCollectionTrace.Add(new Position(trayW / 2.0 - w - trayThickness, trayH / 2.0 - thickness - trayThickness, depth / 2.0));
                pointCollectionTrace.Add(new Position(trayW / 2.0 - w - trayThickness, trayH / 2.0 - thickness - trayThickness, -depth / 2.0));
                pointCollectionTrace.Add(new Position(trayW / 2.0 - w - trayThickness - thickness, trayH / 2.0 - thickness - trayThickness, -depth / 2.0));
                pointCollectionTrace.Add(new Position(trayW / 2.0 - w - trayThickness - thickness, trayH / 2.0 - thickness - trayThickness, depth / 2.0));
                pointCollectionTrace.Add(new Position(trayW / 2.0 - w - trayThickness, trayH / 2.0 - thickness - trayThickness, depth / 2.0));

                Collection<Position> pointCollectionSweep = new Collection<Position>();
                pointCollectionSweep.Add(new Position(trayW / 2.0 - w - trayThickness, trayH / 2.0 - thickness - trayThickness, depth / 2.0));
                pointCollectionSweep.Add(new Position(trayW / 2.0 - w - trayThickness, trayH / 2.0 - thickness - trayThickness + thickness + trayThickness, depth / 2.0));
                pointCollectionSweep.Add(new Position(trayW / 2.0 - w - trayThickness + width - thickness, trayH / 2.0, depth / 2.0));

                Collection<Surface3d> topClampSurfaces = Surface3d.GetSweepSurfacesFromCurve(new LineString3d(pointCollectionSweep), new LineString3d(pointCollectionTrace), (SurfaceSweepOptions)1);               
                int i = 0;
                foreach (Surface3d item in topClampSurfaces)
                {
                    ++i;
                    m_Symbolic.Outputs.Add("Clamps" + i, item);
                }

                //'''''''''''  Generate bottom clamp '''''''''''''''''''''
                pointCollectionTrace = new Collection<Position>();
                pointCollectionTrace.Add(new Position(trayW / 2.0 - w, trayH / 2.0 - trayThickness, depth / 2.0));
                pointCollectionTrace.Add(new Position(trayW / 2.0 - w, trayH / 2.0 - trayThickness, -depth / 2.0));
                pointCollectionTrace.Add(new Position(trayW / 2.0 - w, trayH / 2.0 - trayThickness - thickness, -depth / 2.0));
                pointCollectionTrace.Add(new Position(trayW / 2.0 - w, trayH / 2.0 - trayThickness - thickness, depth / 2.0));
                pointCollectionTrace.Add(new Position(trayW / 2.0 - w, trayH / 2.0 - trayThickness, depth / 2.0));

                pointCollectionSweep = new Collection<Position>();
                pointCollectionSweep.Add(new Position(trayW / 2.0 - w, trayH / 2.0 - trayThickness, depth / 2.0));
                pointCollectionSweep.Add(new Position(trayW / 2.0 - w + width - thickness - trayThickness, trayH / 2.0 - trayThickness, depth / 2.0));

                Collection<Surface3d> bottomClampSurfaces = Surface3d.GetSweepSurfacesFromCurve(new LineString3d(pointCollectionSweep), new LineString3d(pointCollectionTrace), (SurfaceSweepOptions)1);
                
                foreach (Surface3d item in bottomClampSurfaces)
                {
                    ++i;
                    m_Symbolic.Outputs.Add("Clamps" + i, item);
                }

                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(trayW / 2.0 + width - w - thickness - trayThickness - rodOffset, trayH / 2.0 - rodDepth, 0), new Vector(0, 0, 1), new Vector(0, 1, 0));
                m_Symbolic.Outputs["Structure"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(0, 0, 1), new Vector(0, 1, 0));
                m_Symbolic.Outputs["Route"] = port2;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HgrCTAssemblySymbolsLocalizer.GetString(HgrCTAssemblySymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of CClipHoldClamp."));
                    return;
                }
            }
        }

        #endregion

    }
}


