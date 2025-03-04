//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   CHoldSideClamp.cs
//    CTAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.CHoldSideClamp
//   Author       :  Mahanth
//   Creation Date:   05-09-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   05-09-2013     Mahanth    CR-CP-222296- Converted CabletrayAssemblies to C# .Net 
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
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
    public class CHoldSideClamp : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "CTAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.CHoldSideClamp"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        //Following inputs are added by the wizard. Do not remove the catalog part input or change it's index.
        //Indices are sequential and need to be preserved.

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Thickness", "Material Thickness", 0.2)]
        public InputDouble Thickness;
        [InputDouble(3, "Depth", "Material Depth", 2)]
        public InputDouble Depth;
        [InputDouble(4, "RodDiameter", "Rod Diameter", 0.3)]
        public InputDouble RodDiameter;
        [InputDouble(5, "RodOffset", "Rod Offset", 0.5)]
        public InputDouble RodOffset;
        [InputDouble(6, "RodDepth", "Rod Depth", 1)]
        public InputDouble RodDepth;
        [InputDouble(7, "TrayBeamWidth", "Tray Beam Width", 1)]
        public InputDouble TrayBeamWidth;
        [InputDouble(8, "TrayDepth", "Cable Tray Depth", 4)]
        public InputDouble TrayDepth;
        [InputDouble(9, "TrayWidth", "Cable Tray Width", 10)]
        public InputDouble TrayWidth;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("SimplePhysical", "SimplePhysical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Structure", "Rod Port")]
        [SymbolOutput("Route", "Tray Port")]
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

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                double thickness = Thickness.Value;
                double depth = Depth.Value;
                double rodOffset = RodOffset.Value;
                double rodDepth = RodDepth.Value;
                double trayBeamWidth = TrayBeamWidth.Value;
                double trayDepth = TrayDepth.Value;
                double trayWidth = TrayWidth.Value;

                Collection<Position> pointCollectionSweep = new Collection<Position>();
                pointCollectionSweep.Add(new Position(trayWidth / 2.0 + depth, -trayDepth / 2.0, depth / 2.0));
                pointCollectionSweep.Add(new Position(trayWidth / 2.0 + depth, -trayDepth / 2.0, -depth / 2.0));
                pointCollectionSweep.Add(new Position(trayWidth / 2.0 + depth, -trayDepth / 2.0 - thickness, -depth / 2.0));
                pointCollectionSweep.Add(new Position(trayWidth / 2.0 + depth, -trayDepth / 2.0 - thickness, depth / 2.0));
                pointCollectionSweep.Add(new Position(trayWidth / 2.0 + depth, -trayDepth / 2.0, depth / 2.0));
               
                Collection<Position>  pointCollectionTrace = new Collection<Position>();
                pointCollectionTrace.Add(new Position(trayWidth / 2.0 + depth, -trayDepth / 2.0, depth / 2.0));
                pointCollectionTrace.Add(new Position(trayWidth / 2.0 + depth - depth - trayBeamWidth, -trayDepth / 2.0, depth / 2.0));
                pointCollectionTrace.Add(new Position(trayWidth / 2.0 + depth - depth - trayBeamWidth, trayDepth / 2.0, depth / 2.0));
                pointCollectionTrace.Add(new Position(trayWidth / 2.0 + depth - depth - trayBeamWidth + depth + trayBeamWidth, trayDepth / 2.0, depth / 2.0));
               
                Collection<Surface3d> surface = Surface3d.GetSweepSurfacesFromCurve(new LineString3d(pointCollectionTrace), new LineString3d(pointCollectionSweep), (SurfaceSweepOptions)1);
                int i = 1;
                foreach (Surface3d item in surface)
                {
                    m_Symbolic.Outputs.Add("Clamp" + i, item);
                    i++;
                }

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(trayWidth / 2 + depth - rodOffset, -trayDepth / 2.0 - thickness - rodDepth, 0), new Vector(0, 0, 1), new Vector(0, 1, 0));
                m_Symbolic.Outputs["Structure"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(0, 0, 1), new Vector(0, 1, 0));
                m_Symbolic.Outputs["Route"] = port2;

            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HgrCTAssemblySymbolsLocalizer.GetString(HgrCTAssemblySymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of CHoldSideClamp."));
                    return;
                }
            }
        }

        #endregion

    }
}
