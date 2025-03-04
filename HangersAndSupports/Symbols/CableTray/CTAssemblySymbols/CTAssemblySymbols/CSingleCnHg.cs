//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   CSingleCnHg.cs
//   CTAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.CSingleCnHg
//   Author       :  Mahanth
//   Creation Date:   05-09-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   05-09-2013     Mahanth    CR-CP-222296- Converted CabletrayAssemblies to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
    public class CSingleCnHg : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HgrCTAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.CSingleCnHg"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        //Following inputs are added by the wizard. Do not remove the catalog part input or change it's index.
        //Indices are sequential and need to be preserved.

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "A", "A", 2)]
        public InputDouble A;
        [InputDouble(3, "B", "B", 2)]
        public InputDouble B;
        [InputDouble(4, "C", "C", 4)]
        public InputDouble C;
        [InputDouble(5, "Depth", "Material Depth", 2)]
        public InputDouble Depth;
        [InputDouble(6, "Thickness", "Material Thickness", 0.2)]
        public InputDouble Thickness;
        [InputDouble(7, "RodDiameter", "Rod Diameter", 0.3)]
        public InputDouble RodDiameter;
        [InputDouble(8, "TrayDepth", "Cable Tray Height", 2)]
        public InputDouble TrayDepth;
        [InputDouble(9, "TrayWidth", "Cable Tray Width", 4)]
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
                double a = A.Value;
                double b = B.Value;
                double c = C.Value;
                double depth = Depth.Value;
                double thickness = Thickness.Value;           
                double trayDepth = TrayDepth.Value;
                double trayWidth = TrayWidth.Value;

                Collection<Position> pointCollectionSweep = new Collection<Position>();
                pointCollectionSweep.Add(new Position(trayWidth / 2.0, 0, depth / 2.0));
                pointCollectionSweep.Add(new Position(trayWidth / 2.0, 0, -depth / 2.0));
                pointCollectionSweep.Add(new Position(trayWidth / 2.0 + thickness, 0, -depth / 2.0));
                pointCollectionSweep.Add(new Position(trayWidth / 2.0 + thickness, 0, depth / 2.0));
                pointCollectionSweep.Add(new Position(trayWidth / 2.0, 0, depth / 2.0));
                LineString3d plane = new LineString3d(pointCollectionSweep);

                Collection<Position> pointCollectionTrace = new Collection<Position>();
                pointCollectionTrace.Add(new Position(trayWidth / 2.0, 0, depth / 2.0));
                pointCollectionTrace.Add(new Position(trayWidth / 2.0, 0 - trayDepth / 2.0, depth / 2.0));
                pointCollectionTrace.Add(new Position(-trayWidth / 2.0, 0 - trayDepth / 2.0, depth / 2.0));
                pointCollectionTrace.Add(new Position(-trayWidth / 2.0, 0 - trayDepth / 2.0 + c - 2.0 * thickness, depth / 2.0));
                pointCollectionTrace.Add(new Position(a / 2.0, 0 - trayDepth / 2.0 + c - 2.0 * thickness, depth / 2.0));
                pointCollectionTrace.Add(new Position(a / 2.0, 0 - trayDepth / 2.0 + c - 2.0 * thickness + b + thickness, depth / 2.0));
                pointCollectionTrace.Add(new Position(-a / 2.0, 0 - trayDepth / 2.0 + c - 2.0 * thickness + b + thickness, depth / 2.0));               
                LineString3d lineString = new LineString3d(pointCollectionTrace);

                Collection<Surface3d> surface = Surface3d.GetSweepSurfacesFromCurve( new LineString3d(pointCollectionTrace), new LineString3d(pointCollectionSweep), (SurfaceSweepOptions)1);
                int i = 1;
                foreach (Surface3d item in surface)
                {
                    m_Symbolic.Outputs.Add("Clamp" + i,item);
                    i++;
                }
                
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0,(2*( 0 - trayDepth / 2.0 + c - 2.0 * thickness)+b+thickness)/2.0, 0), new Vector(0, 0, 1), new Vector(0, 1, 0));
                m_Symbolic.Outputs["Structure"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(0, 0, 1), new Vector(0, 1, 0));
                m_Symbolic.Outputs["Route"] = port2;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HgrCTAssemblySymbolsLocalizer.GetString(HgrCTAssemblySymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of CSingleCnHg."));
                    return;
                }
            }
        }

        #endregion

    }
}
