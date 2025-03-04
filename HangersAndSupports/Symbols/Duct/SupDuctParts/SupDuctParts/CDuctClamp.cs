//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   CDuctClamp.cs
//    SupDuctParts,Ingr.SP3D.Content.Support.Symbols.CDuctClamp
//   Author       :  BS
//   Creation Date:  18-03-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-03-2012     BS      CR-CP-222295 Converted HgrSupDuctParts VB Project to C# .Net
//   6/11/2013     Vijaya   CR-CP-242533  Provide the ability to store GType outputs as a single blob for H&S .net symbols  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.NonCached)]
    [VariableOutputs]
    public class CDuctClamp : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "SupDuctParts,Ingr.SP3D.Content.Support.Symbols.CDuctClamp"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Length", "Length", 1.965)]
        public InputDouble m_Length;
        [InputDouble(3, "ClampLength", "ClampLength", 1.6)]
        public InputDouble m_ClampLength;
        [InputDouble(4, "LegLength", "LegLength", 0.15)]
        public InputDouble m_LegLength;
        [InputDouble(5, "Radius", "Radius", 0.58)]
        public InputDouble m_Radius;
        [InputDouble(6, "Thickness", "Thickness", 0.03)]
        public InputDouble m_Thickness;
        [InputDouble(7, "Width", "Width", 0.25)]
        public InputDouble m_Width;
        [InputDouble(8, "BoltDia", "BoltDia", 0.1)]
        public InputDouble m_BoltDia;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("DuctClamp", "DuctClamp")]
        [SymbolOutput("LeftBolt", "LeftBolt")]
        [SymbolOutput("RightBolt", "RightBolt")]
        [SymbolOutput("HgrPort_1", "HgrPort_1")]
        [SymbolOutput("HgrPort_2", "HgrPort_2")]
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

                double length = m_Length.Value;
                double clampLength = m_ClampLength.Value;
                double legLength = m_LegLength.Value;
                double radius = m_Radius.Value;
                double thickness = m_Thickness.Value;
                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SupDuctPartsLocalizer.GetString(SupDuctPartSymbolResourceIDs.ErrInvalidThicknessNZero, "Thickness value should not be 0"));
                    return;
                }
                double width = m_Width.Value;
                if (width == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError,SupDuctPartsLocalizer.GetString(SupDuctPartSymbolResourceIDs.ErrInvalidWidthNZero,"Width value should not be 0"));
                    return;
                }
                double boltDiameter = m_BoltDia.Value;
                double boltRadius = boltDiameter / 2;
                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                //Construct in the x/y plane with origin at the center of the Clamp.
                //The z-axis will serve as the "Width" direction of the Clamp.    
                //========================================
                // Clamp and Projection
                //=======================================
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Collection<Position> positionCollection = new Collection<Position>();
                positionCollection.Add(new Position(0,-radius - thickness,0));
                positionCollection.Add(new Position(0, -radius - thickness, length - thickness));
                positionCollection.Add(new Position(0, -clampLength * 0.5 - legLength, length - thickness));
                positionCollection.Add(new Position(0, -clampLength * 0.5 - legLength, length));
                positionCollection.Add(new Position(0, -radius, length));
                positionCollection.Add(new Position(0, -radius, 0));

                curveCollection.Add(new LineString3d(positionCollection));
                curveCollection.Add(new Arc3d(new Position(0, radius, 0),new Position( 0, 0, -radius),new Position( 0, -radius, 0)));

                positionCollection = new Collection<Position>();
                positionCollection.Add(new Position(0, radius, 0));
                positionCollection.Add(new Position(0, radius, length ));
                positionCollection.Add(new Position(0, (clampLength * 0.5 + legLength), length));
                positionCollection.Add(new Position(0, (clampLength * 0.5 + legLength), length - thickness));
                positionCollection.Add(new Position(0, (radius + thickness), length - thickness));
                positionCollection.Add(new Position(0, (radius + thickness), 0));

                curveCollection.Add(new LineString3d(positionCollection));
                curveCollection.Add(new Arc3d(new Position(0, radius + thickness, 0), new Position(0, 0, -radius - thickness), new Position(0, -radius - thickness, 0)));

                Projection3d ductClamp = new Projection3d(new ComplexString3d(curveCollection),new Vector(1,0,0), width, true);
                m_Symbolic.Outputs["DuctClamp"] = ductClamp;
                // ========================================
                // LeftBolt and Projection
                //========================================
                Circle3d leftCircle = new Circle3d(new Position(width * 0.5 + boltRadius, -clampLength * 0.5, length + thickness), new Position(width * 0.5, -clampLength * 0.5 + boltRadius, length + thickness), new Position(width * 0.5 - boltRadius, -clampLength * 0.5 - boltRadius, length + thickness));
                Projection3d leftBolt = new Projection3d(leftCircle, new Vector(0, 0, -1), 3 * thickness, true);
                m_Symbolic.Outputs["LeftBolt"] = leftBolt;
                //========================================
                // RightBolt and Projection
                //========================================
                Circle3d rightCircle = new Circle3d(new Position(width * 0.5 + boltRadius, clampLength * 0.5, length + thickness), new Position(width * 0.5, clampLength * 0.5 + boltRadius, length + thickness), new Position(width * 0.5 - boltRadius, clampLength * 0.5 - boltRadius, length + thickness));
                Projection3d rightBolt = new Projection3d(rightCircle, new Vector(0, 0, -1), 3 * thickness, true);
                m_Symbolic.Outputs["RightBolt"] = rightBolt;

                //=============================================================================
                // create hgrports as part of the output in the case of symbolic representation
                //=============================================================================                
                //First Port Located at the Center of clamp along the Z-Direction
                Port port1 = new Port(OccurrenceConnection, part, "HgrPort_1", new Position(width/2, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["HgrPort_1"] = port1;
                //Second Port Located at the Center of Clamp along Y/Z Direction
                Port port2 = new Port(OccurrenceConnection, part, "HgrPort_2", new Position(width / 2, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["HgrPort_2"] = port2;
               
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SupDuctPartsLocalizer.GetString(SupDuctPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Duct Clamp"));
                }
            }
        }
        #endregion
    }
}
