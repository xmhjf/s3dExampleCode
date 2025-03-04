//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5387.cs
//   FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5387
//   Author       :  Hema
//   Creation Date:  18-03-2013
//   Description:    Converted FINL_Parts VB Project to C#.Net Project 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-03-2013      Hema    Converted FINL_Parts VB Project to C#.Net Project 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
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
    [CacheOption(CacheOptionType.Cached)]
    public class SFS5387 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5387"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "HoleDia", "HoleDia", 0.999999)]
        public InputDouble m_dHoleDia;
        [InputDouble(3, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(4, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        [InputDouble(5, "PlateL", "PlateL", 0.999999)]
        public InputDouble m_dPlateL;
        [InputDouble(6, "HoleInset", "HoleInset", 0.999999)]
        public InputDouble m_dHoleInset;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Plate", "Plate")]
        [SymbolOutput("Hole", "Hole")]
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
                 
               
               
                Double holeDiameter = m_dHoleDia.Value;
                Double length = m_dL.Value;
                Double thickness = m_dThickness.Value;
                Double plateL = m_dPlateL.Value;
                Double holeInset = m_dHoleInset.Value;
                Double height = plateL;
                Matrix4X4 matrix = new Matrix4X4();

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, -(height - holeInset)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (length <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidLengthGTZero, "Length should be greater than zero"));
                    return;
                }
                if (height <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidHeightGTZero, "Height should be greater than zero"));
                    return;
                }
                if (thickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidThicknessGTZero, "Thickness should be greater than zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
               
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0,0,0);
                Projection3d plate = (Projection3d)symbolGeometryHelper.CreateBox(null, thickness, length, height, 9);
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                matrix.Translate(new Vector(thickness / 2, length / 2, -height));
                plate.Transform(matrix);
                m_Symbolic.Outputs["Plate"] = plate;
                
                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(thickness / 2, 0, -(height - holeInset)).Subtract(new Position(-thickness / 2, 0, -(height - holeInset)));
                symbolGeometryHelper.ActivePosition = new Position(-thickness / 2, 0, -(height - holeInset));
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d hole = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holeDiameter / 2, normal.Length);
                m_Symbolic.Outputs["Hole"] = hole;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5387."));
                    return;
                }
            }
        }
        #endregion
    }
}
