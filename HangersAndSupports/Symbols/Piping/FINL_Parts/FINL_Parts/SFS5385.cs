//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5385.cs
//   FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5385
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
    public class SFS5385 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5385"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Width", "Width", 0.999999)]
        public InputDouble m_dWidth;
        [InputDouble(3, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        [InputDouble(4, "HoleDia", "HoleDia", 0.999999)]
        public InputDouble m_dHoleDia;
        [InputDouble(5, "HoleCtoC", "HoleCtoC", 0.999999)]
        public InputDouble m_dHoleCtoC;
        [InputDouble(6, "HoleInset", "HoleInset", 0.999999)]
        public InputDouble m_dHoleInset;
        [InputDouble(7, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(8, "TakeOut", "TakeOut", 0.999999)]
        public InputDouble m_dTakeOut;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("VerPlate", "VerPlate")]
        [SymbolOutput("HorPlate", "HorPlate")]
        [SymbolOutput("Hole1", "Hole1")]
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
                 
               

                Double width = m_dWidth.Value;
                Double thickness = m_dThickness.Value;
                Double holeDiameter = m_dHoleDia.Value;
                Double holeCtoC = m_dHoleCtoC.Value;
                Double holeInset = m_dHoleInset.Value;
                Double length = m_dL.Value;
                Double takeOut = m_dTakeOut.Value;
                Matrix4X4 matrix = new Matrix4X4();

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Hole", new Position(0, holeInset, takeOut + thickness), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, width / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (length <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidLengthGTZero, "Length should be greater zero"));
                    return;
                }
                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidWidthGTZero, "Width should be greater zero"));
                    return;
                }
                if (holeDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidHoleDiameterGTZero, "HoleDia value should be greater zero"));
                    return;
                }
                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidThickness, "Thickness cannot be zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
               
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-length / 2, thickness, -thickness);
                Projection3d verPlate = (Projection3d)symbolGeometryHelper.CreateBox(null, length, width - thickness, thickness, 9);
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                verPlate.Transform(matrix);
                m_Symbolic.Outputs["VerPlate"] = verPlate;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-length / 2, 0, 0);
                Projection3d horPlate = (Projection3d)symbolGeometryHelper.CreateBox(null, length, width, thickness, 9);
                m_Symbolic.Outputs["HorPlate"] = horPlate;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-holeInset,0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d hole1 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holeDiameter / 2, thickness);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                hole1.Transform(matrix);
                m_Symbolic.Outputs["Hole1"] = hole1;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5385."));
                    return;
                }
            }
        }
        #endregion
    }
}
