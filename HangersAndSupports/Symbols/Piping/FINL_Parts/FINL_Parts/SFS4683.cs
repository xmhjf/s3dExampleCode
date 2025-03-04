//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   SFS4683.cs
//    FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS4683
//   Author       :  Vijaya
//   Creation Date:  18/3/2013
//   Description: CR-CP-222272 Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18/3/2013      Vijaya   CR-CP-222272 Initial Creation
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
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class SFS4683 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS4683"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_dROD_DIA;
        [InputDouble(3, "Width", "Width", 0.999999)]
        public InputDouble m_dWidth;
        [InputDouble(4, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        [InputDouble(5, "HoleDia", "HoleDia", 0.999999)]
        public InputDouble m_dHoleDia;
        [InputDouble(6, "TakeOut", "TakeOut", 0.999999)]
        public InputDouble m_dTakeOut;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Washer", "Washer")]
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
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 rotateMatrix = new Matrix4X4();
                Double rodDiameter = m_dROD_DIA.Value;
                Double width = m_dWidth.Value;
                Double thickness = m_dThickness.Value;
                Double holeDiameter = m_dHoleDia.Value;
                Double takeOut = m_dTakeOut.Value;

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "BotStructure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "TopStructure", new Position(0, 0, thickness), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, thickness / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;


                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidWidthGTZero, "Width should be greater than zero"));
                    return;
                }
                if (thickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidThickness, "Thickness should be greater than zero"));
                    return;
                }
                if (holeDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidHoleDiameterGTZero, "Hole Dia value should be greater than zero"));
                    return;
                }


                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                
                symbolGeometryHelper.ActivePosition = new Position(-width / 2, -width / 2, 0);
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d washer = (Projection3d)symbolGeometryHelper.CreateBox(null, width, width, thickness, 9);
                washer.Transform(rotateMatrix);
                m_Symbolic.Outputs["Washer"] = washer;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                Vector normal = new Vector(0, 0, 1);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(0, 0, 0));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d hole = symbolGeometryHelper.CreateCylinder(null, holeDiameter / 2.0, thickness);
                hole.Transform(rotateMatrix);
                m_Symbolic.Outputs["Hole"] = hole;

            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS4683."));
                    return;
                }
            }
        }
        #endregion

    }

}
