//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   SFS5369.cs
//    FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5369
//   Author       :   Vijaya
//   Creation Date:  18/3/2013
//   Description: CR-CP-222272 Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//    18/3/2013      Vijaya   CR-CP-222272 Initial Creation
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
    public class SFS5369 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5369"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(4, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(5, "D", "D", 0.999999)]
        public InputDouble m_dD;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BEND", "BEND")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("RIGHT", "RIGHT")]
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

                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double E = m_dE.Value;
                Double H = m_dH.Value;
                Double D = m_dD.Value;

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, pipeDiameter / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }
                if (H == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidHNZero, "H value cannot be zero "));
                    return;
                }                
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
               
                Revolution3d bend = new Revolution3d((new Circle3d(new Position(E / 2, 0, 0), new Vector(0, 0, 1), D / 2)), new Vector(0, -1, 0), new Position(0, 0, 0), Math.PI * 180 / 180, true);
                rotateMatrix.SetIdentity();
                rotateMatrix.Rotate(Math.PI, new Vector(1, 0, 0));
                rotateMatrix.Translate(new Vector(0, 0, 0));
                bend.Transform(rotateMatrix);
                m_Symbolic.Outputs["BEND"] = bend;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(E / 2, 0, 0);
                Vector normal1 = new Vector(0, 0, 1);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());                
                Projection3d right = symbolGeometryHelper.CreateCylinder(null, D / 2.0, H);
                m_Symbolic.Outputs["RIGHT"] = right;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-E / 2, 0, 0);
                Vector normal2 = new Vector(0, 0, 1);
                symbolGeometryHelper.SetOrientation(normal2, normal2.GetOrthogonalVector());
                Projection3d left = symbolGeometryHelper.CreateCylinder(null, D / 2.0, H);
                m_Symbolic.Outputs["LEFT"] = left;

            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5369."));
                    return;
                }

            }
        }
        #endregion

    }

}
