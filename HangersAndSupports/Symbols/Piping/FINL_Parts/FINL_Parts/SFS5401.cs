//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   SFS5401.cs
//    FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5401
//   Author       :   Rajeswari
//   Creation Date:  18-March-2013
//   Description:    CR-CP-222272-Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  18-March-2013 Rajeswari CR-CP-222272-Initial Creation 
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
    public class SFS5401 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5401"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(3, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(4, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(5, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(6, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(7, "BoltL", "BoltL", 0.999999)]
        public InputDouble m_dBoltL;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP1", "TOP1")]
        [SymbolOutput("TOP2", "TOP2")]
        [SymbolOutput("BOT1", "BOT1")]
        [SymbolOutput("BOT2", "BOT2")]
        [SymbolOutput("TOP_BOLT", "TOP_BOLT")]
        [SymbolOutput("BOT_BOLT", "BOT_BOLT")]
        [SymbolOutput("BODY", "BODY")]
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

                Double A = m_dA.Value;
                Double L = m_dL.Value;
                Double T = m_dT.Value;
                Double B = m_dB.Value;
                Double pipeDiameter = m_dD.Value;
                Double boltL = m_dBoltL.Value;
                Double F = 0.022;
                Double M = 0.024;
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, A / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidTGTZero, "T value should be greater than zero"));
                    return;
                }
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidBGTZero, "B value should be greater than zero"));
                    return;
                }
                if (M <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidMGTZero, "M value should be greater than zero"));
                    return;
                }
                if ((L - pipeDiameter) <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidpipeDiaandLNZero, "L - PipeDiameter value cannot be zero"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                symbolGeometryHelper.ActivePosition = new Position(-(F / 2 + T), -B / 2, pipeDiameter / 2);
                Projection3d top1 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, B, (L / 2 - pipeDiameter / 2), 9);
                m_Symbolic.Outputs["TOP1"] = top1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(F / 2, -B / 2, pipeDiameter / 2);
                Projection3d top2 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, B, (L / 2 - pipeDiameter / 2), 9);
                m_Symbolic.Outputs["TOP2"] = top2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(F / 2 + T), -B / 2, -L / 2);
                Projection3d bot1 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, B, (L / 2 - pipeDiameter / 2), 9);
                m_Symbolic.Outputs["BOT1"] = bot1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(F / 2, -B / 2, -L / 2);
                Projection3d bot2 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, B, (L / 2 - pipeDiameter / 2), 9);
                m_Symbolic.Outputs["BOT2"] = bot2;

                Vector normal = new Position(boltL / 2, 0, A / 2).Subtract(new Position(-boltL / 2, 0, A / 2));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-boltL / 2, 0, A / 2);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topBolt = symbolGeometryHelper.CreateCylinder(null, M / 2, normal.Length);
                m_Symbolic.Outputs["TOP_BOLT"] = topBolt;


                Vector normal1 = new Position(boltL / 2, 0, -A / 2).Subtract(new Position(-boltL / 2, 0, -A / 2));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-boltL / 2, 0, -A / 2);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d botBolt = symbolGeometryHelper.CreateCylinder(null, M / 2, normal1.Length);
                m_Symbolic.Outputs["BOT_BOLT"] = botBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                Vector v1 = new Vector(0, 0, 1);
                symbolGeometryHelper.SetOrientation(v1, v1.GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + T, B);
                Matrix4X4 matrix = new Matrix4X4();
                matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -B / 2, 0));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5401.cs."));
                return;
            }
        }
        #endregion

    }

}
