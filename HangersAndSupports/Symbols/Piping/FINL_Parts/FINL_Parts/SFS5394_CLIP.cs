//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5394_CLIP.cs
//    FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5394_CLIP
//   Author       :  Vijay
//   Creation Date:  18.03.2013
//   Description: CR-CP-222272 Convert HS_FINL_Parts VB Project to C# .Net  

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18.03.2013     Vijay   CR-CP-222272 Convert HS_FINL_Parts VB Project to C# .Net  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

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
    public class SFS5394_CLIP : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5394_CLIP"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "BoltDia", "BoltDia", 0.999999)]
        public InputDouble m_dBoltDia;
        [InputDouble(4, "BoltL", "BoltL", 0.999999)]
        public InputDouble m_dBoltL;
        [InputDouble(5, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(6, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(7, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(8, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(9, "B", "B", 0.999999)]
        public InputDouble m_dB;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("TOP1", "TOP1")]
        [SymbolOutput("TOP2", "TOP2")]
        [SymbolOutput("TOP_BOLT", "TOP_BOLT")]
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
                 
               

                Double pipeDia = m_dPIPE_DIA.Value;
                Double boltDia = m_dBoltDia.Value;
                Double boltL = m_dBoltL.Value;
                Double D = m_dD.Value;
                Double F = m_dF.Value;
                Double E = m_dE.Value;
                Double T = m_dT.Value;
                Double B = m_dB.Value;
                //ports 

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, E), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Port port3 = new Port(OccurrenceConnection, part, "Weld", new Position(0, 0, -pipeDia / 2 - T), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

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
                if (boltDia <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidBoltDiaGTZero, "BoltDia value should be greater than zero"));
                    return;
                }
                if ((pipeDia + T) <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidPipeDiaandTGTZero, "PipeDiameter + T value should be greater than zero"));
                    return;
                }
                if (boltL == 0 && E == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidBoltLandE, "BoltL and E values cannot be zero"));
                    return;
                }                
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                
                symbolGeometryHelper.ActivePosition = new Position(-(F / 2 + T), -B / 2, pipeDia / 2);
                Projection3d top1Box = symbolGeometryHelper.CreateBox(null, T, B, (E - pipeDia / 2) + B, 9);
                m_Symbolic.Outputs["TOP1"] = top1Box;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(F / 2, -B / 2, pipeDia / 2);
                Projection3d top2Box = symbolGeometryHelper.CreateBox(null, T, B, (E - pipeDia / 2) + B, 9);
                m_Symbolic.Outputs["TOP2"] = top2Box;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(boltL / 2, 0, E).Subtract(new Position(-boltL / 2, 0, E));
                symbolGeometryHelper.ActivePosition = new Position(-boltL / 2, 0, E);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topBoltCylinder = symbolGeometryHelper.CreateCylinder(null, boltDia / 2, normal.Length);
                m_Symbolic.Outputs["TOP_BOLT"] = topBoltCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                normal = new Vector(0, 0, 1);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -B / 2, 0));
                Projection3d bodyCylinder = symbolGeometryHelper.CreateCylinder(null, pipeDia / 2.0 + T, B);
                bodyCylinder.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = bodyCylinder;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5394_CLIP.cs."));
                    return;
                }
            }
        }
        #endregion
    }
}
