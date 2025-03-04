//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   SFS5370.cs
//    FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5370
//    Author       :   Vijaya
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
    public class SFS5370 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5370"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "CtoC", "CtoC", 0.999999)]
        public InputDouble m_dCtoC;
        [InputDouble(4, "Gap", "Gap", 0.999999)]
        public InputDouble m_dGap;
        [InputDouble(5, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(6, "BoltDia", "BoltDia", 0.999999)]
        public InputDouble m_dBoltDia;
        [InputDouble(7, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        [InputDouble(8, "Width", "Width", 0.999999)]
        public InputDouble m_dWidth;
        [InputDouble(9, "HoleInset", "HoleInset", 0.999999)]
        public InputDouble m_dHoleInset;
        [InputDouble(10, "BoltL", "BoltL", 0.999999)]
        public InputDouble m_dBoltL;
        [InputDouble(11, "MinInsulat", "MinInsulat", 0.999999)]
        public InputDouble m_dMinInsulat;
        [InputString(12, "BoltSize", "BoltSize", "No Value")]
        public InputString m_BoltSize;
        [InputString(13, "Material", "Material", "No Value")]
        public InputString m_Material;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
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
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 rotateMatrix = new Matrix4X4();

                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double CtoC = m_dCtoC.Value;
                Double gap = m_dGap.Value;
                Double length = m_dL.Value;
                Double boltDiameter = m_dBoltDia.Value;
                Double thickness = m_dThickness.Value;
                Double width = m_dWidth.Value;
                Double holeInset = m_dHoleInset.Value;
                Double boltL = m_dBoltL.Value;
                Double minInsulat = m_dMinInsulat.Value;
                String boltSize = m_BoltSize.Value;
                String material = m_Material.Value;

                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, CtoC / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "Pin2", new Position(0, 0, -CtoC / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;


                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidWidthGTZero, "Width should be greater than zero"));
                    return;
                }
                if (thickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidThicknessGTZero, "Thickness should be greater than zero"));
                    return;
                }
                if (boltDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidBoltDiaGTZero, "BoltDia value should be greater than zero"));
                    return;
                }


                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                
                symbolGeometryHelper.ActivePosition = new Position(-(gap / 2 + thickness), -width / 2, pipeDiameter / 2);
                Projection3d top1 = (Projection3d)symbolGeometryHelper.CreateBox(null, thickness, width, (CtoC / 2 - pipeDiameter / 2) + holeInset, 9);
                m_Symbolic.Outputs["TOP1"] = top1;


                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(gap / 2, -width / 2, pipeDiameter / 2);
                Projection3d top2 = (Projection3d)symbolGeometryHelper.CreateBox(null, thickness, width, (CtoC / 2 - pipeDiameter / 2) + holeInset, 9);
                m_Symbolic.Outputs["TOP2"] = top2;


                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(gap / 2 + thickness), -width / 2, -length / 2);
                Projection3d bot1 = (Projection3d)symbolGeometryHelper.CreateBox(null, thickness, width, (CtoC / 2 - pipeDiameter / 2) + holeInset, 9);
                m_Symbolic.Outputs["BOT1"] = bot1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(gap / 2, -width / 2, -length / 2);
                Projection3d bot2 = (Projection3d)symbolGeometryHelper.CreateBox(null, thickness, width, (CtoC / 2 - pipeDiameter / 2) + holeInset, 9);
                m_Symbolic.Outputs["BOT2"] = bot2;


                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                Vector normal1 = new Vector(0, 0, 1);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                rotateMatrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                rotateMatrix.Translate(new Vector(0, -width / 2, 0));
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2.0 + thickness, width);
                body.Transform(rotateMatrix);
                m_Symbolic.Outputs["BODY"] = body;


                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector topnormal = new Position(boltL / 2, 0, CtoC / 2).Subtract(new Position(-(boltL / 2), 0, CtoC / 2));
                symbolGeometryHelper.ActivePosition = new Position(-(boltL / 2), 0, CtoC / 2);
                symbolGeometryHelper.SetOrientation(topnormal, topnormal.GetOrthogonalVector());
                Projection3d topBolt = symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2, topnormal.Length);
                m_Symbolic.Outputs["TOP_BOLT"] = topBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector bottomnormal = new Position(boltL / 2, 0, -CtoC / 2).Subtract(new Position(-(boltL / 2), 0, -CtoC / 2));
                symbolGeometryHelper.ActivePosition = new Position(-(boltL / 2), 0, -CtoC / 2);
                symbolGeometryHelper.SetOrientation(bottomnormal, bottomnormal.GetOrthogonalVector());
                Projection3d botBolt = symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2, bottomnormal.Length);
                m_Symbolic.Outputs["BOT_BOLT"] = botBolt;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5370."));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                string material = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAFINL_Material", "Material")).PropValue;
                double pipeND = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAFINL_PipeND_mm", "PipeND")).PropValue;
                string pipeNDiameter = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, pipeND, UnitName.DISTANCE_MILLIMETER);
                string[] s = pipeNDiameter.Split('.');
                bomDescription = "Pipe Clamp A SFS5370 DN " + s[0];
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of SFS5370."));
                return "";
            }
        }
        #endregion

    }

}
