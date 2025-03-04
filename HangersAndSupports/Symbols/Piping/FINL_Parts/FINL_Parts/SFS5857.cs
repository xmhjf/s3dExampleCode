//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   SFS5857.cs
//    FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5857
//   Author       :   Rajeswari
//   Creation Date:  18-March-2013
//   Description:    CR-CP-222272-Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 18-March-2013 Rajeswari CR-CP-222272-Initial Creation
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
    [VariableOutputs]
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class SFS5857 : CustomSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5857"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(4, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(5, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(6, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(7, "M", "M", 0.999999)]
        public InputDouble m_dM;
        [InputDouble(8, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(9, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(10, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(11, "BoltL", "BoltL", 0.999999)]
        public InputDouble m_dBoltL;
        [InputDouble(12, "MaxLoad200", "MaxLoad200", 0.999999)]
        public InputDouble m_dMaxLoad200;
        [InputDouble(13, "MaxLoad300", "MaxLoad300", 0.999999)]
        public InputDouble m_dMaxLoad300;
        [InputDouble(14, "MaxLoad480", "MaxLoad480", 0.999999)]
        public InputDouble m_dMaxLoad480;
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
                 
               

                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double A = m_dA.Value;
                Double F = m_dF.Value;
                Double L = m_dL.Value;
                Double E = m_dE.Value;
                Double M = m_dM.Value;
                Double T = m_dT.Value;
                Double B = m_dB.Value;
                Double C = m_dC.Value;
                Double boltL = m_dBoltL.Value;

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, A / 2.0 + E), new Vector(1, 0, 0), new Vector(0, 0, 1));
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
                if (A <= 0 && pipeDiameter <= 0 && C <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidAandpipeDiaandCGTZero, "A, pipeDiameter and C values should be greater than zero"));
                    return;
                }
                //=================================================
                //Construction of Physical Aspect 
                //=================================================              

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

               

                symbolGeometryHelper.ActivePosition = new Position(-(F / 2.0 + T), -B / 2.0, pipeDiameter / 2.0);
                Projection3d top1 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, B, (A / 2.0 - pipeDiameter / 2.0) + E + C, 9);
                m_Symbolic.Outputs["TOP1"] = top1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(F / 2.0, -B / 2.0, pipeDiameter / 2.0);
                Projection3d top2 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, B, (A / 2.0 - pipeDiameter / 2.0) + E + C, 9);
                m_Symbolic.Outputs["TOP2"] = top2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(F / 2.0 + T), -B / 2.0, -A / 2.0 - C);
                Projection3d bot1 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, B, (A / 2.0 - pipeDiameter / 2.0) + C, 9);
                m_Symbolic.Outputs["BOT1"] = bot1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(F / 2.0, -B / 2.0, -A / 2.0 - C);
                Projection3d bot2 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, B, (A / 2.0 - pipeDiameter / 2.0) + C, 9);
                m_Symbolic.Outputs["BOT2"] = bot2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                Vector v1 = new Vector(0, 0, 1);
                symbolGeometryHelper.SetOrientation(v1, v1.GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2.0 + T, B);
                Matrix4X4 matrix = new Matrix4X4();
                matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -B / 2.0, 0));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;

                Vector normal = new Position(boltL / 2.0, 0, A / 2.0 + E).Subtract(new Position(-boltL / 2.0, 0, A / 2.0 + E));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-boltL / 2.0, 0, A / 2.0 + E);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topBolt = symbolGeometryHelper.CreateCylinder(null, M / 2.0, normal.Length);
                m_Symbolic.Outputs["TOP_BOLT"] = topBolt;

                Vector normal1 = new Position(boltL / 2.0, 0, -A / 2.0).Subtract(new Position(-boltL / 2.0, 0, -A / 2.0));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-boltL / 2.0, 0, -A / 2.0);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d botBolt = symbolGeometryHelper.CreateCylinder(null, M / 2.0, normal1.Length);
                m_Symbolic.Outputs["BOT_BOLT"] = botBolt;

                Vector normal2 = new Position(boltL / 2.0, 0, A / 2.0).Subtract(new Position(-boltL / 2.0, 0, A / 2.0));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-boltL / 2.0, 0, A / 2.0);
                symbolGeometryHelper.SetOrientation(normal2, normal2.GetOrthogonalVector());
                Projection3d botBolt1 = symbolGeometryHelper.CreateCylinder(null, M / 2.0, normal2.Length);
                m_Symbolic.Outputs["BOT_BOLT_"] = botBolt1;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5857.cs."));
                return;
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
                 double temp = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAFINL_Temperature", "Temperature")).PropValue;
                 temp = temp * 1000;
                 double maxLoad = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJOAFINL_MaxLoad200", "MaxLoad200")).PropValue;
                 if (temp > 20.0)
                 {
                     maxLoad = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJOAFINL_MaxLoad300", "MaxLoad300")).PropValue;
                 }
                 if (temp > 300.0)
                 {
                     maxLoad = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJOAFINL_MaxLoad480", "MaxLoad480")).PropValue;
                 }
                 oSupportOrComponent.SetPropertyValue(maxLoad, "IJOAFINL_MaxLoad", "MaxLoad");

                 string material = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAFINL_Material", "Material")).PropValue;
                 double pipeNDValue = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAFINL_PipeND_mm", "PipeND")).PropValue;
                 string pipeND = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, pipeNDValue, UnitName.DISTANCE_MILLIMETER).Trim();
                 string[] pipeNDArray = pipeND.Split('.');
                 bomDescription = "Pipe clamp E SFS 5857 DN " + pipeNDArray[0];

                 return bomDescription;
             }
             catch
             {
                 ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of SFS5857.cs."));
                 return "";
             }
        }
        #endregion

    }

}
