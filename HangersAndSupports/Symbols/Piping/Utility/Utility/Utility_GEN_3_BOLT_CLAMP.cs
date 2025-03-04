//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_GEN_3_BOLT_CLAMP.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_3_BOLT_CLAMP
//   Author       : Hema 
//   Creation Date: 31.10.2012 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   31.10.2012      Hema    CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		 Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
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
    public class Utility_GEN_3_BOLT_CLAMP : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_3_BOLT_CLAMP"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "R", "R", 0.999999)]
        public InputDouble m_dR;
        [InputDouble(3, "J", "J", 0.999999)]
        public InputDouble m_dJ;
        [InputDouble(4, "K", "K", 0.999999)]
        public InputDouble m_dK;
        [InputDouble(5, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(6, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(7, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(8, "BOLT_DIA", "BOLT_DIA", 0.999999)]
        public InputDouble m_dBOLT_DIA;
        [InputDouble(9, "BOLT_L", "BOLT_L", 0.999999)]
        public InputDouble m_dBOLT_L;
        [InputDouble(10, "BOLT_DIA2", "BOLT_DIA2", 0.999999)]
        public InputDouble m_dBOLT_DIA2;
        [InputDouble(11, "BOLT_L2", "BOLT_L2", 0.999999)]
        public InputDouble m_dBOLT_L2;
        [InputDouble(12, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(13, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(14, "OPT_BOM", "OPT_BOM", 1)]
        public InputDouble m_oOPT_BOM;

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
        [SymbolOutput("MIDDLE_BOLT", "MIDDLE_BOLT")]
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

                Double R = m_dR.Value;
                Double J = m_dJ.Value;
                Double K = m_dK.Value;
                Double A = m_dA.Value;
                Double B = m_dB.Value;
                Double H = m_dH.Value;
                Double bolt_Dia = m_dBOLT_DIA.Value;
                Double bolt_L = m_dBOLT_L.Value;
                Double bolt_Dia2 = m_dBOLT_DIA2.Value;
                Double bolt_L2 = m_dBOLT_L2.Value;
                Double T = m_dT.Value;
                Double W = m_dW.Value;

                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrStockTGTZero, "Stock T should be greater than zero"));
                    return;
                }
                if (A <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrAGTZero, "Pipe CL to Bottom should be greater than zero"));
                    return;
                }
                if (R <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrRGTZero, "Radius should be greater than zero"));
                    return;
                }
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBGTZero, "Pipe CL to Top should be greater than zero"));
                    return;
                }
                if (bolt_Dia <= 0 || bolt_Dia2 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBoltDiameterGTZero, "BoltDiameter should be greater than zero"));
                    return;
                }
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidStockWidth, "Stock Width should be greater than zero"));
                    return;
                }
                if (bolt_L == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidBOlTLENGTH, "Bolt Length cannot be zero"));
                    return;
                }
                if (bolt_L2 == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidTOPBOlTLENGTH, "Top Bolt Length cannot be zero"));
                    return;
                }
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, K - bolt_Dia2 / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, (H / 2 + T), -((A - R) / 2 + R));
                symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(0, 0, 1));
                Projection3d top1 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, A - R, W);
                m_Symbolic.Outputs["TOP1"] = top1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -(H / 2 + T), -((A - R) / 2 + R));
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                Projection3d top2 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, A - R, W);
                m_Symbolic.Outputs["TOP2"] = top2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, (H / 2 + T), ((B - R) / 2 + R));
                symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(0, 0, 1));
                Projection3d bottom1 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, B - R, W);
                m_Symbolic.Outputs["BOT1"] = bottom1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -(H / 2 + T), ((B - R) / 2 + R));
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                Projection3d bottom2 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, B - R, W);
                m_Symbolic.Outputs["BOT2"] = bottom2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, -bolt_L2 / 2, K).Subtract(new Position(0, bolt_L2 / 2, K));
                symbolGeometryHelper.ActivePosition = new Position(0, -bolt_L2 / 2, K);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d top_bolt = (Projection3d)symbolGeometryHelper.CreateCylinder(null, bolt_Dia2 / 2, normal.Length);
                m_Symbolic.Outputs["TOP_BOLT"] = top_bolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal1 = new Position(0, -bolt_L / 2, J).Subtract(new Position(0, bolt_L / 2, J));
                symbolGeometryHelper.ActivePosition = new Position(0, -bolt_L / 2, J);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d middle_bolt = (Projection3d)symbolGeometryHelper.CreateCylinder(null, bolt_Dia / 2, normal1.Length);
                m_Symbolic.Outputs["MIDDLE_BOLT"] = middle_bolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal2 = new Position(0, -bolt_L / 2, -J).Subtract(new Position(0, bolt_L / 2, -J));
                symbolGeometryHelper.ActivePosition = new Position(0, -bolt_L / 2, -J);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d bottom_bolt = (Projection3d)symbolGeometryHelper.CreateCylinder(null, bolt_Dia / 2, normal2.Length);
                m_Symbolic.Outputs["BOT_BOLT"] = bottom_bolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-W / 2, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Projection3d body = (Projection3d)symbolGeometryHelper.CreateCylinder(null, R + T, W);
                m_Symbolic.Outputs["BODY"] = body;

            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_GEN_3_BOLT_CLAMP"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Double weight, cogX, cogY, cogZ;
                const int getSteelDensityKGPerM = 7900;
                double R = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "R")).PropValue;
                double A = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "A")).PropValue;
                double H = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "H")).PropValue;
                double boltDiameter = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "BOLT_DIA")).PropValue;
                double boltL = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "BOLT_L")).PropValue;
                double T = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "T")).PropValue;
                double W = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "W")).PropValue;

                weight = (((R + T) * (R + T) * Math.PI * W) - (R * R * Math.PI * W) - ((T * W * H) * 2) + (((A / 2 - R - T) * W * T) * 4) + ((Math.PI * boltL * (boltDiameter / 2) * (boltDiameter / 2)) * 2)) * getSteelDensityKGPerM;
                cogX = 0;
                cogY = 0;
                cogZ = 0;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_GEN_3_BOLT_CLAMP"));
                }
            }
        }
    }
        #endregion
}


