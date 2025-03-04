//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_FPR.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_FPR
//   Author       :  Manikanth
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who         change description
//   -----------     ---         ------------------
//   21-08-2013     Manikanth    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net   
//   30-Dec-2014    PVK          TR-CP-264951	Resolve P3 coverity issues found in November 2014 report 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
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
    public class PSL_FPR : HangerComponentSymbolDefinition, ICustomWeightCG, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_FPR"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "F", "F", 0.999999)]
        public InputDouble F;
        [InputString(3, "TEMP2", "TEMP2", "No Value")]
        public InputString TEMP2;
        [InputString(4, "PIPE_NOM_DIA", "PIPE_NOM_DIA", "No Value")]
        public InputString PIPE_NOM_DIA;
        [InputString(5, "LOAD_GRP", "LOAD_GRP", "No Value")]
        public InputString LOAD_GRP;
        [InputDouble(6, "MIN_F", "MIN_F", 0.999999)]
        public InputDouble MIN_F;
        [InputDouble(7, "MAX_F", "MAX_F", 0.999999)]
        public InputDouble MAX_F;
        [InputDouble(8, "MIN_E", "MIN_E", 0.999999)]
        public InputDouble MIN_E;
        [InputDouble(9, "MAX_E", "MAX_E", 0.999999)]
        public InputDouble MAX_E;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("L_BOX", "L_BOX")]
        [SymbolOutput("R_BOX", "R_BOX")]
        public AspectDefinition m_Symbolic;
        #endregion

        #region "Construct Outputs"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)PartInput.Value;
                double f = F.Value;
                double minF = MIN_F.Value;
                double maxF = MAX_F.Value;
                double minE = MIN_E.Value;
                double maxE = MAX_E.Value;

                if ((f < minF) || (f > maxF))
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "Rod Centre must be between " + (minF * 1000) + " mm and " + (maxF * 1000) + "mm.");
                    return;
                }
                string temp2 = TEMP2.Value;
                string f1 = (f * 1000).ToString();
                if (HgrCompareDoubleService.cmpdbl((f * 1000 / 100) * 100.0 , f * 1000)==false)
                    f1 = ((f * 1000 / 100.0) * 100.0 + 100).ToString();
                string partclass = string.Empty, auxInterface = string.Empty;
                string partNumber = "FPR-" + PIPE_NOM_DIA.Value + "/" + f1.Trim() + "-" + LOAD_GRP.Value + "-" + temp2;
                if (temp2.Equals("400"))
                {
                    partclass = "PSL_FPR_AUX_400";
                    auxInterface = "IJUAHgrPSL_FPR_AUX_400";
                }
                if (temp2.Equals("490"))
                {
                    partclass = "PSL_FPR_AUX_490";
                    auxInterface = "IJUAHgrPSL_FPR_AUX_490";
                }
                if (temp2.Equals("530"))
                {
                    partclass = "PSL_FPR_AUX_530";
                    auxInterface = "IJUAHgrPSL_FPR_AUX_530";
                }
                if (temp2.Equals("560"))
                {
                    partclass = "PSL_FPR_AUX_560";
                    auxInterface = "IJUAHgrPSL_FPR_AUX_560";
                }
                if (temp2.Equals("600"))
                {
                    partclass = "PSL_FPR_AUX_600";
                    auxInterface = "IJUAHgrPSL_FPR_AUX_600";
                }

                Dictionary<KeyValuePair<string, string>, Object> parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>(auxInterface, "PART_NUMBER"), partNumber.Trim());

                double gE = (double)PSLSymbolServices.GetDataByMultipleConditions(partclass, auxInterface, "GE", parameter);
                double kW = (double)PSLSymbolServices.GetDataByMultipleConditions(partclass, auxInterface, "KW", parameter);
                double hG = (double)PSLSymbolServices.GetDataByMultipleConditions(partclass, auxInterface, "HG", parameter);
                double D = (double)PSLSymbolServices.GetDataByMultipleConditions(partclass, auxInterface, "D", parameter);
                double gH = (double)PSLSymbolServices.GetDataByMultipleConditions(partclass, auxInterface, "GH", parameter);
                double B = (double)PSLSymbolServices.GetDataByMultipleConditions(partclass, auxInterface, "B", parameter);
                double A = (double)PSLSymbolServices.GetDataByMultipleConditions(partclass, auxInterface, "A", parameter);
                double C = (double)PSLSymbolServices.GetDataByMultipleConditions(partclass, auxInterface, "C", parameter);
                double E = (double)PSLSymbolServices.GetDataByMultipleConditions(partclass, auxInterface, "E", parameter);

                if (HgrCompareDoubleService.cmpdbl(maxF - minF , 0)==false)
                    E = minE + (maxE - minE) * (f - minF) / (maxF - minF);

                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBGTZ, "B Value should be greater than zero"));
                    return;
                }
                if (E <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEGTZ, "E Value should be greater than zero"));
                    return;
                }
                if (kW <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidKWGTZ, "KW Value should be greater than zero"));
                    return;
                }
                if (hG <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidHGGTZ, "HG Value should be greater than zero"));
                    return;
                }
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Eye", new Position(0, -f / 2, C + A / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Port port3 = new Port(OccurrenceConnection, part, "Eye", new Position(0, f / 2, C + A / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;


                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 rotateMatrix = new Matrix4X4();

                rotateMatrix.Translate(new Vector(-(f / 2 + gE / 2 + gH), -hG / 2, -E));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d basePlate = symbolGeometryHelper.CreateBox(null, f + gE + gH * 2, hG, E, 9);
                basePlate.Transform(rotateMatrix);
                m_Symbolic.Outputs["BODY"] = basePlate;

                symbolGeometryHelper = new SymbolGeometryHelper();
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(-(f / 2 + kW / 2), -B / 2, -E));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d web = symbolGeometryHelper.CreateBox(null, kW, B, E + C + D, 9);
                web.Transform(rotateMatrix);
                m_Symbolic.Outputs["L_BOX"] = web;

                symbolGeometryHelper = new SymbolGeometryHelper();
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(f / 2 - kW / 2, -B / 2, -E));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d ep1 = symbolGeometryHelper.CreateBox(null, kW, B, E + C + D, 9);
                ep1.Transform(rotateMatrix);
                m_Symbolic.Outputs["R_BOX"] = ep1;

            }
            catch (Exception)
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_FPR.cs."));
                return;
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomDescrition = "";
            try
            {
                Part part = (Part)SupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                double length = (double)((PropertyValueDouble)SupportOrComponent.GetPropertyValue("IJOAHgrPSL_FPR", "F")).PropValue;

                string size = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_SIZE", "SIZE")).PropValue;
                bomDescrition = "PSL " + size + " Riser Clamp Flat Plate Type, Rod Centre =" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_MILLIMETER); ;
                return bomDescrition;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_FPR."));
                return "";
            }
        }
        #endregion

        #region "ICustomHgrWeightCG Members"
        public void EvaluateWeightCG(BusinessObject SupportOrComponent)
        {
            try
            {
                Part catalogPart = (Part)SupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                double weight = 0;
                double f = (double)((PropertyValueDouble)SupportOrComponent.GetPropertyValue("IJOAHgrPSL_FPR", "F")).PropValue;
                double minWeight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrPSL_FPR", "MIN_WEIGHT")).PropValue;
                double maxWeight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrPSL_FPR", "MAX_WEIGHT")).PropValue;
                double minF = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrPSL_FPR", "MIN_F")).PropValue;
                double maxF = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrPSL_FPR", "MAX_F")).PropValue;
                if (HgrCompareDoubleService.cmpdbl(maxF - minF, 0) == false)
                    weight = minWeight + (maxWeight - minWeight) * (f - minF) / (maxF - minF);
                double cogZ = 0;
                double cogY = 0;
                double cogX = 0;
                SupportComponent supportComponent = (SupportComponent)SupportOrComponent;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);

            }
            catch (Exception)
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PSL_FPR.cs."));
                return;
            }
        }
        #endregion
    }
}
