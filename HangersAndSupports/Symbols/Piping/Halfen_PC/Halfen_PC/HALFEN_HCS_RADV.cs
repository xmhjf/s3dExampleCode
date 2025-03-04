//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   HALFEN_HCS_RADV.cs
//    Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HCS_RADV
//   Author       :Sasidahr  
//   Creation Date:21-11-2012  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-11-2012      Sasidahr  CR-CP-222275 Converted VB HS_HALFEN_PC Project to C#.Net
//   20-03-2013      Vijay     DI-CP-228142  Modify the error handling for delivered H&S symbols
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
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
    [VariableOutputs]
    public class HALFEN_HCS_RADV : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {

        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HCS_RADV"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Thermal_Expansion_Type", "Thermal_Expansion_Type", 1)]
        public InputDouble m_dThermal_Expansion_Type;
        [InputDouble(3, "Finish", "Finish", 1)]
        public InputDouble m_dFinish;
        [InputDouble(4, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputString(5, "SIZE", "SIZE", "No Value")]
        public InputString m_sSIZE;
        [InputDouble(6, "K_2", "K_2", 0.999999)]
        public InputDouble m_dK_2;
        [InputDouble(7, "K_2D", "K_2D", 0.999999)]
        public InputDouble m_dK_2D;
        [InputDouble(8, "L_2", "L_2", 0.999999)]
        public InputDouble m_dL_2;
        [InputDouble(9, "L_2D", "L_2D", 0.999999)]
        public InputDouble m_dL_2D;
        [InputDouble(10, "Th", "Th", 0.999999)]
        public InputDouble m_dTh;
        [InputDouble(11, "Depth", "Depth", 0.999999)]
        public InputDouble m_dDepth;
        [InputDouble(12, "Rubber_Depth", "Rubber_Depth", 0.999999)]
        public InputDouble m_dRubber_Depth;
        [InputDouble(13, "Rubber_Th", "Rubber_Th", 0.999999)]
        public InputDouble m_dRubber_Th;
        [InputDouble(14, "Spacer_W", "Spacer_W", 0.999999)]
        public InputDouble m_dSpacer_W;
        [InputDouble(15, "Spacer_Th", "Spacer_Th", 0.999999)]
        public InputDouble m_dSpacer_Th;
        [InputString(16, "STOCK_NUMBER_2_FV", "STOCK_NUMBER_2_FV", "No Value")]
        public InputString m_sSTOCK_NUMBER_2_FV;
        [InputString(17, "STOCK_NUMBER_2_A4", "STOCK_NUMBER_2_A4", "No Value")]
        public InputString m_sSTOCK_NUMBER_2_A4;
        [InputString(18, "STOCK_NUMBER_2D_FV", "STOCK_NUMBER_2D_FV", "No Value")]
        public InputString m_sSTOCK_NUMBER_2D_FV;
        [InputString(19, "STOCK_NUMBER_2D_A4", "STOCK_NUMBER_2D_A4", "No Value")]
        public InputString m_sSTOCK_NUMBER_2D_A4;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("BODY1", "BODY1")]
        [SymbolOutput("BODY2", "BODY2")]
        [SymbolOutput("BODY3", "BODY3")]
        [SymbolOutput("SPACER1", "SPACER1")]
        [SymbolOutput("SPACER2", "SPACER2")]
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
                SP3DConnection connection = default(SP3DConnection);
                connection = OccurrenceConnection;

                Double pipeDia = m_dPIPE_DIA.Value;
                Double K_2 = m_dK_2.Value;
                Double K_2D = m_dK_2D.Value;
                Double L_2 = m_dL_2.Value;
                Double L_2D = m_dL_2D.Value;
                Double th = m_dTh.Value;
                Double depth = m_dDepth.Value;
                Double rubberDepth = m_dRubber_Depth.Value;
                Double rubberThickness = m_dRubber_Th.Value;
                Double spacerWidth = m_dSpacer_W.Value;
                Double spacerThickness = m_dSpacer_Th.Value;
                double spacerDepth = 0;
                String type = Convert.ToString(m_dThermal_Expansion_Type.Value);
                String finish = Convert.ToString(m_dFinish.Value);
                Double K, L;
                String stock_Number_2_FV = m_sSTOCK_NUMBER_2_FV.Value;
                String stock_Number_2_A4 = m_sSTOCK_NUMBER_2_A4.Value;
                String stock_Number_2D_FV = m_sSTOCK_NUMBER_2D_FV.Value;
                String stock_Number_2D_A4 = m_sSTOCK_NUMBER_2D_A4.Value;

                if (m_dThermal_Expansion_Type.Value < 1 || m_dThermal_Expansion_Type.Value > 2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrThermalValue, "Thermal Expansion Type CodeList value should between 1 to 2"));
                    return;
                }
                if (type == "2")
                {
                    K = K_2;
                    L = L_2;
                }
                else
                {
                    K = K_2D;
                    L = L_2D;
                }
                rubberThickness = L / 2.0 - pipeDia / 2.0;
                Double saddleWidth = L / 2.0;
                Double saddleCalc = Math.Sqrt(Math.Abs(((pipeDia / 2.0 * pipeDia / 2.0) - (saddleWidth / 2.0 * saddleWidth / 2.0))));
                Double saddleThickness = rubberThickness / 2.0;
                if (type == "1")
                {
                    spacerWidth = 0;
                    spacerThickness = 0;
                }
                Line3d line;
                Collection<ICurve> curvecollection = new Collection<ICurve>();

                //ports

                Port port1 = new Port(connection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(connection, part, "RightHole", new Position(0, K / 2 - (K / 2 - L / 2) / 2, pipeDia / 2 + saddleThickness), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Port port3 = new Port(connection, part, "LeftHole", new Position(0, -(K / 2 - (K / 2 - L / 2) / 2), pipeDia / 2 + saddleThickness), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                line = new Line3d(new Position(-rubberDepth / 2, -L / 2 - th, 0), new Position(-rubberDepth / 2, -L / 2 - th, -pipeDia / 4));
                curvecollection.Add(line);

                line = new Line3d(new Position(-rubberDepth / 2, -L / 2 - th, -pipeDia / 4), new Position(-rubberDepth / 2, -pipeDia / 2, -pipeDia / 4));
                curvecollection.Add(line);

                line = new Line3d(new Position(-rubberDepth / 2, -pipeDia / 2, -pipeDia / 4), new Position(-rubberDepth / 2, -pipeDia / 2, 0));
                curvecollection.Add(line);

                line = new Line3d(new Position(-rubberDepth / 2, pipeDia / 2, 0), new Position(-rubberDepth / 2, pipeDia / 2, -pipeDia / 4));
                curvecollection.Add(line);

                line = new Line3d(new Position(-rubberDepth / 2, pipeDia / 2, -pipeDia / 4), new Position(-rubberDepth / 2, L / 2 + th, -pipeDia / 4));
                curvecollection.Add(line);

                line = new Line3d(new Position(-rubberDepth / 2, L / 2 + th, 0), new Position(-rubberDepth / 2, L / 2 + th, -pipeDia / 4));
                curvecollection.Add(line);

                Arc3d outerArc = new Arc3d(new Position(-rubberDepth / 2, 0, 0), new Vector(-1, 0, 0), new Position(-rubberDepth / 2, -L / 2 - th, 0), new Position(-rubberDepth / 2, L / 2 + th, 0));
                curvecollection.Add(outerArc);

                Arc3d innerArc = new Arc3d(new Position(-rubberDepth / 2, 0, 0), new Vector(-1, 0, 0), new Position(-rubberDepth / 2, -pipeDia / 2, 0), new Position(-rubberDepth / 2, pipeDia / 2, 0));
                curvecollection.Add(innerArc);

                ComplexString3d lineString = new ComplexString3d(curvecollection);
                Projection3d body1 = new Projection3d(lineString, new Vector(1, 0, 0), rubberDepth, true);
                m_Symbolic.Outputs["BODY1"] = body1;

                curvecollection = new Collection<ICurve>();

                Line3d bottom = new Line3d(new Position(-rubberDepth / 2, -saddleWidth / 2, -pipeDia / 2 - saddleThickness), new Position(-rubberDepth / 2, saddleWidth / 2, -pipeDia / 2 - saddleThickness));
                curvecollection.Add(bottom);

                line = new Line3d(new Position(-rubberDepth / 2, saddleWidth / 2, -saddleCalc), new Position(-rubberDepth / 2, 0, -pipeDia / 2));
                curvecollection.Add(line);

                line = new Line3d(new Position(-rubberDepth / 2, 0, -pipeDia / 2), new Position(-rubberDepth / 2, -saddleWidth / 2, -saddleCalc));
                curvecollection.Add(line);

                line = new Line3d(new Position(-rubberDepth / 2, saddleWidth / 2, -pipeDia / 2 - saddleThickness), new Position(-rubberDepth / 2, saddleWidth / 2, -saddleCalc));
                curvecollection.Add(line);

                line = new Line3d(new Position(-rubberDepth / 2, -saddleWidth / 2, -pipeDia / 2 - saddleThickness), new Position(-rubberDepth / 2, -saddleWidth / 2, -saddleCalc));
                curvecollection.Add(line);

                lineString = new ComplexString3d(curvecollection);
                Projection3d body2 = new Projection3d(lineString, new Vector(1, 0, 0), rubberDepth, true);
                m_Symbolic.Outputs["BODY2"] = body2;

                curvecollection = new Collection<ICurve>();
                outerArc = new Arc3d(new Position(-depth / 2, 0, 0), new Vector(-1, 0, 0), new Position(-depth / 2, -L / 2 - th, 0), new Position(-depth / 2, L / 2 + th, 0));
                curvecollection.Add(outerArc);

                line = new Line3d(new Position(-depth / 2, -L / 2 - th, 0), new Position(-depth / 2, -L / 2 - th, -pipeDia / 2 - saddleThickness + spacerThickness + th));
                curvecollection.Add(line);

                Line3d topLeft = new Line3d(new Position(-depth / 2, -L / 2 - th, -pipeDia / 2 - saddleThickness + spacerThickness + th), new Position(-depth / 2, -K / 2, -pipeDia / 2 - saddleThickness + spacerThickness + th));
                curvecollection.Add(topLeft);

                Line3d veryLeft = new Line3d(new Position(-depth / 2, -K / 2, -pipeDia / 2 - saddleThickness + spacerThickness + th), new Position(-depth / 2, -K / 2, -pipeDia / 2 - saddleThickness + spacerThickness));
                curvecollection.Add(veryLeft);

                Line3d bottomLeft = new Line3d(new Position(-depth / 2, -K / 2, -pipeDia / 2 - saddleThickness + spacerThickness), new Position(-depth / 2, -L / 2, -pipeDia / 2 - saddleThickness + spacerThickness));
                curvecollection.Add(bottomLeft);

                line = new Line3d(new Position(-depth / 2, -L / 2, -pipeDia / 2 - saddleThickness + spacerThickness), new Position(-depth / 2, -L / 2, 0));
                curvecollection.Add(line);

                innerArc = new Arc3d(new Position(-depth / 2, 0, 0), new Vector(-1, 0, 0), new Position(-depth / 2, -L / 2, 0), new Position(-depth / 2, L / 2, 0));
                curvecollection.Add(innerArc);

                line = new Line3d(new Position(-depth / 2, L / 2, 0), new Position(-depth / 2, L / 2, -pipeDia / 2 - saddleThickness + spacerThickness));
                curvecollection.Add(line);

                Line3d bottomRight = new Line3d(new Position(-depth / 2, L / 2, -pipeDia / 2 - saddleThickness + spacerThickness), new Position(-depth / 2, K / 2, -pipeDia / 2 - saddleThickness + spacerThickness));
                curvecollection.Add(bottomRight);

                Line3d veryRight = new Line3d(new Position(-depth / 2, K / 2, -pipeDia / 2 - saddleThickness + spacerThickness + th), new Position(-depth / 2, K / 2, -pipeDia / 2 - saddleThickness + spacerThickness));
                curvecollection.Add(veryRight);

                Line3d topRight = new Line3d(new Position(-depth / 2, K / 2, -pipeDia / 2 - saddleThickness + spacerThickness + th), new Position(-depth / 2, L / 2 + th, -pipeDia / 2 - saddleThickness + spacerThickness + th));
                curvecollection.Add(topRight);

                line = new Line3d(new Position(-depth / 2, L / 2 + th, -pipeDia / 2 - saddleThickness + spacerThickness + th), new Position(-depth / 2, L / 2 + th, 0));
                curvecollection.Add(line);

                lineString = new ComplexString3d(curvecollection);
                Projection3d body3 = new Projection3d(lineString, new Vector(1, 0, 0), depth, true);
                m_Symbolic.Outputs["BODY3"] = body3;

                if (type.Equals("2"))
                {
                    spacerDepth = (K / 2 - L / 2) + 0.01;
                    SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(K / 2 - (K / 2 - L / 2) / 2 - spacerDepth / 2, -spacerWidth / 2, -pipeDia / 2 - saddleThickness + spacerThickness);
                    Projection3d spacer1 = (Projection3d)symbolGeometryHelper.CreateBox(null, spacerDepth, spacerWidth, spacerThickness,9);
                    Matrix4X4 rotateMatrix = new Matrix4X4();
                    rotateMatrix.SetIdentity();
                    rotateMatrix.Rotate( 3*Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));                    
                    spacer1.Transform(rotateMatrix);
                    m_Symbolic.Outputs["SPACER1"] = spacer1;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-(K / 2 - (K / 2 - L / 2) / 2) - spacerDepth / 2, -spacerWidth / 2, -pipeDia / 2 - saddleThickness + spacerThickness);
                    Projection3d spacer2 = (Projection3d)symbolGeometryHelper.CreateBox(null, spacerDepth, spacerWidth, spacerThickness, 9);
                    rotateMatrix.SetIdentity();
                    rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    spacer2.Transform(rotateMatrix);
                    m_Symbolic.Outputs["SPACER2"] = spacer2;
                }
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of HALFEN_HCS_RADV.cs."));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrFinish_List", "Finish");
                String finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;
                PropertyValueCodelist thermalCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrRADVType", "Thermal_Expansion_Type");
                String partType = thermalCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(thermalCodelist.PropValue).DisplayName;
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                String stockNumber = ((PropertyValueString)part.GetPropertyValue("IJUAHgrSTOCK_NUMBER_" + partType + "_" + (finish.ToUpper()), "STOCK_NUMBER_" + partType + "_" + (finish.ToUpper()))).PropValue;

                bomString = part.PartDescription + "/" + partType + " - " + finish + " - " + stockNumber;

                return bomString;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrBOMDescription, "Error in BOMDescription of HALFEN_HCS_RADV.cs."));
                }
                return "";
            }
        }
        #endregion
    
        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Double weight, cogX, cogY, cogZ;
                PropertyValueCodelist thermalCodelist = (PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrRADVType", "Thermal_Expansion_Type");
                String type = thermalCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(thermalCodelist.PropValue).DisplayName;
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrWEIGHT_" + type, "WEIGHT_" + type)).PropValue;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrWeightCG, "Error in WeightCG of HALFEN_HCS_RADV.cs."));
                }
            }
        }
        #endregion

    }

}
