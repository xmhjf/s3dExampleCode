//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_U_BOLT_PLATE.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_U_BOLT_PLATE
//   Author       :  Sasidhar
//   Creation Date:  04/11/2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   04/11/2012   Sasidhar   CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//   30/Dec/2014    PVK       TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using System.Linq;
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
    public class Utility_U_BOLT_PLATE : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_U_BOLT_PLATE"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "THICKNESS", "THICKNESS", 1)]
        public InputDouble m_THICKNESS;
        [InputDouble(3, "WIDTH", "WIDTH", 0.999999)]
        public InputDouble m_WIDTH;
        [InputDouble(4, "DEPTH", "DEPTH", 0.999999)]
        public InputDouble m_DEPTH;
        [InputDouble(5, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_PIPE_DIA;
        [InputDouble(6, "SPACING", "SPACING", 0.999999)]
        public InputDouble m_SPACING;
        [InputDouble(7, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_ROD_DIA;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("LINE", "LINE")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("R_HOLE", "R_HOLE")]
        [SymbolOutput("L_HOLE", "L_HOLE")]
        public AspectDefinition m_oSymbolic;

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
                Part part =(Part)m_PartInput.Value ;

                Double thickness = m_THICKNESS.Value;
                Double width = m_WIDTH.Value;
                Double depth = m_DEPTH.Value;
                Double pipeDiameter = m_PIPE_DIA.Value;
                Double spacing = m_SPACING.Value;
                Double rodDiameter = m_ROD_DIA.Value;
                Double C = (width - spacing) / 2;
                Double holeSize = rodDiameter + 0.0015875;
                Double T;

                PropertyValueCodelist thicknessCodeList = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrUtility_U_BOLT_PLATE", "THICKNESS");
                CodelistItem codeList = thicknessCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)thickness);

                if (codeList != null)
                {
                    if (codeList.Value < 1 || codeList.Value > 12)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidPlateThickness, "Thickness should be between 1 to 12"));
                    }
                }
                if (depth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrDepthGTZero, "Depth should be greater than zero"));
                    return;
                }
                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWidthGTZero, "Width should be greater than zero"));
                    return;
                }
                if (codeList != null)
                    T = Double.Parse(codeList.ShortDisplayName) * 25.4 / 1000;
                else
                    T = 0.0127;

                Port port1 = new Port(OccurrenceConnection, part, "TopStructure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oSymbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "BotStructure", new Position(0, 0, T), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oSymbolic.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d body = (Projection3d)symbolGeometryHelper.CreateBox(null, T, depth, width);
                m_oSymbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, (width / 2 - C), -0.00001);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d rhole = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holeSize / 2, T + 0.00002);
                m_oSymbolic.Outputs["R_HOLE"] = rhole;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -(width / 2 - C), -0.00001);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d lHole = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holeSize / 2, T + 0.00002);
                m_oSymbolic.Outputs["L_HOLE"] = lHole;

                Line3d line = new Line3d(OccurrenceConnection, new Position(-depth / 2, 0, T / 2), new Position(depth / 2, 0, T / 2));
                m_oSymbolic.Outputs["LINE"] = line;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_U_BOLT_PLATE"));
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
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_U_BOLT_PLATE", "THICKNESS");
                string T = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;

                Double CValue, holeSizeValue;

                String[] bomUnits;
                String[] bomHoleSizeUnits;

                Double widthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_U_BOLT_PLATE", "WIDTH")).PropValue;
                Double depthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_U_BOLT_PLATE", "DEPTH")).PropValue;

                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                Double rodDiameterValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrUtility_U_BOLT_PLATE", "ROD_DIA")).PropValue;

                Double spacingValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrUtility_U_BOLT_PLATE", "SPACING")).PropValue;

                string width = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, widthValue, UnitName.DISTANCE_INCH);
                string depth = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, depthValue, UnitName.DISTANCE_INCH);
                string rodDiameter = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, rodDiameterValue, UnitName.DISTANCE_INCH);

                bomUnits = width.Split(' ');
                bomHoleSizeUnits = rodDiameter.Split(' ');

                CValue = (widthValue - spacingValue) / 2;
                holeSizeValue = rodDiameterValue + 0.0015875;

                CValue = Double.Parse(MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, CValue, UnitName.DISTANCE_INCH).Split(' ').First());
                holeSizeValue = Double.Parse(MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, holeSizeValue, UnitName.DISTANCE_INCH).Split(' ').First());

                bomString = T + "in Plate Steel, " + width + " X " + depth + " with Two " + Microsoft.VisualBasic.Conversion.Str(holeSizeValue) + " " + bomHoleSizeUnits[1] + " " + " Holes, inset " + Microsoft.VisualBasic.Conversion.Str(CValue) + " " + bomUnits[1] + " from ends ";

                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_U_BOLT_PLATE"));
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
                const int getSteelDensityKGPerM = 7900;

                Double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_U_BOLT_PLATE", "WIDTH")).PropValue;
                Double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_U_BOLT_PLATE", "DEPTH")).PropValue;
                PropertyValueCodelist thicknessCodeList = (PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrUtility_U_BOLT_PLATE", "THICKNESS");

                Double T = Double.Parse((thicknessCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(thicknessCodeList.PropValue).DisplayName)) * 25.4 / 1000;

                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                Double holeSize = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrUtility_U_BOLT_PLATE", "ROD_DIA")).PropValue + 0.0625;

                weight = ((depth * width * T) - (2 * (holeSize / 2 * holeSize / 2 * T * Math.PI))) * getSteelDensityKGPerM;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_U_BOLT_PLATE"));
                }
            }
        }

        #endregion
    }
}
