//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_FOUR_HOLE_PLATE2.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_FOUR_HOLE_PLATE2
//   Author       :   
//   Creation Date: 15.10.2014  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   15.10.2014      Sasidhar  CR-CP-262765  New 
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
    public class Utility_FOUR_HOLE_PLATE2 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_FOUR_HOLE_PLATE2"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "C1", "C1", 0.999999)]
        public InputDouble m_C1;
        [InputDouble(3, "C2", "C2", 0.999999)]
        public InputDouble m_C2;
        [InputDouble(4, "HOLE_SIZE", "HOLE_SIZE", 0.999999)]
        public InputDouble m_HOLE_SIZE;
        [InputDouble(5, "THICKNESS", "THICKNESS", 0.999999)]
        public InputDouble m_THICKNESS;
        [InputDouble(6, "WIDTH", "WIDTH", 0.999999)]
        public InputDouble m_WIDTH;
        [InputDouble(7, "DEPTH", "DEPTH", 0.999999)]
        public InputDouble m_DEPTH;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("LINE", "LINE")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("R_F_HOLE", "R_F_HOLE")]
        [SymbolOutput("R_B_HOLE", "R_B_HOLE")]
        [SymbolOutput("L_F_HOLE", "L_F_HOLE")]
        [SymbolOutput("L_B_HOLE", "L_B_HOLE")]
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
                Part part =(Part)m_PartInput.Value ;

                Double C1 = m_C1.Value;
                Double C2 = m_C2.Value;
                Double holesize = m_HOLE_SIZE.Value;
                Double width = m_WIDTH.Value;
                Double thickness = m_THICKNESS.Value;
                Double depth = m_DEPTH.Value;

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
                if (holesize <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrHoleDiameterGTZero, "Hole Diameter should be greater than zero"));
                    return;
                }

                Port port1 = new Port(OccurrenceConnection, part, "TopStructure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "BotStructure", new Position(0, 0, thickness), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d body = (Projection3d)symbolGeometryHelper.CreateBox(null, thickness, depth, width);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position((depth / 2 - C1), (width / 2 - C2), -0.00001);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d topleft = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holesize / 2, thickness + 0.00002);
                m_Symbolic.Outputs["L_B_HOLE"] = topleft;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(depth / 2 - C1), (width / 2 - C2), -0.00001);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d topright = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holesize / 2, thickness + 0.00002);
                m_Symbolic.Outputs["R_B_HOLE"] = topright;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position((depth / 2 - C1), -(width / 2 - C2), -0.00001);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d bottomleft = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holesize / 2, thickness + 0.00002);
                m_Symbolic.Outputs["L_F_HOLE"] = bottomleft;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(depth / 2 - C1), -(width / 2 - C2), -0.00001);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d bottomright = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holesize / 2, thickness + 0.00002);
                m_Symbolic.Outputs["R_F_HOLE"] = bottomright;

                Line3d line = new Line3d(OccurrenceConnection, new Position(-depth / 2, 0, thickness / 2), new Position(depth / 2, 0, thickness / 2));
                m_Symbolic.Outputs["LINE"] = line;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_FOUR_HOLE_PLATE"));
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
                Double widthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_FOUR_HOLE_PLAT2", "WIDTH")).PropValue;
                Double depthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_FOUR_HOLE_PLAT2", "DEPTH")).PropValue;
                Double holeSizeValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_FOUR_HOLE_PLAT2", "HOLE_SIZE")).PropValue;
                Double C1Value = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_FOUR_HOLE_PLAT2", "C1")).PropValue;
                Double C2Value = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_FOUR_HOLE_PLAT2", "C2")).PropValue;

                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                Double thicknessValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLAT2", "THICKNESS")).PropValue;

                String width = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, widthValue, UnitName.DISTANCE_INCH);
                String depth = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, depthValue, UnitName.DISTANCE_INCH);
                String holeSize = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, holeSizeValue, UnitName.DISTANCE_INCH);
                String C1 = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, C1Value, UnitName.DISTANCE_INCH);
                String C2 = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, C2Value, UnitName.DISTANCE_INCH);
                String thickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, thicknessValue, UnitName.DISTANCE_INCH);

                bomString = thickness + " Plate Steel, " + width + " X " + depth + " 4 " + holeSize + " dia. holes spaced " + C1 + " and " + C2 + " from edges";
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_FOUR_HOLE_PLATE"));
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
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                Double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_FOUR_HOLE_PLAT2", "WIDTH")).PropValue;
                Double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_FOUR_HOLE_PLAT2", "DEPTH")).PropValue;
                Double holesize = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_FOUR_HOLE_PLAT2", "HOLE_SIZE")).PropValue;
                Double T = (Double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLAT2", "THICKNESS")).PropValue;

                weight = ((depth * width * T) - (4 * (holesize / 2 * holesize / 2 * T * (Math.PI)))) * getSteelDensityKGPerM;

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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_FOUR_HOLE_PLATE"));
                }
            }
        }
    }
        #endregion
}
