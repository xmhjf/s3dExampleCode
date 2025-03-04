//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG98F.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG98F
//   Author       :  Rajeswari
//   Creation Date:  13-05-2013
//   Description:
//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   13-05-2013      Rajeswari  CR-CP-233113 Convert HS_Anvil VB Project to C# .Net
//   11-Dec-2014     PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using System.Collections.ObjectModel;
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

    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    public class Anvil_FIG98F : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG98F"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "DIR", "DIR", 1)]
        public InputDouble m_oDIR;
        [InputDouble(3, "WORKING_TRAV", "WORKING_TRAV", 0.999999)]
        public InputDouble m_dWORKING_TRAV;
        [InputDouble(4, "HOT_LOAD", "HOT_LOAD", 0.999999)]
        public InputDouble m_dHOT_LOAD;
        [InputDouble(5, "COL_TYP", "COL_TYP", 1)]
        public InputDouble m_oCOL_TYP;
        [InputDouble(6, "TOP", "TOP", 1)]
        public InputDouble m_oTOP;
        [InputDouble(7, "ROLL_MATERIAL", "ROLL_MATERIAL", 1)]
        public InputDouble m_oROLL_MATERIAL;
        [InputDouble(8, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputString(9, "SIZE", "SIZE", "No Value")]
        public InputString m_oSIZE;
        [InputDouble(10, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(11, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(12, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(13, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(14, "COL_DIA", "COL_DIA", 0.999999)]
        public InputDouble m_dCOL_DIA;
        [InputDouble(15, "FLANGE_T", "FLANGE_T", 0.999999)]
        public InputDouble m_dFLANGE_T;
        [InputDouble(16, "FLANGE_DIA", "FLANGE_DIA", 0.999999)]
        public InputDouble m_dFLANGE_DIA;
        [InputDouble(17, "X", "X", 0.999999)]
        public InputDouble m_dX;
        [InputDouble(18, "ADJ", "ADJ", 0.999999)]
        public InputDouble m_dADJ;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("LOAD_FLANGE", "LOAD_FLANGE")]
        [SymbolOutput("PIPE_ROLL", "PIPE_ROLL")]
        [SymbolOutput("ROLL_1", "ROLL_1")]
        [SymbolOutput("ROLL_2", "ROLL_2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("BOTTOM", "BOTTOM")]
        [SymbolOutput("LOAD_COLUMN", "LOAD_COLUMN")]
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

                Double workingTrav = m_dWORKING_TRAV.Value;
                Double hotLoad = m_dHOT_LOAD.Value;
                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double B = m_dB.Value;
                Double C = m_dC.Value;
                Double E = m_dE.Value;
                Double T = m_dT.Value;
                Double colDiameter = m_dCOL_DIA.Value;
                Double flangeT = m_dFLANGE_T.Value;
                Double flangeDiameter = m_dFLANGE_DIA.Value;
                Double X = m_dX.Value;
                int rollMaterial = (int)m_oROLL_MATERIAL.Value;
                int top = (int)m_oTOP.Value;
                int colTyp = (int)m_oCOL_TYP.Value;
                string size = m_oSIZE.Value;
                if (m_oDIR.Value < 1 || m_oDIR.Value > 2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidDIRCodelist, "DIR codelist value should be between 1 & 2."));
                    return;
                }
                CodelistItem codelist = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG98F", "DIR")).PropertyInfo.CodeListInfo.GetCodelistItem((int)m_oDIR.Value);
                string dir = codelist.DisplayName;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                string actualRollMaterial = string.Empty, actualTop = string.Empty;
                if (rollMaterial < 1 || rollMaterial > 2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRollMaterialCodelist, "Rollmaterial codelist value should be between 1 & 2."));
                    return;
                }
                if (top < 1 || top > 3)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTopCodelist, "Top codelist value should be between 1 & 3."));
                    return;
                }
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                {
                    actualRollMaterial = metadataManager.GetCodelistInfo("Anvil_Variable_RollMat", "UDP").GetCodelistItem(rollMaterial).ShortDisplayName;
                    actualTop = metadataManager.GetCodelistInfo("Anvil_Variable_Top", "UDP").GetCodelistItem(top).ShortDisplayName;
                }
                else
                {
                    actualTop = "None";
                    actualRollMaterial = "Cast Iron";
                }
                if (pipeDiameter < 0.175 && actualRollMaterial == "Steel")
                    actualRollMaterial = "Cast Iron";
                double springRate = 0, loadMax = 0, loadMin = 0, coldload = 0, preset = 0, takeOut = 0;

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass anvilfig98FParts = (PartClass)catalogBaseHelper.GetPartClass("Anvil_SPRING_RATE");
                ReadOnlyCollection<BusinessObject> fig98FParts, fig98FSpringRollerStrokeParts;
                fig98FParts = anvilfig98FParts.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject part1 in fig98FParts)
                {
                    if (((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrAnvil_SPRING_RATE", "SIZE")).PropValue == m_oSIZE.Value))
                    {
                        springRate = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_SPRING_RATE", "RATE_" + "FIG98F".Substring(3, 2) + "")).PropValue;
                        loadMax = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_SPRING_RATE", "LOAD_MAX")).PropValue;
                        loadMin = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_SPRING_RATE", "LOAD_MIN")).PropValue;
                        break;
                    }
                }

                //Ports

                Port port1 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                coldload = (hotLoad - (workingTrav * springRate));

                if (dir == "Up")
                    coldload = (hotLoad + (workingTrav * springRate));

                preset = (coldload - loadMin) / springRate;

                if (HgrCompareDoubleService.cmpdbl(hotLoad, 0) == true && HgrCompareDoubleService.cmpdbl(workingTrav, 0) == true)
                {
                    coldload = 0;
                    preset = 0;
                }

                Double A = 0, M = 0, P = 0, R = 0, S = 0;
                PartClass anvilfig98FSpringRollerParts = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FSPRING_ROLLER");
                fig98FSpringRollerStrokeParts = anvilfig98FSpringRollerParts.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject part1 in fig98FSpringRollerStrokeParts)
                {
                    if (((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FSPRING_ROLLER", "PIPE_DIA")).PropValue > pipeDiameter - 0.001) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FSPRING_ROLLER", "PIPE_DIA")).PropValue < pipeDiameter + 0.001) && ((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrAnvil_FSPRING_ROLLER", "MATERIAL")).PropValue == actualRollMaterial))
                    {
                        A = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FSPRING_ROLLER", "A")).PropValue;
                        M = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FSPRING_ROLLER", "M")).PropValue;
                        P = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FSPRING_ROLLER", "P")).PropValue;
                        R = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FSPRING_ROLLER", "R")).PropValue;
                        S = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FSPRING_ROLLER", "S")).PropValue;
                        break;
                    }
                }
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, X + takeOut + preset), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (flangeDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFlangeDiameterGTZero, "Flange diameter should be greater than zero"));
                    return;
                }
                if (R <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRGTZero, "R value should be greater than zero"));
                    return;
                }
                if (S <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidSGTZero, "S value should be greater than zero"));
                    return;
                }
                if (M <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidMGTZero, "M value should be greater than zero"));
                    return;
                }
                if (C <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidCGTZero, "C value should be greater than zero"));
                    return;
                }
                if (E <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidEGTZero, "E value should be greater than zero"));
                    return;
                }
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTGTZero, "T value should be greater than zero"));
                    return;
                }
                if (colDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidColDiameterGTZero, "Col diameter should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(flangeT , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFlangeTNZero, "FlangeT value cannot be zero"));
                    return;
                }
                if (A <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidAGTZero, "A value should be greater than zero"));
                    return;
                }
                if (A <= 0 && pipeDiameter <= 0 && P <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidAPipeDiaPGTZero, "A value,PipeDiameter and P value should be greater than zero"));
                    return;
                }
                if (actualTop == "Load Flange")
                {
                    takeOut = flangeT;

                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                    Projection3d loadFlange = symbolGeometryHelper.CreateCylinder(null, flangeDiameter / 2, flangeT);
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    loadFlange.Transform(matrix);
                    m_Symbolic.Outputs["LOAD_FLANGE"] = loadFlange;
                }

                if (actualTop == "Pipe Roll")
                {
                    takeOut = A;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-R / 2, -S / 2, takeOut - M);
                    Projection3d pipeRoll = symbolGeometryHelper.CreateBox(null, R, S, M, 9);
                    matrix = new Matrix4X4();
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    pipeRoll.Transform(matrix);
                    m_Symbolic.Outputs["PIPE_ROLL"] = pipeRoll;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                    Projection3d roll1 = symbolGeometryHelper.CreateCylinder(null, A - pipeDiameter / 2 - P, R);
                    matrix = new Matrix4X4();
                    matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0));
                    matrix.Translate(new Vector(-R / 2, -(A - pipeDiameter / 2 - P) * 1.25, takeOut - P));
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    roll1.Transform(matrix);
                    m_Symbolic.Outputs["ROLL_1"] = roll1;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                    Projection3d roll2 = symbolGeometryHelper.CreateCylinder(null, A - pipeDiameter / 2 - P, R);
                    matrix = new Matrix4X4();
                    matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0));
                    matrix.Translate(new Vector(-R / 2, (A - pipeDiameter / 2 - P) * 1.25, takeOut - P));
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    roll2.Transform(matrix);
                    m_Symbolic.Outputs["ROLL_2"] = roll2;
                }
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, X - B + takeOut + preset);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, C / 2, B - T);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-E / 2, -E / 2, X + takeOut - T + preset);
                Projection3d bottom = symbolGeometryHelper.CreateBox(null, E, E, T, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                bottom.Transform(matrix);
                m_Symbolic.Outputs["BOTTOM"] = bottom;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, takeOut);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d loadColumn = symbolGeometryHelper.CreateCylinder(null, colDiameter / 2, X - B + (B * 0.75));
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                loadColumn.Transform(matrix);
                m_Symbolic.Outputs["LOAD_COLUMN"] = loadColumn;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG98F"));
                return;
            }
        }
        #endregion

    }

}
