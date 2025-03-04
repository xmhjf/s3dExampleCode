//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG201.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG201
//   Author       :  Rajeswari
//   Creation Date:  09-May-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   09-May-2013  Rajeswari  CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
//   11-Dec-2014     PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//   05-May-2016     PVK	 TR-CP-293853	Copy/Pasting Cable Tray Deprecated supports results record exceptions
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

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
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class Anvil_FIG201 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG201"
        //----------------------------------------------------------------------------------
        const double const_1 = 0.2032;
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "RESERVOIR_TYP", "RESERVOIR_TYP", 1)]
        public InputDouble m_oRESERVOIR_TYP;
        [InputString(3, "OPT", "OPT", "No Value")]
        public InputString m_oOPT;
        [InputDouble(4, "STROKE", "STROKE", 0.999999)]
        public InputDouble m_dSTROKE;
        [InputDouble(5, "PISTON_SETTING", "PISTON_SETTING", 0.999999)]
        public InputDouble m_dPISTON_SETTING;
        [InputDouble(6, "ROTATION", "ROTATION", 0.999999)]
        public InputDouble m_dROTATION;
        [InputDouble(7, "Length", "Length", 0.999999)]
        public InputDouble m_dLength;
        [InputDouble(8, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputString(9, "SIZE", "SIZE", "No Value")]
        public InputString m_oSIZE;
        [InputDouble(10, "Z", "Z", 0.999999)]
        public InputDouble m_dZ;
        [InputDouble(11, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(12, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(13, "Q", "Q", 0.999999)]
        public InputDouble m_dQ;
        [InputDouble(14, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(15, "D1", "D1", 0.999999)]
        public InputDouble m_dD1;
        [InputDouble(16, "N", "N", 0.999999)]
        public InputDouble m_dN;
        [InputDouble(17, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(18, "R", "R", 0.999999)]
        public InputDouble m_dR;
        [InputDouble(19, "S", "S", 0.999999)]
        public InputDouble m_dS;
        [InputDouble(20, "SIZE_INDEX", "SIZE_INDEX", 0.999999)]
        public InputDouble m_dSIZE_INDEX;
        [InputDouble(21, "I_TRANS", "I_TRANS", 0.999999)]
        public InputDouble m_dI_TRANS;
        [InputDouble(22, "I_PRESS", "I_PRESS", 0.999999)]
        public InputDouble m_dI_PRESS;
        [InputDouble(23, "I_METAL", "I_METAL", 0.999999)]
        public InputDouble m_dI_METAL;
        [InputDouble(24, "K", "K", 0.999999)]
        public InputDouble m_dK;
        [InputDouble(25, "VALVE_TYP", "VALVE_TYP", 1)]
        public InputDouble m_oVALVE_TYP;
        [InputDouble(26, "RES_ORIENTATION", "RES_ORIENTATION", 1)]
        public InputDouble m_oRES_ORIENTATION;
        [InputDouble(27, "HOT_PISTON_SETTING", "HOT_PISTON_SETTING", 0.999999)]
        public InputDouble m_dHOT_PISTON_SETTING;
        [InputDouble(28, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("LUG2", "LUG2")]
        [SymbolOutput("LUG2", "LUG2")]
        [SymbolOutput("LUG2", "LUG2")]
        [SymbolOutput("LUG", "LUG")]
        [SymbolOutput("EXTENSION", "EXTENSION")]
        [SymbolOutput("VALVE", "VALVE")]
        [SymbolOutput("CYL", "CYL")]
        [SymbolOutput("RIGHT", "RIGHT")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("END", "END")]
        [SymbolOutput("BOLT", "BOLT")]
        [SymbolOutput("PORT1", "PORT1")]
        [SymbolOutput("PORT2", "PORT2")]
        [SymbolOutput("OPTIONALPART1", "OPTIONALPART1")]
        [SymbolOutput("OPTIONALPART2", "OPTIONALPART2")]
        [SymbolOutput("OPTIONALPART3", "OPTIONALPART3")]
        [SymbolOutput("OPTIONALPART4", "OPTIONALPART4")]
        [SymbolOutput("OPTIONALPART5", "OPTIONALPART5")]
        [SymbolOutput("OPTIONALPART6", "OPTIONALPART6")]
        [SymbolOutput("OPTIONALPART7", "OPTIONALPART7")]
        [SymbolOutput("OPTIONALPART8", "OPTIONALPART8")]
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

                Double stroke = m_dSTROKE.Value;
                Double pistonSetting = m_dPISTON_SETTING.Value;
                Double rotation;
                Double length = m_dLength.Value;
                Double L = m_dL.Value;
                Double Z = m_dZ.Value;
                Double sizeIndex = m_dSIZE_INDEX.Value;
                Double iTrans = m_dI_TRANS.Value;
                Double iPress = m_dI_PRESS.Value;
                Double hotPisttonSetting = m_dHOT_PISTON_SETTING.Value;
                int reservoirType = (int)m_oRESERVOIR_TYP.Value;
                String size = m_oSIZE.Value;
                String opt = m_oOPT.Value;
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                String actualReservoirtype;
                Double actualStroke, actualPipeDiameter = 0, cMin = 0, C, W = 0;
                rotation = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Angle, m_dROTATION.Value, UnitName.ANGLE_DEGREE);

                //Validating Inputs
                if (reservoirType < 1 || reservoirType > 4)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidReservoirTypeCodelist, "ReservoirType codelist value should be between 1 & 4."));
                    return;
                }
                if (stroke < 1 || stroke > 4)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidStrokeCodelist, "Stroke codelist value should be between 1 & 4."));
                    return;
                }

                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                {
                    actualReservoirtype = metadataManager.GetCodelistInfo("Anvil_Dyn_ResType", "UDP").GetCodelistItem(reservoirType).ShortDisplayName.Trim();
                    actualStroke = Double.Parse(metadataManager.GetCodelistInfo("Anvil_Dyn_Stroke", "UDP").GetCodelistItem((int)stroke).ShortDisplayName) * 25.4 / 1000;
                }
                else
                {
                    actualReservoirtype = "Transparent";
                    actualStroke = 5 * 25.4 / 1000;
                }

                if (actualStroke > 0.254 && size == "1 1/2")
                    actualStroke = 0.127;

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

                PartClass anvilfig201StrokeParts = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG201_STROKE");
                ReadOnlyCollection<BusinessObject> fig201StrokeParts = anvilfig201StrokeParts.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject part1 in fig201StrokeParts)
                {
                    if (((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_STROKE", "SIZE")).PropValue == size) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_STROKE", "STROKE")).PropValue > actualStroke - 0.001) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_STROKE", "STROKE")).PropValue < (actualStroke + 0.001)))
                    {
                        cMin = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_STROKE", "CMIN")).PropValue;
                        break;
                    }
                }
                C = pistonSetting + cMin - Z;
                Double A = m_dA.Value;
                Double F = m_dF.Value;
                Double Q = m_dQ.Value;
                Double D = m_dD.Value;
                Double D1 = m_dD1.Value;
                Double iMetal = m_dI_METAL.Value;
                Double N = m_dN.Value;
                Double T = m_dT.Value;
                Double R = m_dR.Value;
                Double S = m_dS.Value;
                Double K = m_dK.Value;
                Double E = -Q;

                if (S <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidSGTZero, "S value should be greater than zero"));
                    return;
                }
                if (R <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRGTZero, "R value should be greater than zero"));
                    return;
                }
                if (K < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidKLTZero, "K value should not be less than zero"));
                    return;
                }
                if (iMetal < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidIMetalLTZero, "IMetal value should not be less than zero"));
                    return;
                }
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTGTZero, "T value should be greater than zero"));
                    return;
                }
                if (N <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidNGTZero, "N value should be greater than zero"));
                    return;
                }
                if (A <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidAGTZero, "A value should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }
                if (D1 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidD1GTZero, "D1 value should be greater than zero"));
                    return;
                }
                if (L <= 0 && K <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidLandKGTZero, "L and K values should be greater than zero"));
                    return;
                }
                if (L <= 0 && iMetal <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidLandIMetalGTZero, "L and IMetal values should be greater than zero"));
                    return;
                }
                if (W == 0 && F == 0 && R == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidWandFandRNZero, "W, F and R values cannot be zero"));
                    return;
                }

                if (opt == "1")
                {
                    E = 0;
                    symbolGeometryHelper.ActivePosition = new Position(-S / 4, -R, E - R);
                    Projection3d lug2 = symbolGeometryHelper.CreateBox(null, S / 2, R * 2, Q + R, 9);
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    lug2.Transform(matrix);
                    m_Symbolic.Outputs["LUG2"] = lug2;
                }
                if (opt == "2")
                {
                    E = A;
                    symbolGeometryHelper.ActivePosition = new Position(-S / 4, -R, E - R);
                    Projection3d lug2 = symbolGeometryHelper.CreateBox(null, S / 2, R * 2, Q + R, 9);
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    lug2.Transform(matrix);
                    m_Symbolic.Outputs["LUG2"] = lug2;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-D1 / 2, -D / 2, 0);
                    Projection3d optionalPart1 = symbolGeometryHelper.CreateBox(null, N, D, A * 1.5, 9);
                    matrix = new Matrix4X4();
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    optionalPart1.Transform(matrix);
                    m_Symbolic.Outputs["OPTIONALPART1"] = optionalPart1;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(D1 / 2 - N, -D / 2, 0);
                    Projection3d optionalPart2 = symbolGeometryHelper.CreateBox(null, N, D, A * 1.5, 9);
                    matrix = new Matrix4X4();
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    optionalPart2.Transform(matrix);
                    m_Symbolic.Outputs["OPTIONALPART2"] = optionalPart2;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-D1 / 2 + N, -D / 2, 0);
                    Projection3d optionalPart3 = symbolGeometryHelper.CreateBox(null, D1 - 2 * N, D, N, 9);
                    matrix = new Matrix4X4();
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    optionalPart3.Transform(matrix);
                    m_Symbolic.Outputs["OPTIONALPART3"] = optionalPart3;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                    Projection3d optionalPart4 = symbolGeometryHelper.CreateCylinder(null, T / 2, D1 + T);
                    matrix = new Matrix4X4();
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 1, 0));
                    matrix.Translate(new Vector(D1 / 2 + T / 2, 0, A));
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    optionalPart4.Transform(matrix);
                    m_Symbolic.Outputs["OPTIONALPART4"] = optionalPart4;
                }

                if (size == "8")
                {
                    if (opt == "3")
                        sizeIndex = 6;
                    else
                        sizeIndex = 8;
                }

                if (actualStroke > 0.254 && actualReservoirtype == "Pressurized")
                    iMetal = 0;

                if (actualReservoirtype == "Transparent")
                    iMetal = iTrans;

                if (actualReservoirtype == "Pressurized" && actualStroke < 15)
                    iMetal = iPress;

                if (actualReservoirtype == "Remote" && actualStroke < 15)
                    iMetal = L;

                if (opt == "3")
                {
                    // Get the pipe dia (it will be in meters)
                    try
                    {
                        RelationCollection hgrRelation = Occurrence.GetRelationship("SupportHasComponents", "Support");
                        BusinessObject businessObject = hgrRelation.TargetObjects[0];
                        SupportedHelper supportedHelper = new SupportedHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                        SupportHelper supportHelper = new SupportHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                        PipeObjectInfo pipeInfo = null;
                        if (supportHelper.SupportedObjects.Count > 0 && supportedHelper.SupportedObjectInfo(1) != null)
                        {
                            pipeInfo = ((PipeObjectInfo)supportedHelper.SupportedObjectInfo(1));
                        }
                        if (pipeInfo != null)
                        {
                            if (pipeInfo.NominalDiameter.Units == "in")
                                actualPipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, pipeInfo.NominalDiameter.Size, UnitName.DISTANCE_INCH);
                            else if (pipeInfo.NominalDiameter.Units == "mm")
                                actualPipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, pipeInfo.NominalDiameter.Size, UnitName.DISTANCE_MILLIMETER);
                        }
                        else
                        {
                            actualPipeDiameter = const_1;
                        }
                    }
                    catch
                    {
                        actualPipeDiameter = const_1;
                    }

                    PartClass anvil_fig201Clamp2Parts = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG201_CLAMP1");
                    ReadOnlyCollection<BusinessObject> fig201Clamp2Parts = anvil_fig201Clamp2Parts.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    foreach (BusinessObject part1 in fig201Clamp2Parts)
                    {
                        if (((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_CLAMP1", "PIPE_DIA2")).PropValue > actualPipeDiameter - 0.001) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_CLAMP1", "PIPE_DIA2")).PropValue < (actualPipeDiameter + 0.001)))
                        {
                            E = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_CLAMP1", "E_" + sizeIndex.ToString().Trim())).PropValue;
                            break;
                        }
                    }

                    if (E < 0)
                        E = 0.02;

                    symbolGeometryHelper.ActivePosition = new Position(-S / 4, -R, E - R);
                    Projection3d lug2 = symbolGeometryHelper.CreateBox(null, S / 2, R * 2, Q + R, 9);
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    lug2.Transform(matrix);
                    m_Symbolic.Outputs["LUG2"] = lug2;

                    if ((1.5 * T + E) >= (actualPipeDiameter / 2))
                    {
                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.ActivePosition = new Position(-(S / 2 + 0.4 * T), -1.5 * T, actualPipeDiameter / 2);
                        Projection3d optionalPart1 = symbolGeometryHelper.CreateBox(null, 0.4 * T, 3 * T, 1.5 * T + E - actualPipeDiameter / 2, 9);
                        matrix = new Matrix4X4();
                        matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                        optionalPart1.Transform(matrix);
                        m_Symbolic.Outputs["OPTIONALPART1"] = optionalPart1;

                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.ActivePosition = new Position(S / 2, -1.5 * T, actualPipeDiameter / 2);
                        Projection3d optionalPart2 = symbolGeometryHelper.CreateBox(null, 0.4 * T, 3 * T, 1.5 * T + E - actualPipeDiameter / 2, 9);
                        matrix = new Matrix4X4();
                        matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                        optionalPart2.Transform(matrix);
                        m_Symbolic.Outputs["OPTIONALPART2"] = optionalPart2;
                    }
                    else
                    {
                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.ActivePosition = new Position(-(S / 2 + 0.4 * T), -1.5 * T, actualPipeDiameter - (1.5 * T + E));
                        Projection3d optionalPart1 = symbolGeometryHelper.CreateBox(null, 0.4 * T, 3 * T, actualPipeDiameter / 2 - (1.5 * T + E), 9);
                        matrix = new Matrix4X4();
                        matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                        optionalPart1.Transform(matrix);
                        m_Symbolic.Outputs["OPTIONALPART1"] = optionalPart1;

                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.ActivePosition = new Position(S / 2, -1.5 * T, actualPipeDiameter - (1.5 * T + E));
                        Projection3d topRight = symbolGeometryHelper.CreateBox(null, 0.4 * T, 3 * T, actualPipeDiameter / 2 - (1.5 * T + E), 9);
                        matrix = new Matrix4X4();
                        matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                        topRight.Transform(matrix);
                        m_Symbolic.Outputs["OPTIONALPART2"] = topRight;
                    }

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-(S / 2 + 0.4 * T), -1.5 * T, -actualPipeDiameter / 2 - 3 * T);
                    Projection3d optionalPart3 = symbolGeometryHelper.CreateBox(null, 0.4 * T, 3 * T, 3 * T, 9);
                    matrix = new Matrix4X4();
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    optionalPart3.Transform(matrix);
                    m_Symbolic.Outputs["OPTIONALPART3"] = optionalPart3;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(S / 2, -1.5 * T, -actualPipeDiameter / 2 - 3 * T);
                    Projection3d optionalPart4 = symbolGeometryHelper.CreateBox(null, 0.4 * T, 3 * T, 3 * T, 9);
                    matrix = new Matrix4X4();
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    optionalPart4.Transform(matrix);
                    m_Symbolic.Outputs["OPTIONALPART4"] = optionalPart4;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                    Projection3d optionalPart5 = symbolGeometryHelper.CreateCylinder(null, actualPipeDiameter / 2 + 0.4 * T, 3 * T);
                    matrix = new Matrix4X4();
                    matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                    matrix.Translate(new Vector(0, -1.5 * T, 0));
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    optionalPart5.Transform(matrix);
                    m_Symbolic.Outputs["OPTIONALPART5"] = optionalPart5;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    Vector normal = new Position(0, -(S / 2 + T), E).Subtract(new Position(0, S / 2 + T, E));
                    symbolGeometryHelper.ActivePosition = new Position(0, S / 2 + T, E);
                    symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                    Projection3d optionalPart6 = symbolGeometryHelper.CreateCylinder(null, T / 2, normal.Length);
                    m_Symbolic.Outputs["OPTIONALPART6"] = optionalPart6;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    normal = new Position(0, -(S / 2 + T), actualPipeDiameter / 2 + 1.5 * T).Subtract(new Position(0, S / 2 + T, actualPipeDiameter / 2 + 1.5 * T));
                    symbolGeometryHelper.ActivePosition = new Position(0, S / 2 + T, actualPipeDiameter / 2 + 1.5 * T);
                    symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                    Projection3d optionalPart7 = symbolGeometryHelper.CreateCylinder(null, T / 2, normal.Length);
                    m_Symbolic.Outputs["OPTIONALPART7"] = optionalPart7;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    normal = new Position(0, -(S / 2 + T), -actualPipeDiameter / 2 - 1.5 * T).Subtract(new Position(0, S / 2 + T, -actualPipeDiameter / 2 - 1.5 * T));
                    symbolGeometryHelper.ActivePosition = new Position(0, S / 2 + T, -actualPipeDiameter / 2 - 1.5 * T);
                    symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                    Projection3d optionalPart8 = symbolGeometryHelper.CreateCylinder(null, T / 2, normal.Length);
                    m_Symbolic.Outputs["OPTIONALPART8"] = optionalPart8;
                }
                W = length - A - F - C - E;

                if (opt == "0")
                {
                    Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["PORT1"] = port1;
                }
                if (opt == "1")
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["PORT1"] = port1;
                }
                if (opt == "2")
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["PORT1"] = port1;
                }
                if (opt == "3")
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["PORT1"] = port1;
                }
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["PORT2"] = port2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-D1 / 2, -D / 2, 0);
                Projection3d right1 = symbolGeometryHelper.CreateBox(null, N, D, A * 1.5, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(rotation * Math.PI / 180, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                matrix.Translate(new Vector(0, 0, length));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                right1.Transform(matrix);
                m_Symbolic.Outputs["RIGHT"] = right1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(D1 / 2 - N, -D / 2, 0);
                Projection3d left1 = symbolGeometryHelper.CreateBox(null, N, D, A * 1.5, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(rotation * Math.PI / 180, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                matrix.Translate(new Vector(0, 0, length));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                left1.Transform(matrix);
                m_Symbolic.Outputs["LEFT"] = left1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-D1 / 2 + N, -D / 2, 0);
                Projection3d end1 = symbolGeometryHelper.CreateBox(null, D1 - 2 * N, D, N, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(rotation * Math.PI / 180, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI, new Vector(0, 1, 0));
                matrix.Translate(new Vector(0, 0, length));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                end1.Transform(matrix);
                m_Symbolic.Outputs["END"] = end1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d bolt1 = symbolGeometryHelper.CreateCylinder(null, T / 2, D1 + T);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(D1 / 2 + T / 2, 0, A));
                matrix.Rotate(rotation * Math.PI / 180, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                matrix.Translate(new Vector(0, 0, length));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                bolt1.Transform(matrix);
                m_Symbolic.Outputs["BOLT"] = bolt1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d lug = symbolGeometryHelper.CreateCylinder(null, R, S / 2);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(-S / 4, 0, E + C + W + F));
                matrix.Rotate(-rotation * Math.PI / 180, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                lug.Transform(matrix);
                m_Symbolic.Outputs["LUG"] = lug;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d extension = symbolGeometryHelper.CreateCylinder(null, R, W + F - R);
                matrix = new Matrix4X4();
                matrix.Rotate(rotation * Math.PI / 180, new Vector(0, 0, 1));
                matrix.Translate(new Vector(0, 0, E + C));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                extension.Transform(matrix);
                m_Symbolic.Outputs["EXTENSION"] = extension;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-L, -L, E + Q + pistonSetting);
                Projection3d valve = symbolGeometryHelper.CreateBox(null, L + K, L + iMetal, cMin - Q - Z, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(-rotation * Math.PI / 180, new Vector(0, 0, 1));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                valve.Transform(matrix);
                m_Symbolic.Outputs["VALVE"] = valve;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                double fraction = FeettInchFractiontoMeter("0" + size) / 2;
                if (base.ToDoListMessage != null)
                {
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        return;
                    }
                }
                Projection3d cylinder = symbolGeometryHelper.CreateCylinder(null, MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, fraction, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER), pistonSetting);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(0, 0, E + Q));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                cylinder.Transform(matrix);
                m_Symbolic.Outputs["CYL"] = cylinder;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG201"));
                return;
            }
        }

        /// <summary>
        /// Converts fraction value to double value .
        /// </summary>
        /// <param name="fraction">fraction value-String.</param>
        /// <example>
        /// double c=FeettInchFractiontoMeter(size);
        /// </example>
        double FeettInchFractiontoMeter(string fraction)
        {
            double result;

            if (double.TryParse(fraction, out result))
                return result;

            string[] split = fraction.Split(new char[] { ' ', '/' });
            int a, b, c, d;

            if (split.Length == 2 || split.Length == 3)
            {
                if (int.TryParse(split[0], out a) && int.TryParse(split[1], out b))
                {
                    if (split.Length == 2)
                        return (double)a / b;
                    if (int.TryParse(split[2], out c))
                        return a + (double)b / c;
                }
            }
            else if (split.Length == 4)
            {

                if (int.TryParse(split[0], out a) && int.TryParse(split[1], out b) && int.TryParse(split[2], out c))
                {
                    if (int.TryParse(split[3], out d))
                        return (a * 304.8 + (b + (double)c / d) * 25.4) / 25.4;
                }
            }
            ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFraction, "The Input Fraction is not in Correct Format"));
            return 0;
        }

        #endregion

        #region "ICustomHgrBOMDescription Members"
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                // To get SIZE_INDEX
                string sizeIndex, ndUnitType;
                Double length, A, F, E = 0;
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                sizeIndex = Convert.ToString((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrAnvil_FIG201", "SIZE_INDEX")).PropValue);
                ndUnitType = (string)((PropertyValueString)part.GetPropertyValue("IJHgrDiameterSelection", "NDUnitType")).PropValue;
                length = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                A = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrAnvil_FIG201", "A")).PropValue;
                F = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrAnvil_FIG201", "F")).PropValue;

                // To get FINISH
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG201", "FINISH");
                if (finishCodelist.PropValue == 0)
                    finishCodelist.PropValue = 1;
                if (finishCodelist.PropValue < 1 && finishCodelist.PropValue > 2)
                {
                    ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFinishvalue, "Finish value should be 1 or 2"));
                    return "";
                }
                string finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;

                // To get RESERVOIR_TYP
                PropertyValueCodelist reservoirTypeCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG201", "RESERVOIR_TYP");
                if (reservoirTypeCodelist.PropValue < 1 || reservoirTypeCodelist.PropValue > 4)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidReservoirTypeCodelist, "ReservoirType codelist value should be between 1 & 4."));
                    return "";
                }
                string actualReservoirType = reservoirTypeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(reservoirTypeCodelist.PropValue).DisplayName;

                // To get OPT
                string opt = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrAnvil_FIG201", "OPT")).PropValue;

                // To get ACTUAL_STROKE
                PropertyValueCodelist actualStrokeCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG201", "STROKE");
                if (actualStrokeCodelist.PropValue < 1 || actualStrokeCodelist.PropValue > 4)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidStrokeCodelist, "Stroke codelist value should be between 1 & 4."));
                    return "";
                }
                string actualStroke = actualStrokeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(actualStrokeCodelist.PropValue).DisplayName;
                double actualStrokeValue = Convert.ToDouble(actualStroke) * 25.4 / 1000;

                // To get PISTON_SETTING
                double pistonSettingValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG201", "PISTON_SETTING")).PropValue;
                string pistonSetting = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, pistonSettingValue, UnitName.DISTANCE_INCH);

                // To get HOT_PISTON_SETTING
                double hotPistonSettingValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG201", "HOT_PISTON_SETTING")).PropValue;
                string hotPistonSetting = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, hotPistonSettingValue, UnitName.DISTANCE_INCH);

                // To get VALVE_TYP
                PropertyValueCodelist valveTypeStrokeCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG201", "VALVE_TYP");
                if (valveTypeStrokeCodelist.PropValue < 1 || valveTypeStrokeCodelist.PropValue > 2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidValveTypeCodelist, "ValveType codelist value should be between 1 & 2."));
                    return "";
                }
                string actualValveType = valveTypeStrokeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(valveTypeStrokeCodelist.PropValue).DisplayName;

                // To get RES_ORIENTATION
                PropertyValueCodelist resOrientationCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG201", "RES_ORIENTATION");
                if (resOrientationCodelist.PropValue < 1 || resOrientationCodelist.PropValue > 2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidResOrientationCodelist, "ResOrientation codelist value should be between 1 & 2."));
                    return "";
                }
                string actualresOrientation = resOrientationCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(resOrientationCodelist.PropValue).DisplayName;

                // To get SIZE
                string size = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrSize", "SIZE")).PropValue;

                // To get Z
                double Z = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrAnvil_FIG201", "Z")).PropValue;

                double cMin = 0, CValue, WValue = 0;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass anvil_fig201StrokeParts = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG201_STROKE");
                ReadOnlyCollection<BusinessObject> fig201StrokeParts = anvil_fig201StrokeParts.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject part1 in fig201StrokeParts)
                {
                    if (((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_STROKE", "SIZE")).PropValue == size) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_STROKE", "STROKE")).PropValue > actualStrokeValue - 0.001) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_STROKE", "STROKE")).PropValue < (actualStrokeValue + 0.001)))
                    {
                        cMin = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_STROKE", "CMIN")).PropValue;
                        break;
                    }
                }

                CValue = pistonSettingValue + cMin - Z;
                string C = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, pistonSettingValue, UnitName.DISTANCE_INCH);

                if (opt == "1")
                    E = 0;
                if (opt == "2")
                    E = A;

                WValue = length - A - F - CValue - E;
                string W = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, pistonSettingValue, UnitName.DISTANCE_INCH);

                string valveBom="A", reservoirBom ="M", resOrientationBom="O", bomPositions, bomEnd;

                bomPositions = ", Hot Set=" + hotPistonSetting + ", Cold Set=" + pistonSetting;
                bomEnd = ", W=" + W + ", Finish=" + finish + bomPositions + ", C=" + C;

                if (actualValveType == "Temperature Compensating")
                    valveBom = "T";
                if (HgrCompareDoubleService.cmpdbl(actualStrokeValue, 0.127) == true)
                    actualStroke = "05";
                if (actualReservoirType == "Pressurized")
                    reservoirBom = "P";
                if (actualReservoirType == "Transparent")
                    reservoirBom = "L";
                if (actualReservoirType == "Remote")
                    reservoirBom = "R";

                if (opt == "3")
                {
                    // Get the pipe dia (it will be in meters)
                    double actualPipeDiameter = 0;
                    try
                    {
                        RelationCollection hgrRelation = oSupportOrComponent.GetRelationship("SupportHasComponents", "Support");
                        BusinessObject businessObject = hgrRelation.TargetObjects[0];
                        SupportedHelper supportedHelper = new SupportedHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                        SupportHelper supportHelper = new SupportHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                        PipeObjectInfo pipeInfo = null;
                        if (supportHelper.SupportedObjects.Count > 0 && supportedHelper.SupportedObjectInfo(1) != null)
                        {
                            pipeInfo = ((PipeObjectInfo)supportedHelper.SupportedObjectInfo(1));
                        }
                        if (pipeInfo != null)
                        {
                            if (pipeInfo.NominalDiameter.Units == "in")
                                actualPipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, pipeInfo.NominalDiameter.Size, UnitName.DISTANCE_INCH);
                            else if (pipeInfo.NominalDiameter.Units == "mm")
                                actualPipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, pipeInfo.NominalDiameter.Size, UnitName.DISTANCE_MILLIMETER);
                        }
                        else
                        {
                            actualPipeDiameter = const_1;
                        }
                    }
                    catch
                    {
                        actualPipeDiameter = const_1;
                    }


                    string pipeNominalDiameter = null;
                    PartClass anvilPNDParts = (PartClass)catalogBaseHelper.GetPartClass("Anvil_PND");
                    ReadOnlyCollection<BusinessObject> PNDParts = anvilPNDParts.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    foreach (BusinessObject part1 in PNDParts)
                    {
                        if (((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_PND", "Pipe_Dia_M")).PropValue > actualPipeDiameter - 0.001) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_PND", "Pipe_Dia_M")).PropValue < (actualPipeDiameter + 0.001)))
                        {
                            pipeNominalDiameter = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrAnvil_PND", "Pipe_Nom_Dia")).PropValue;
                            break;
                        }
                    }

                    bomEnd = ", W=" + W + ", Pipe Size=" + pipeNominalDiameter + " " + ndUnitType + ", Finish=" + finish + bomPositions + ", C=" + C;
                }
                if (actualresOrientation == "Rod up or horizontal" && actualReservoirType != "Pressurized" && actualReservoirType != "Remote")
                    resOrientationBom = "U";
                if (actualresOrientation == "Rod down" && actualReservoirType != "Pressurized" && actualReservoirType != "Remote")
                    resOrientationBom = "D";

                bomDescription = "Anvil FIG201 Size " + size + " Hydraulic Snubber with Extension Piece (Option " + opt + "), Part No: 201" + sizeIndex + actualStroke + opt + valveBom + reservoirBom + resOrientationBom + bomEnd;

            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of Anvil_FIG201"));
            }
            return bomDescription;
        }
        #endregion


        #region "ICustomHgrWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double weight, cogX, cogY, cogZ, weightPerUnitLength = 0;
                double length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                string size = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJUAHgrSize", "SIZE")).PropValue;
                PropertyValueCodelist actualStrokeCodelist = (PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrAnvil_FIG201", "STROKE");
                CodelistItem codeList = actualStrokeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(actualStrokeCodelist.PropValue);
                double actualStroke = Double.Parse(codeList.ShortDisplayName.Trim()) * 25.4 / 1000;

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass anvil_fig201StrokeParts = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG201_STROKE");
                ReadOnlyCollection<BusinessObject> fig201StrokeParts = anvil_fig201StrokeParts.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject part1 in fig201StrokeParts)
                {
                    if (((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_STROKE", "SIZE")).PropValue == size) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_STROKE", "STROKE")).PropValue > actualStroke - 0.001) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_STROKE", "STROKE")).PropValue < (actualStroke + 0.001)))
                    {
                        weightPerUnitLength = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG201_STROKE", "WEIGHT")).PropValue;
                        break;
                    }
                }
                weight = weightPerUnitLength * length;
                cogX = 0;
                cogY = 0;
                cogZ = -length / 2;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrWeightCG, "Error while defining weightCG of Anvil_FIG201"));
            }
        }
        #endregion

    }

}
