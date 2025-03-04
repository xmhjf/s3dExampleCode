//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG211.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG211
//   Author       :  Rajeswari
//   Creation Date:  09-May-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   09-May-2013  Rajeswari  CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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
    public class Anvil_FIG211 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG211"
        //----------------------------------------------------------------------------------
        const double const_1 = 0.2032;
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "CONFIG", "CONFIG", "No Value")]
        public InputString m_oCONFIG;
        [InputDouble(3, "X", "X", 0.999999)]
        public InputDouble m_dX;
        [InputDouble(4, "Y", "Y", 0.999999)]
        public InputDouble m_dY;
        [InputDouble(5, "Length", "Length", 0.999999)]
        public InputDouble m_dLength;
        [InputDouble(6, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputString(7, "SIZE", "SIZE", "No Value")]
        public InputString m_oSIZE;
        [InputDouble(8, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(9, "EXT_DIA", "EXT_DIA", 0.999999)]
        public InputDouble m_dEXT_DIA;
        [InputDouble(10, "ROD_END", "ROD_END", 0.999999)]
        public InputDouble m_dROD_END;
        [InputDouble(11, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(12, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(13, "D1", "D1", 0.999999)]
        public InputDouble m_dD1;
        [InputDouble(14, "N", "N", 0.999999)]
        public InputDouble m_dN;
        [InputDouble(15, "R", "R", 0.999999)]
        public InputDouble m_dR;
        [InputDouble(16, "S", "S", 0.999999)]
        public InputDouble m_dS;
        [InputDouble(17, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("TOP_EYE", "TOP_EYE")]
        [SymbolOutput("BOT_EYE", "BOT_EYE")]
        [SymbolOutput("TOP_END", "TOP_END")]
        [SymbolOutput("BOT_END", "BOT_END")]
        [SymbolOutput("EXT_PIECE", "EXT_PIECE")]
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

                Double length = m_dLength.Value;
                Double T = m_dT.Value;
                Double A = m_dA.Value;
                Double extDiameter = m_dEXT_DIA.Value;
                Double rodEnd = m_dROD_END.Value;
                Double F = m_dF.Value;
                Double D = m_dD.Value;
                Double D1 = m_dD1.Value;
                Double N = m_dN.Value;
                Double R = m_dR.Value;
                Double S = m_dS.Value;
                String config = m_oCONFIG.Value;
                String size = m_oSIZE.Value;
                Double X = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Angle, m_dX.Value, UnitName.ANGLE_DEGREE);
                Double Y = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Angle, m_dY.Value, UnitName.ANGLE_DEGREE);

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                double takeOff, actualPipeDiameter = 0, E = 0, CC = 0;

                takeOff = A / Math.Cos(Y * Math.PI / 180);

                if (Math.Abs(Y) > 60)
                    takeOff = 0;

                //Validating Inputs
                if (R <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRGTZero, "R value should be greater than zero"));
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
                if (rodEnd <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRodEndGTZero, "Rod end value should be greater than zero"));
                    return;
                }
                if (extDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidExtDiameterGTZero, "Ext diameter should be greater than zero"));
                    return;
                }

                if (config == "3")
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

                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    PartClass anvilfig211E1Parts = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG211_E1");
                    ReadOnlyCollection<BusinessObject> fig211E1Parts = anvilfig211E1Parts.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    foreach (BusinessObject part1 in fig211E1Parts)
                    {
                        if (((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG211_E1", "PIPE_DIA2")).PropValue > actualPipeDiameter - 0.001) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG211_E1", "PIPE_DIA2")).PropValue < (actualPipeDiameter + 0.001)))
                        {
                            E = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG211_E1", "SIZE_" + size)).PropValue;
                            break;
                        }
                    }

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
                        Projection3d optionalPart2 = symbolGeometryHelper.CreateBox(null, 0.4 * T, 3 * T, actualPipeDiameter / 2 - (1.5 * T + E), 9);
                        matrix = new Matrix4X4();
                        matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                        optionalPart2.Transform(matrix);
                        m_Symbolic.Outputs["OPTIONALPART2"] = optionalPart2;
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
                CC = length - (takeOff + E);

                if (config == "1")
                    CC = length - takeOff;

                if (config == "2")
                {
                    CC = length - A - takeOff;
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

                if (config == "1")
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["PORT1"] = port1;
                }
                if (config == "2")
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["PORT1"] = port1;
                }
                if (config == "3")
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["PORT1"] = port1;
                }
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, length + A * Math.Cos(Y * Math.PI / 180) - takeOff), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["PORT2"] = port2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-D1 / 2, -D / 2, 0);
                Projection3d right1 = symbolGeometryHelper.CreateBox(null, N, D, A * 1.5, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI + Y * Math.PI / 180, new Vector(0, 1, 0));
                matrix.Translate(new Vector(A * Math.Sin(Y * Math.PI / 180), 0, length + A * Math.Cos(Y * Math.PI / 180) - takeOff));
                matrix.Rotate(X * Math.PI / 180, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                right1.Transform(matrix);
                m_Symbolic.Outputs["RIGHT"] = right1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(D1 / 2 - N, -D / 2, 0);
                Projection3d left1 = symbolGeometryHelper.CreateBox(null, N, D, A * 1.5, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI + Y * Math.PI / 180, new Vector(0, 1, 0));
                matrix.Translate(new Vector(A * Math.Sin(Y * Math.PI / 180), 0, length + A * Math.Cos(Y * Math.PI / 180) - takeOff));
                matrix.Rotate(X * Math.PI / 180, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                left1.Transform(matrix);
                m_Symbolic.Outputs["LEFT"] = left1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-D1 / 2 + N, -D / 2, 0);
                Projection3d end1 = symbolGeometryHelper.CreateBox(null, D1 - 2 * N, D, N, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI + Y * Math.PI / 180, new Vector(0, 1, 0));
                matrix.Translate(new Vector(A * Math.Sin(Y * Math.PI / 180), 0, length + A * Math.Cos(Y * Math.PI / 180) - takeOff));
                matrix.Rotate(X * Math.PI / 180, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                end1.Transform(matrix);
                m_Symbolic.Outputs["END"] = end1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d bolt1 = symbolGeometryHelper.CreateCylinder(null, T / 2, D1 + T);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(D1 / 2 + T / 2, 0, A));
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(Math.PI + Y * Math.PI / 180, new Vector(0, 1, 0), new Position(0, 0, 0));
                matrix.Translate(new Vector(A * Math.Sin(Y * Math.PI / 180), 0, length + A * Math.Cos(Y * Math.PI / 180) - takeOff));
                matrix.Rotate(X * Math.PI / 180, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                bolt1.Transform(matrix);
                m_Symbolic.Outputs["BOLT"] = bolt1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal1 = new Position(-S / 4, 0, length - takeOff).Subtract(new Position(S / 4, 0, length - takeOff));
                symbolGeometryHelper.ActivePosition = new Position(S / 4, 0, length - takeOff);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d topEye = symbolGeometryHelper.CreateCylinder(null, R, normal1.Length);
                m_Symbolic.Outputs["TOP_EYE"] = topEye;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal1 = new Position(0, -S / 4, length - (takeOff + CC)).Subtract(new Position(0, S / 4, length - (takeOff + CC)));
                symbolGeometryHelper.ActivePosition = new Position(0, S / 4, length - (takeOff + CC));
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d bottomEye = symbolGeometryHelper.CreateCylinder(null, R, normal1.Length);
                m_Symbolic.Outputs["BOT_EYE"] = bottomEye;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal1 = new Position(0, 0, length - (takeOff + F)).Subtract(new Position(0, 0, length - (takeOff + R)));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, length - (takeOff + R));
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d topEnd = symbolGeometryHelper.CreateCylinder(null, rodEnd / 2, normal1.Length);
                m_Symbolic.Outputs["TOP_END"] = topEnd;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal1 = new Position(0, 0, length - takeOff - CC + F).Subtract(new Position(0, 0, length - takeOff - CC + R));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, length - takeOff - CC + R);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d bottomEnd = symbolGeometryHelper.CreateCylinder(null, rodEnd / 2, normal1.Length);
                m_Symbolic.Outputs["BOT_END"] = bottomEnd;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal1 = new Position(0, 0, length - (takeOff + F)).Subtract(new Position(0, 0, length - takeOff - CC + F));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, length - takeOff - CC + F);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d extPiece = symbolGeometryHelper.CreateCylinder(null, extDiameter / 2, normal1.Length);
                m_Symbolic.Outputs["EXT_PIECE"] = extPiece;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG211"));
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
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                // To get FINISH
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG211", "FINISH");
                if (finishCodelist.PropValue == 0)
                    finishCodelist.PropValue = 1;
                if (finishCodelist.PropValue < 1 && finishCodelist.PropValue > 2)
                {
                    ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFinishvalue, "Finish value should be 1 or 2"));
                    return "";
                }
                string finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;

                double Y = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG211", "Y")).PropValue;
                string ndUnitType = (string)((PropertyValueString)part.GetPropertyValue("IJHgrDiameterSelection", "NDUnitType")).PropValue;
                string size = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrSize", "SIZE")).PropValue;
                double length = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                double A = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrAnvil_FIG211", "A")).PropValue;
                double F = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrAnvil_FIG211", "F")).PropValue;

                //  To get CONFIG
                string actualConfig = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrAnvil_FIG211", "CONFIG")).PropValue;

                // Get the pipe dia (it will be in meters)
                Double pipeDiameter = 0, E = 0, CC = 0, takeOff = 0, WValue = 0;
                string PipeNominalDiameter = null;
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
                            pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, pipeInfo.NominalDiameter.Size, UnitName.DISTANCE_INCH);
                        else if (pipeInfo.NominalDiameter.Units == "mm")
                            pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, pipeInfo.NominalDiameter.Size, UnitName.DISTANCE_MILLIMETER);
                    }
                    else
                    {
                        pipeDiameter = const_1;
                    }
                }
                catch
                {
                    pipeDiameter = const_1;
                }
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

                string clampSize;
                if (actualConfig == "3")
                {
                    PartClass anvilPNDParts = (PartClass)catalogBaseHelper.GetPartClass("Anvil_PND");
                    ReadOnlyCollection<BusinessObject> PNDParts = anvilPNDParts.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    foreach (BusinessObject part1 in PNDParts)
                    {
                        if (((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_PND", "Pipe_Dia_M")).PropValue > pipeDiameter - 0.001) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_PND", "Pipe_Dia_M")).PropValue < (pipeDiameter + 0.001)))
                        {
                            PipeNominalDiameter = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrAnvil_PND", "Pipe_Nom_Dia")).PropValue;
                            break;
                        }
                    }
                    clampSize = ", " + PipeNominalDiameter + " " + ndUnitType + " Clamp, ";
                }
                else
                    clampSize = ",";

                PartClass anvilFig211E1Parts = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG211_E1");
                ReadOnlyCollection<BusinessObject> fig211E1Parts = anvilFig211E1Parts.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject part1 in fig211E1Parts)
                {
                    if (((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG211_E1", "PIPE_DIA2")).PropValue > pipeDiameter - 0.001) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG211_E1", "PIPE_DIA2")).PropValue < (pipeDiameter + 0.001)))
                    {
                        E = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG211_E1", "SIZE_" + size)).PropValue;
                        break;
                    }
                }

                takeOff = A / Math.Cos(Y * Math.PI / 180);

                if (Math.Abs(Y) > 60)
                    takeOff = 0;

                CC = length - (takeOff + E);

                if (actualConfig == "1")
                    CC = length - takeOff;
                if (actualConfig == "2")
                    CC = length - A - takeOff;

                WValue = CC - 2 * F;
                string W = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, WValue, UnitName.DISTANCE_INCH);

                bomDescription = "Anvil FIG211 Size " + size + " Sway Strut Assembly (Option " + actualConfig + ")" + clampSize + "W: " + W + ", " + finish;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of Anvil_FIG211"));
            }
            return bomDescription;
        }
        #endregion


        #region "ICustomHgrWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {

                double weight, cogX, cogY, cogZ, weightPerUnitLength = 0;
                double length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;

                weight = weightPerUnitLength * length;
                cogX = 0;
                cogY = 0;
                cogZ = -length / 2;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrWeightCG, "Error while defining weightCG of SFS5381"));
            }
        }
        #endregion

    }

}
