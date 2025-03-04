//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   CS1_Clamp_2.cs
//    Power,Ingr.SP3D.Content.Support.Symbols.CS1_Clamp_2
//   Author       :  Rajeswari
//   Creation Date:  14-Dec-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   14-Dec-2012   Rajeswari CR222282 .Net HS_Power project creation
//	 25-Mar-2013    Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   05/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//   30-Dec-2014    PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    [VariableOutputs]
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class CS1_Clamp_2 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Power,Ingr.SP3D.Content.Support.Symbols.CS1_Clamp_2"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Dw", "Dw", 0.999999)]
        public InputDouble m_dDw;
        [InputDouble(3, "e", "e", 0.999999)]
        public InputDouble m_de;
        [InputDouble(4, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(5, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(6, "d", "d", 0.999999)]
        public InputDouble m_dd;
        [InputDouble(7, "G1", "G1", 0.999999)]
        public InputDouble m_dG1;
        [InputDouble(8, "ClampHeight", "ClampHeight", 0.999999)]
        public InputDouble m_dClampHeight;
        [InputDouble(9, "b", "b", 0.999999)]
        public InputDouble m_db;
        [InputDouble(10, "FINISH", "FINISH", 0.999999)]
        public InputDouble m_dFINISH;
        [InputDouble(11, "Angle", "Angle", 0.999999)]
        public InputDouble m_dAngle;
        [InputString(12, "LoadType", "LoadType", "No Value")]
        public InputString m_sLoadType;
        [InputDouble(13, "G", "G", 0.999999)]
        public InputDouble m_dG;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("LEFT_FRONT", "LEFT_FRONT")]
        [SymbolOutput("LEFT_BACK", "LEFT_BACK")]
        [SymbolOutput("RIGHT_FRONT", "RIGHT_FRONT")]
        [SymbolOutput("RIGHT_BACK", "RIGHT_BACK")]
        [SymbolOutput("LEFT_OUTER", "LEFT_OUTER")]
        [SymbolOutput("LEFT_INNER", "LEFT_INNER")]
        [SymbolOutput("RIGHT_OUTER", "RIGHT_OUTER")]
        [SymbolOutput("RIGHT_INNER", "RIGHT_INNER")]
        [SymbolOutput("LINE", "LINE")]
        [SymbolOutput("STIFFNER1", "STIFFNER1")]
        [SymbolOutput("STIFFNER2", "STIFFNER2")]
        [SymbolOutput("STIFFNER3", "STIFFNER3")]
        [SymbolOutput("STIFFNER4", "STIFFNER4")]
        [SymbolOutput("PLATE1", "PLATE1")]
        [SymbolOutput("PLATE2", "PLATE2")]
        [SymbolOutput("PLATE3", "PLATE3")]
        [SymbolOutput("PLATE4", "PLATE4")]
        [SymbolOutput("PLATE5", "PLATE5")]
        [SymbolOutput("PLATE6", "PLATE6")]
        [SymbolOutput("PLATE7", "PLATE7")]
        [SymbolOutput("PLATE8", "PLATE8")]

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
              
                Double pipeDiameter = m_dDw.Value;
                Double e = m_de.Value;
                Double L = m_dL.Value;
                Double C = m_dC.Value;
                Double D = m_dd.Value;
                Double G1 = m_dG1.Value;
                Double B = m_dClampHeight.Value;
                Double S = m_db.Value;
                Double angle = m_dAngle.Value;
                Double G = m_dG.Value;
                Double H = 0.05;
                G1 = 0.01;
                if (pipeDiameter <= 0 && L <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidPipeDiameterAndLGTZero, "Pipe Diameter and L should be greater than zero"));
                    return;
                }
                Double OParameter = (L / 2 - C / 2);

                string[] outputString = new string[23];

                if (G1 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidG1GTZero, "G1 should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidDGTZero, "D should be greater than zero"));
                    return;
                }
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidB, "B should be greater than zero"));
                    return;
                }
                if (L <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidLGTZero, "L should be greater than zero"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "LeftPin", new Position(0, C / 2, B / 2.0 - e - D / 2.0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "RightPin", new Position(0, -C / 2, B / 2.0 - e - D / 2.0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;


                symbolGeometryHelper.ActivePosition = new Position(0, 0, -B / 2.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2.0 + G1, B);
                m_Symbolic.Outputs["BODY"] = body;
                outputString[0] = "BODY";

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-S / 2.0 - G1 / 2.0, pipeDiameter / 2, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d leftFront = (Projection3d)symbolGeometryHelper.CreateBox(null, C / 2 + OParameter - pipeDiameter / 2.0, G1, B);
                m_Symbolic.Outputs["LEFT_FRONT"] = leftFront;
                outputString[1] = "LEFT_FRONT";

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(S / 2.0 + G1 / 2.0, pipeDiameter / 2, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d leftBack = (Projection3d)symbolGeometryHelper.CreateBox(null, C / 2 + OParameter - pipeDiameter / 2.0, G1, B);
                m_Symbolic.Outputs["LEFT_BACK"] = leftBack;
                outputString[2] = "LEFT_BACK";

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-S / 2.0 - G1 / 2.0, -pipeDiameter / 2, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(1, 0, 0));
                Projection3d rightFront = (Projection3d)symbolGeometryHelper.CreateBox(null, C / 2 + OParameter - pipeDiameter / 2.0, G1, B);
                m_Symbolic.Outputs["RIGHT_FRONT"] = rightFront;
                outputString[3] = "RIGHT_FRONT";

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(S / 2.0 + G1 / 2.0, -pipeDiameter / 2, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(1, 0, 0));
                Projection3d rightBack = (Projection3d)symbolGeometryHelper.CreateBox(null, C / 2 + OParameter - pipeDiameter / 2.0, G1, B);
                m_Symbolic.Outputs["RIGHT_BACK"] = rightBack;
                outputString[4] = "RIGHT_BACK";

                Vector normal1 = new Position((S / 2.0 + 2.0 * G1), -C / 2, B / 2.0 - e).Subtract(new Position(-(S / 2.0 + 2.0 * G1), -C / 2, B / 2.0 - e));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(S / 2.0 + 2.0 * G1), -C / 2, B / 2.0 - e);
                symbolGeometryHelper.SetOrientation(normal1, new Vector(0, 1, 0));
                Projection3d leftOuter = symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal1.Length);
                m_Symbolic.Outputs["LEFT_OUTER"] = leftOuter;
                outputString[5] = "LEFT_OUTER";

                Vector normal2 = new Position((S / 2.0 + 2.0 * G1), -(G / 2), 0).Subtract(new Position(-(S / 2.0 + 2.0 * G1), -(G / 2), 0));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(S / 2.0 + 2.0 * G1), -(G / 2), 0);
                symbolGeometryHelper.SetOrientation(normal2, new Vector(0, 1, 0));
                Projection3d leftInner = symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal2.Length);
                m_Symbolic.Outputs["LEFT_INNER"] = leftInner;
                outputString[6] = "LEFT_INNER";

                Vector normal3 = new Position((S / 2.0 + 2.0 * G1), C / 2, B / 2.0 - e).Subtract(new Position(-(S / 2.0 + 2.0 * G1), C / 2, B / 2.0 - e));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(S / 2.0 + 2.0 * G1), C / 2, B / 2.0 - e);
                symbolGeometryHelper.SetOrientation(normal3, new Vector(0, 1, 0));
                Projection3d rightOuter = symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal3.Length);
                m_Symbolic.Outputs["RIGHT_OUTER"] = rightOuter;
                outputString[7] = "RIGHT_OUTER";

                Vector normal4 = new Position((S / 2.0 + 2.0 * G1), (G / 2), 0).Subtract(new Position(-(S / 2.0 + 2.0 * G1), (G / 2), 0));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(S / 2.0 + 2.0 * G1), (G / 2), 0);
                symbolGeometryHelper.SetOrientation(normal4, new Vector(0, 1, 0));
                Projection3d rightInner = symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal4.Length);
                m_Symbolic.Outputs["RIGHT_INNER"] = rightInner;
                outputString[8] = "RIGHT_INNER";

                Line3d line = new Line3d(new Position(0, 0, 0), new Position(0, 0, B));
                m_Symbolic.Outputs["LINE"] = line;
                outputString[9] = "LINE";

                Double tempAngle1, tempAngle2, temp1, temp2;
                tempAngle1 = (G1 / pipeDiameter);
                tempAngle2 = (G1 / 2) / ((pipeDiameter / 2) + G1);
                temp1 = Math.Sqrt((pipeDiameter / 2) * (pipeDiameter / 2) + (G1 / 2) * (G1 / 2));
                temp2 = Math.Sqrt(((pipeDiameter / 2) + G1) * ((pipeDiameter / 2) + G1) + (G1 / 2) * (G1 / 2));

                int iCount;
                string output;
                for (iCount = 1; iCount <= 4; iCount++)
                {
                    Line3d line1;
                    Collection<ICurve> curveCollection = new Collection<ICurve>();
                    //const double PI = 3.1415926;
                    line1 = new Line3d(new Position(temp1 * Math.Sin((2 * iCount - 1) * (Math.PI / 4) + tempAngle1), temp1 * Math.Cos((2 * iCount - 1) * (Math.PI / 4) + tempAngle1), B / 2), new Position(temp1 * Math.Sin((2 * iCount - 1) * (Math.PI / 4) - tempAngle1), temp1 * Math.Cos((2 * iCount - 1) * (Math.PI / 4) - tempAngle1), B / 2));
                    curveCollection.Add(line1);
                    line1 = new Line3d(new Position(temp1 * Math.Sin((2 * iCount - 1) * (Math.PI / 4) - tempAngle1), temp1 * Math.Cos((2 * iCount - 1) * (Math.PI / 4) - tempAngle1), B / 2), new Position(temp2 * Math.Sin((2 * iCount - 1) * (Math.PI / 4) - tempAngle2), temp2 * Math.Cos((2 * iCount - 1) * (Math.PI / 4) - tempAngle2), B / 2));
                    curveCollection.Add(line1);
                    line1 = new Line3d(new Position(temp2 * Math.Sin((2 * iCount - 1) * (Math.PI / 4) - tempAngle2), temp2 * Math.Cos((2 * iCount - 1) * (Math.PI / 4) - tempAngle2), B / 2), new Position(temp2 * Math.Sin((2 * iCount - 1) * (Math.PI / 4) + tempAngle2), temp2 * Math.Cos((2 * iCount - 1) * (Math.PI / 4) + tempAngle2), B / 2));
                    curveCollection.Add(line1);
                    line1 = new Line3d(new Position(temp2 * Math.Sin((2 * iCount - 1) * (Math.PI / 4) + tempAngle2), temp2 * Math.Cos((2 * iCount - 1) * (Math.PI / 4) + tempAngle2), B / 2), new Position(temp1 * Math.Sin((2 * iCount - 1) * (Math.PI / 4) + tempAngle1), temp1 * Math.Cos((2 * iCount - 1) * (Math.PI / 4) + tempAngle1), B / 2));
                    curveCollection.Add(line1);

                    output = "STIFFNER" + iCount;

                    ComplexString3d lineString = new ComplexString3d(curveCollection);
                    Vector lineVector = new Vector(0, 0, H);
                    Projection3d lines = new Projection3d(lineString, lineVector, lineVector.Length, false);
                    m_Symbolic.Outputs[output] = lines;
                    outputString[9 + iCount] = output;
                }

                int itemp = 1;
                Double plateZ = 0.9 * B / 2.0;
                double tempangle1value;
                for (iCount = 0; iCount <= 2; iCount++)
                {
                    //plate1
                    tempangle1value = ((S / 2.0) + G1) / ((pipeDiameter / 2.0) + G1);
                    tempAngle1 = Math.Asin(tempangle1value);

                    Line3d line1;
                    Collection<ICurve> curveCollection = new Collection<ICurve>();

                    line1 = new Line3d(new Position(((S / 2.0) + G1), ((pipeDiameter) + G1), plateZ), new Position(((S / 2.0) + G1), ((pipeDiameter / 2) + G1) * Math.Cos(tempAngle1), plateZ));
                    curveCollection.Add(line1);

                    Matrix4X4 matrix = new Matrix4X4();
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    matrix.SetIdentity();
                    matrix.Rotate(tempAngle1, new Vector(0, 0, 1));

                    Arc3d arc1 = symbolGeometryHelper.CreateArc(null, ((pipeDiameter / 2) + G1), ((Math.PI / 4) - tempAngle1));
                    arc1.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, 0, plateZ));
                    arc1.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(((Math.PI / 4) - tempAngle1), new Vector(0, 0, 1));
                    arc1.Transform(matrix);

                    curveCollection.Add(arc1);

                    line1 = new Line3d(new Position(((S / 2.0) + G1), ((pipeDiameter) + G1), plateZ), new Position(((pipeDiameter / 2) + G1) * Math.Sin(Math.PI / 4), ((pipeDiameter / 2) + G1) * Math.Cos(Math.PI / 4), plateZ));
                    curveCollection.Add(line1);

                    output = "PLATE" + itemp;

                    ComplexString3d lineStringPlate1 = new ComplexString3d(curveCollection);
                    Vector lineVectorPlate1 = new Vector(0, 0, -0.1 * B / 2);
                    Projection3d plate1 = new Projection3d(lineStringPlate1, lineVectorPlate1, lineVectorPlate1.Length, true);
                    m_Symbolic.Outputs[output] = plate1;
                    outputString[13 + itemp] = output;
                    itemp = itemp + 1;
                    //plate 2
                    Line3d line2;
                    Collection<ICurve> curveCollection2 = new Collection<ICurve>();
                    line2 = new Line3d(new Position(-((S / 2) + G1), ((pipeDiameter) + G1), plateZ), new Position(-((S / 2) + G1), ((pipeDiameter / 2) + G1) * Math.Cos(tempAngle1), plateZ));
                    curveCollection2.Add(line2);

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(tempAngle1, new Vector(0, 0, 1));

                    Arc3d arc2 = symbolGeometryHelper.CreateArc(null, ((pipeDiameter / 2) + G1), (Math.PI / 4) - tempAngle1);
                    arc2.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, 0, plateZ));
                    arc2.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI / 2), new Vector(0, 0, 1));
                    arc2.Transform(matrix);
                    curveCollection2.Add(arc2);

                    line2 = new Line3d(new Position(-((S / 2) + G1), ((pipeDiameter) + G1), plateZ), new Position(-((pipeDiameter / 2) + G1) * Math.Sin(Math.PI / 4), ((pipeDiameter / 2) + G1) * Math.Cos(Math.PI / 4), plateZ));
                    curveCollection2.Add(line2);

                    output = "PLATE" + itemp;

                    ComplexString3d lineStringPlate2 = new ComplexString3d(curveCollection2);
                    Vector lineVectorPlate2 = new Vector(0, 0, -0.1 * B / 2);
                    Projection3d plate2 = new Projection3d(lineStringPlate2, lineVectorPlate2, lineVectorPlate2.Length, true);
                    m_Symbolic.Outputs[output] = plate2;
                    outputString[13 + itemp] = output;
                    itemp = itemp + 1;

                    //plate 3
                    Line3d line3;
                    Collection<ICurve> curveCollection3 = new Collection<ICurve>();
                    line3 = new Line3d(new Position(-((S / 2) + G1), -((pipeDiameter) + G1), plateZ), new Position(-((S / 2) + G1), -((pipeDiameter / 2) + G1) * Math.Cos(tempAngle1), plateZ));
                    curveCollection3.Add(line3);

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(tempAngle1, new Vector(0, 0, 1));

                    Arc3d arc3 = symbolGeometryHelper.CreateArc(null, ((pipeDiameter / 2) + G1), (Math.PI / 4) - tempAngle1);
                    arc3.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, 0, plateZ));
                    arc3.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(((Math.PI) + (Math.PI / 4) - tempAngle1), new Vector(0, 0, 1));
                    arc3.Transform(matrix);
                    curveCollection3.Add(arc3);

                    line3 = new Line3d(new Position(-((S / 2) + G1), -((pipeDiameter) + G1), plateZ), new Position(-((pipeDiameter / 2) + G1) * Math.Sin(Math.PI / 4), -((pipeDiameter / 2) + G1) * Math.Cos(Math.PI / 4), plateZ));
                    curveCollection3.Add(line3);

                    output = "PLATE" + itemp;

                    ComplexString3d lineStringPlate3 = new ComplexString3d(curveCollection3);
                    Vector lineVectorPlate3 = new Vector(0, 0, -0.1 * B / 2);
                    Projection3d plate3 = new Projection3d(lineStringPlate3, lineVectorPlate3, lineVectorPlate3.Length, true);
                    m_Symbolic.Outputs[output] = plate3;
                    outputString[13 + itemp] = output;
                    itemp = itemp + 1;

                    //plate 4
                    Line3d line4;
                    Collection<ICurve> curveCollection4 = new Collection<ICurve>();
                    line4 = new Line3d(new Position(((S / 2) + G1), -((pipeDiameter) + G1), plateZ), new Position(((S / 2) + G1), -((pipeDiameter / 2) + G1) * Math.Cos(tempAngle1), plateZ));
                    curveCollection4.Add(line4);

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(tempAngle1, new Vector(0, 0, 1));

                    Arc3d arc4 = symbolGeometryHelper.CreateArc(null, ((pipeDiameter / 2) + G1), (Math.PI / 4) - tempAngle1);
                    arc4.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, 0, plateZ));
                    arc4.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                    arc4.Transform(matrix);
                    curveCollection4.Add(arc4);

                    line4 = new Line3d(new Position(((S / 2) + G1), -((pipeDiameter) + G1), plateZ), new Position(((pipeDiameter / 2) + G1) * Math.Sin(Math.PI / 4), -((pipeDiameter / 2) + G1) * Math.Cos(Math.PI / 4), plateZ));
                    curveCollection4.Add(line4);

                    output = "PLATE" + itemp;

                    ComplexString3d lineStringPlate4 = new ComplexString3d(curveCollection4);
                    Vector lineVectorPlate4 = new Vector(0, 0, -0.1 * B / 2);
                    Projection3d plate4 = new Projection3d(lineStringPlate4, lineVectorPlate4, lineVectorPlate4.Length, true);
                    m_Symbolic.Outputs[output] = plate4;
                    outputString[13 + itemp] = output;
                    itemp = itemp + 1;

                    plateZ = -(0.8 * B / 2);

                    iCount++;
                }

                if (angle > 0)
                {
                    Matrix4X4 matrix = new Matrix4X4();
                    matrix.SetIdentity();

                    matrix.Rotate(angle, new Vector(0, 1, 0));

                    for (int i = 0; i < outputString.Length; i++)
                        if (outputString[i] != null)
                        {
                            Geometry3d transformObject = (Geometry3d)m_Symbolic.Outputs[outputString[i]];
                            transformObject.Transform(matrix);
                        }
                }
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of CS1_Clamp_2"));
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
                string inputBomDescripttion = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJUAHgrPowerBomDesc", "InputBomDesc")).PropValue;

                if (inputBomDescripttion == "None")
                {
                    bomDescription = "";
                }
                else
                {
                    if (inputBomDescripttion == null)
                    {
                        bomDescription = catalogPart.PartDescription;
                    }
                    else
                    {
                        bomDescription = inputBomDescripttion.Trim();
                    }
                }
                return bomDescription;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of CS1_Clamp_2"));
                }
                return "";
            }
        }
        #endregion
    }
}
