//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_PB1_CM.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_PB1_CM
//   Author       :  Manikanth
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013     Manikanth    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
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
    public class PSL_PB1_CM : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_PB1_CM"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "COMLIN_T", "COMLIN_T", 0.999999)]
        public InputDouble COMLIN_T;
        [InputString(3, "CLAMP_SIZE", "CLAMP_SIZE", "No Value")]
        public InputString CLAMP_SIZE;
        [InputDouble(4, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble PIPE_DIA;
        [InputDouble(5, "END_T", "END_T", 0.999999)]
        public InputDouble END_T;
        [InputDouble(6, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(7, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(8, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(9, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(10, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(11, "WEB_T", "WEB_T", 0.999999)]
        public InputDouble WEB_T;
        [InputDouble(12, "FLANGE_T", "FLANGE_T", 0.999999)]
        public InputDouble FLANGE_T;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BASEPLATE", "BASEPLATE")]
        [SymbolOutput("WEB", "WEB")]
        [SymbolOutput("EP1", "EP1")]
        [SymbolOutput("EP2", "EP2")]
        [SymbolOutput("CYL", "CYL")]
        [SymbolOutput("CYL2", "CYL2")]
        [SymbolOutput("COMLIN1", "COMLIN1")]
        [SymbolOutput("COMLIN2", "COMLIN2")]
        [SymbolOutput("BOX1", "BOX1")]
        [SymbolOutput("BOX2", "BOX2")]
        [SymbolOutput("BOX3", "BOX3")]
        [SymbolOutput("BOX4", "BOX4")]
        public AspectDefinition m_Symbolic;
        #endregion

        #region "Construct Outputs"
        protected override void ConstructOutputs()
        {
            try
            {
                double flangeThickness = FLANGE_T.Value;
                double endThickness = END_T.Value;
                double pipeDiameter = PIPE_DIA.Value;
                string clampSize = CLAMP_SIZE.Value;
                double a = A.Value;
                double b = B.Value;
                double c = C.Value;
                double d = D.Value;
                double e = E.Value;
                double webThickness = WEB_T.Value;
                double comlinThickness = COMLIN_T.Value;

                Dictionary<KeyValuePair<string, string>, Object> parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_PB1_CM_AUX", "SIZE"), clampSize.Trim());

                double clampF = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_PB1_CM_AUX", "IJUAHgrPSL_PB1_CM_AUX", "F", parameter);
                double clampE = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_PB1_CM_AUX", "IJUAHgrPSL_PB1_CM_AUX", "E", parameter);
                double clampB = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_PB1_CM_AUX", "IJUAHgrPSL_PB1_CM_AUX", "B", parameter);

                Part part = (Part)PartInput.Value;
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -c), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (a <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidAGTZ, "A value should be greater than zero"));
                    return;
                }
                if (flangeThickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidflangeTGTZ, "FLANGE_T value should be greater than zero"));
                    return;
                }
                if (webThickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidwebTGTZ, "WEB_T value should be greater than zero"));
                    return;
                }
                if (endThickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEndTGTZ, "END_T value should be greater than zero"));
                    return;
                }
                if (clampE <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEGTZ, "E value should be greater than zero"));
                    return;
                }

                Matrix4X4 rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(-a / 2, -(e - clampE - endThickness) / 2, -c));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d basePlateBox = symbolGeometryHelper.CreateBox(null, a, e - clampE - endThickness, flangeThickness, 9);
                basePlateBox.Transform(rotateMatrix);
                m_Symbolic.Outputs["BASEPLATE"] = basePlateBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(-webThickness / 2, -(e - clampE - endThickness) / 2, -c + flangeThickness));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d webBox = symbolGeometryHelper.CreateBox(null, webThickness, e - clampE - endThickness, (c - pipeDiameter / 2.0 - flangeThickness) * 0.75, 9);
                webBox.Transform(rotateMatrix);
                m_Symbolic.Outputs["WEB"] = webBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(-a / 2.0, -e / 2.0 + clampE / 2.0 - endThickness / 2.0, -c));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d ep1Box = symbolGeometryHelper.CreateBox(null, a, endThickness, c - pipeDiameter / 2.0 - clampF, 9);
                ep1Box.Transform(rotateMatrix);
                m_Symbolic.Outputs["EP1"] = ep1Box;

                symbolGeometryHelper = new SymbolGeometryHelper();
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(-a / 2, e / 2 - clampE / 2 - endThickness / 2, -c));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d ep2Box = symbolGeometryHelper.CreateBox(null, a, endThickness, c - pipeDiameter / 2.0 - clampF, 9);
                ep2Box.Transform(rotateMatrix);
                m_Symbolic.Outputs["EP2"] = ep2Box;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(e / 2.0, 0, 0).Subtract(new Position(e / 2.0 - clampE, 0, 0));
                symbolGeometryHelper.ActivePosition = new Position(e / 2.0 - clampE, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d cyl = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2.0 + clampF, normal.Length);
                m_Symbolic.Outputs["CYL"] = cyl;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(-e / 2.0 + clampE, 0, 0).Subtract(new Position(-e / 2.0, 0, 0));
                symbolGeometryHelper.ActivePosition = new Position(-e / 2.0, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d cyl2 = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2.0 + clampF, normal.Length);
                m_Symbolic.Outputs["CYL2"] = cyl2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(e / 2 - clampE / 2.0 + clampE * 0.75, 0, 0).Subtract(new Position(e / 2 - clampE / 2 - clampE * 0.75, 0, 0));
                symbolGeometryHelper.ActivePosition = new Position(e / 2 - clampE / 2 - clampE * 0.75, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d comlin1Box = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2.0 + comlinThickness, normal.Length);
                m_Symbolic.Outputs["COMLIN1"] = comlin1Box;


                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(-e / 2 + clampE / 2 + clampE * 0.75, 0, 0).Subtract(new Position(-e / 2 + clampE / 2 - clampE * 0.75, 0, 0));
                symbolGeometryHelper.ActivePosition = new Position(-e / 2 + clampE / 2 - clampE * 0.75, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d comlin2Box = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2.0 + comlinThickness, normal.Length);
                m_Symbolic.Outputs["COMLIN2"] = comlin2Box;

                symbolGeometryHelper = new SymbolGeometryHelper();
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(pipeDiameter / 2, -e / 2, -(clampB + clampF * 2) / 2));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d box1 = symbolGeometryHelper.CreateBox(null, d / 2.0 - pipeDiameter / 2.0, clampE, clampB + clampF * 2, 9);
                box1.Transform(rotateMatrix);
                m_Symbolic.Outputs["BOX1"] = box1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(-d / 2, -e / 2, -(clampB + clampF * 2) / 2));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d box2 = symbolGeometryHelper.CreateBox(null, d / 2.0 - pipeDiameter / 2.0, clampE, clampB + clampF * 2, 9);
                box2.Transform(rotateMatrix);
                m_Symbolic.Outputs["BOX2"] = box2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(pipeDiameter / 2, e / 2 - clampE, -(clampB + clampF * 2) / 2));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d box3 = symbolGeometryHelper.CreateBox(null, d / 2.0 - pipeDiameter / 2.0, clampE, clampB + clampF * 2, 9);
                box3.Transform(rotateMatrix);
                m_Symbolic.Outputs["BOX3"] = box3;

                symbolGeometryHelper = new SymbolGeometryHelper();
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(-d / 2, e / 2 - clampE, -(clampB + clampF * 2) / 2));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d box4 = symbolGeometryHelper.CreateBox(null, d / 2.0 - pipeDiameter / 2.0, clampE, clampB + clampF * 2, 9);
                box4.Transform(rotateMatrix);
                m_Symbolic.Outputs["BOX4"] = box4;


            }
            catch (Exception)
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_PB1_CM.cs."));
                return;
            }
        }

        #endregion

    }
}
