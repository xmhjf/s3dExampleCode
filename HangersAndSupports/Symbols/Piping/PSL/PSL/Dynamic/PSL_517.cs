﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_517.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_517
//   Author       :  Rajeswari
//   Creation Date:  22-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   22-08-2013   Rajeswari  CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
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

    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class PSL_517 : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_517"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "Length", "Length", 0.999999)]
        public InputDouble Length;
        [InputDouble(3, "END_ROT", "END_ROT", 0.999999)]
        public InputDouble END_ROT;
        [InputDouble(4, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(5, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(6, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(7, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(8, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(9, "MIN_L", "MIN_L", 0.999999)]
        public InputDouble MIN_L;
        [InputDouble(10, "MAX_LENGTH", "MAX_LENGTH", 0.999999)]
        public InputDouble MAX_LENGTH;
        [InputDouble(11, "STROKE", "STROKE", 0.999999)]
        public InputDouble STROKE;
        [InputDouble(12, "RET", "RET", 0.999999)]
        public InputDouble RET;
        [InputDouble(13, "EXT", "EXT", 0.999999)]
        public InputDouble EXT;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("L_END_BOX", "L_END_BOX")]
        [SymbolOutput("L_END_CYL", "L_END_CYL")]
        [SymbolOutput("L_CYL", "L_CYL")]
        [SymbolOutput("L_MID", "L_MID")]
        [SymbolOutput("MID", "MID")]
        [SymbolOutput("R_MID", "R_MID")]
        [SymbolOutput("R_END_BOX", "R_END_BOX")]
        [SymbolOutput("R_END_CYL", "R_END_CYL")]
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
                Part part = (Part)PartInput.Value;
                Double a = A.Value;
                Double b = B.Value;
                Double c = C.Value;
                Double d = D.Value;
                Double l = Length.Value;
                Double endRot = END_ROT.Value;
                Double stroke = STROKE.Value;
                Double ret = RET.Value;
                Double ext = EXT.Value;
                Double minimumLength = MIN_L.Value;
                Double maximumLength = MAX_LENGTH.Value;

                // checking retraction
                if (l - ret < minimumLength)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidRetraction, "Retraction must be less than (Length - MIN_L)"));

                // checking extension
                if (l + ext > maximumLength)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidExtension, "Extension must be less than (MAX_LENGTH - Length)"));

                // checking ext+ret is less than stroke
                if (ret + ext > stroke)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidExtentionAndretraction, "(Extension + Retraction) must be less than STROKE : " + (stroke * 1000).ToString()));

                // checking overall length
                if (l < minimumLength || l > maximumLength)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning,"Length is out of Min and Max range: MIN:" + (minimumLength * 1000).ToString() + " Max:" + (maximumLength * 1000).ToString());

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, a / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, l + a / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (b <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBGTZ, "B value should be greater than zero"));
                    return;
                }
                if (c <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCGTZ, "C value should be greater than zero"));
                    return;
                }
                if (d <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDGTZ, "D value should be greater than zero"));
                    return;
                }
                Projection3d lEndBox = symbolGeometryHelper.CreateBox(null, b, c, b, 9);
                matrix.Translate(new Vector(-b / 2.0, -c / 2.0, a / 2.0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                lEndBox.Transform(matrix);
                m_Symbolic.Outputs["L_END_BOX"] = lEndBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d lEndCylinder = symbolGeometryHelper.CreateCylinder(null, b / 2.0, c);
                matrix.SetIdentity();
                matrix.Rotate(3 * Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -c / 2.0, a / 2.0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                lEndCylinder.Transform(matrix);
                m_Symbolic.Outputs["L_END_CYL"] = lEndCylinder;

                Vector normal = new Position(0, 0, a / 2.0 + b * 2.0).Subtract(new Position(0, 0, a / 2.0 + b));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, a / 2.0 + b);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d lCylinder = symbolGeometryHelper.CreateCylinder(null, d / 4.0, normal.Length);
                m_Symbolic.Outputs["L_CYL"] = lCylinder;

                normal = new Position(0, 0, a / 2.0 + b * 3.0).Subtract(new Position(0, 0, a / 2.0 + b * 2.0));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, a / 2.0 + b * 2.0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d lmid = symbolGeometryHelper.CreateCylinder(null, d / 2.0, normal.Length);
                m_Symbolic.Outputs["L_MID"] = lmid;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Projection3d mid = symbolGeometryHelper.CreateBox(null, d, d, b * 2.0, 9);
                matrix.SetIdentity();
                matrix.Translate(new Vector(-d / 2.0, -d / 2.0, a / 2.0 + b * 3.0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                mid.Transform(matrix);
                m_Symbolic.Outputs["MID"] = mid;

                normal = new Position(0, 0, a / 2.0 + l - b).Subtract(new Position(0, 0, a / 2.0 + b * 5.0));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, a / 2.0 + b * 5.0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d rmid = symbolGeometryHelper.CreateCylinder(null, d / 4.0, normal.Length);
                m_Symbolic.Outputs["R_MID"] = rmid;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Projection3d rEndBox = symbolGeometryHelper.CreateBox(null, b, c, b, 9);
                matrix.SetIdentity();
                matrix.Translate(new Vector(-b / 2.0, -c / 2.0, l + a / 2.0 - b));
                matrix.Rotate((endRot * 180 / Math.PI) * Math.PI / 180, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                rEndBox.Transform(matrix);
                m_Symbolic.Outputs["R_END_BOX"] = rEndBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d rEndCylinder = symbolGeometryHelper.CreateCylinder(null, b / 2.0, c);
                matrix.SetIdentity();
                matrix.Rotate(3 * Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -c / 2.0, l + a / 2.0));
                matrix.Rotate((endRot * 180 / Math.PI) * Math.PI / 180, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                rEndCylinder.Transform(matrix);
                m_Symbolic.Outputs["R_END_CYL"] = rEndCylinder;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_516.cs"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                Double minimumWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_517", "MIN_WEIGHT")).PropValue;
                Double maximumWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_517", "MAX_WEIGHT")).PropValue;
                Double minimumLength = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_517", "MIN_L")).PropValue;
                Double maximumLength = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_517", "MAX_LENGTH")).PropValue;
                Double L = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrOccLength_mm", "Length")).PropValue;

                // Calculate the weight by interpolating MIN_WEIGHT and MAX_WEIGHT
                double weight = minimumWeight + (maximumWeight - minimumWeight) * (L - minimumLength) / (maximumLength - minimumLength);

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, 0, 0, 0);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PSL_517.cs."));
            }
        }
        #endregion
    }
}
