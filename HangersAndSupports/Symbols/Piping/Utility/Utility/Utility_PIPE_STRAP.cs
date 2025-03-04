//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_PIPE_STRAP.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_PIPE_STRAP
//   Author       :  Rajeswari
//   Creation Date:  05/11/2012 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   05/11/2012   Rajeswari  CR-CP-222288  Converted HS_Utility VB Project to C# .Net 
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols 
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
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
    public class Utility_PIPE_STRAP : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_PIPE_STRAP"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "NO_HOLES", "NO_HOLES", 1)]
        public InputDouble m_NO_HOLES;
        [InputDouble(3, "STRAP_O", "STRAP_O", 0.999999)]
        public InputDouble m_STRAP_O;
        [InputDouble(4, "STRAP_T", "STRAP_T", 0.999999)]
        public InputDouble m_STRAP_T;
        [InputDouble(5, "INNER_DIA", "INNER_DIA", 0.999999)]
        public InputDouble m_INNER_DIA;
        [InputDouble(6, "STRAP_L", "STRAP_L", 0.999999)]
        public InputDouble m_STRAP_L;
        [InputDouble(7, "STRAP_W", "STRAP_W", 0.999999)]
        public InputDouble m_STRAP_W;
        [InputDouble(8, "HOLE_INSET1", "HOLE_INSET1", 0.999999)]
        public InputDouble m_HOLE_INSET1;
        [InputDouble(9, "HOLE_INSET2", "HOLE_INSET2", 0.999999)]
        public InputDouble m_HOLE_INSET2;
        [InputDouble(10, "HOLE_SIZE", "HOLE_SIZE", 0.999999)]
        public InputDouble m_HOLE_SIZE;
        [InputDouble(11, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_PIPE_DIA;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("HOLE1", "HOLE1")]
        [SymbolOutput("HOLE2", "HOLE2")]
        [SymbolOutput("HOLE1", "HOLE1")]
        [SymbolOutput("HOLE2", "HOLE2")]
        [SymbolOutput("HOLE3", "HOLE3")]
        [SymbolOutput("HOLE4", "HOLE4")]
        [SymbolOutput("BODY", "BODY")]
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

                Double noOfHoles = m_NO_HOLES.Value;
                Double strapO = m_STRAP_O.Value;
                Double strapT = m_STRAP_T.Value;
                Double strapDiameter = m_INNER_DIA.Value;
                Double strapL = m_STRAP_L.Value;
                Double strapW = m_STRAP_W.Value;
                Double holeInset1 = m_HOLE_INSET1.Value;
                Double holeInset2 = m_HOLE_INSET2.Value;
                Double holeSize = m_HOLE_SIZE.Value;
                Double pipeDiameter = m_PIPE_DIA.Value;
                Matrix4X4 matrix = new Matrix4X4();

                if (holeSize <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrHoleSizeGTZero, "Hole Size should be greater than zero"));
                    return;
                }
                if (strapT == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidStrapT, "Strap Thickness cannot be zero"));
                    return;
                }
                if (strapL == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidStrapL, "Strap Length cannot be zero"));
                    return;
                }
                if (strapW <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidStrapW, "Strap Ear Width should be greater than zero"));
                    return;
                }
                Double holeLeninset;
                string actualNH;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                PropertyValueCodelist holesCodeList = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrUtility_PIPE_STRAP", "NO_HOLES");
                CodelistItem codelist = holesCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)noOfHoles);

                if (codelist != null)
                    actualNH = codelist.ShortDisplayName.Trim();
                else
                    actualNH = "0";

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                if (strapDiameter != 0)
                {
                    pipeDiameter = strapDiameter;
                }
                Port port2 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, pipeDiameter / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (actualNH == "2")
                {
                    holeLeninset = strapL / 2;
                    symbolGeometryHelper.ActivePosition = new Position(pipeDiameter / 2 + strapW - holeInset2, -strapL / 2 + holeLeninset, -(pipeDiameter / 2 - strapO - strapT));
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                    Projection3d hole1 = symbolGeometryHelper.CreateCylinder(null, holeSize / 2, strapT);
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    hole1.Transform(matrix);
                    m_Symbolic.Outputs["HOLE1"] = hole1;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-pipeDiameter / 2 - strapW + holeInset2, -strapL / 2 + holeLeninset, -(pipeDiameter / 2 - strapO - strapT));
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                    Projection3d hole2 = symbolGeometryHelper.CreateCylinder(null, holeSize / 2, strapT);
                    matrix = new Matrix4X4();
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    hole2.Transform(matrix);
                    m_Symbolic.Outputs["HOLE2"] = hole2;
                }

                if (actualNH == "4")
                {
                    //holeLeninset = strapL / 2;
                    holeLeninset = holeInset1;
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(pipeDiameter / 2 + strapW - holeInset2, -strapL / 2 + holeLeninset, -(pipeDiameter / 2 - strapO - strapT));
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                    Projection3d hole1 = symbolGeometryHelper.CreateCylinder(null, holeSize / 2, strapT);
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    hole1.Transform(matrix);
                    m_Symbolic.Outputs["HOLE1"] = hole1;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-pipeDiameter / 2 - strapW + holeInset2, -strapL / 2 + holeLeninset, -(pipeDiameter / 2 - strapO - strapT));
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                    Projection3d hole2 = symbolGeometryHelper.CreateCylinder(null, holeSize / 2, strapT);
                    matrix = new Matrix4X4();
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    hole2.Transform(matrix);
                    m_Symbolic.Outputs["HOLE2"] = hole2;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-pipeDiameter / 2 - strapW + holeInset2, strapL / 2 - holeLeninset, -(pipeDiameter / 2 - strapO - strapT));
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                    Projection3d hole3 = symbolGeometryHelper.CreateCylinder(null, holeSize / 2, strapT);
                    matrix = new Matrix4X4();
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    hole3.Transform(matrix);
                    m_Symbolic.Outputs["HOLE3"] = hole3;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(pipeDiameter / 2 + strapW - holeInset2, strapL / 2 - holeLeninset, -(pipeDiameter / 2 - strapO - strapT));
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                    Projection3d hole4 = symbolGeometryHelper.CreateCylinder(null, holeSize / 2, strapT);
                    matrix = new Matrix4X4();
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    hole4.Transform(matrix);
                    m_Symbolic.Outputs["HOLE4"] = hole4;
                }

                if (strapW < strapT)
                {
                    strapW = strapT;
                }

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                curveCollection.Add(new Line3d(new Position(-strapL / 2, -pipeDiameter / 2 - strapT, 0), new Position(-strapL / 2, -pipeDiameter / 2 - strapT, -(pipeDiameter / 2 - strapO - strapT))));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate((Math.PI), new Vector(0, 0, 1));

                Arc3d Arc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2 + strapT, -Math.PI);
                Arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                Arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, -strapL / 2, 0));
                Arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                Arc.Transform(matrix);
                curveCollection.Add(Arc);


                //calculated midpoint for arc
                curveCollection.Add(new Line3d(new Position(-strapL / 2, pipeDiameter / 2 + strapT, 0), new Position(-strapL / 2, pipeDiameter / 2 + strapT, -(pipeDiameter / 2 - strapO - strapT))));
                curveCollection.Add(new Line3d(new Position(-strapL / 2, pipeDiameter / 2 + strapW, -(pipeDiameter / 2 - strapO - strapT)), new Position(-strapL / 2, pipeDiameter / 2 + strapT, -(pipeDiameter / 2 - strapO - strapT))));
                curveCollection.Add(new Line3d(new Position(-strapL / 2, pipeDiameter / 2 + strapW, -(pipeDiameter / 2 - strapO)), new Position(-strapL / 2, pipeDiameter / 2 + strapW, -(pipeDiameter / 2 - strapO - strapT))));
                curveCollection.Add(new Line3d(new Position(-strapL / 2, pipeDiameter / 2, -(pipeDiameter / 2 - strapO)), new Position(-strapL / 2, pipeDiameter / 2 + strapW, -(pipeDiameter / 2 - strapO))));
                curveCollection.Add(new Line3d(new Position(-strapL / 2, pipeDiameter / 2, 0), new Position(-strapL / 2, pipeDiameter / 2, -(pipeDiameter / 2 - strapO))));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate((Math.PI), new Vector(0, 0, 1));

                Arc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2, -Math.PI);
                Arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                Arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, -strapL / 2, 0));
                Arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                Arc.Transform(matrix);
                curveCollection.Add(Arc);

                curveCollection.Add(new Line3d(new Position(-strapL / 2, -pipeDiameter / 2, 0), new Position(-strapL / 2, -pipeDiameter / 2, -(pipeDiameter / 2 - strapO))));
                curveCollection.Add(new Line3d(new Position(-strapL / 2, -pipeDiameter / 2, -(pipeDiameter / 2 - strapO)), new Position(-strapL / 2, -pipeDiameter / 2 - strapW, -(pipeDiameter / 2 - strapO))));
                curveCollection.Add(new Line3d(new Position(-strapL / 2, -pipeDiameter / 2 - strapW, -(pipeDiameter / 2 - strapO)), new Position(-strapL / 2, -pipeDiameter / 2 - strapW, -(pipeDiameter / 2 - strapO - strapT))));
                curveCollection.Add(new Line3d(new Position(-strapL / 2, -pipeDiameter / 2 - strapW, -(pipeDiameter / 2 - strapO - strapT)), new Position(-strapL / 2, -pipeDiameter / 2 - strapT, -(pipeDiameter / 2 - strapO - strapT))));

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), strapL, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_PIPE_STRAP"));
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
                double strapLValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_PIPE_STRAP", "STRAP_L")).PropValue;
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                double pipeDiameterValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue;
                double strapTValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_PIPE_STRAP", "STRAP_T")).PropValue;

                string pipediameter = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, pipeDiameterValue, UnitName.DISTANCE_INCH);
                string strapL = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, strapLValue, UnitName.DISTANCE_INCH);
                string strapT = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, strapTValue, UnitName.DISTANCE_INCH);

                string[] bomUnits = pipediameter.Split(' ');
                string bomHoles;
                PropertyValueCodelist noholesCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_PIPE_STRAP", "NO_HOLES");
                string noHoles = noholesCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(noholesCodelist.PropValue).DisplayName;

                if (noHoles == "0")
                {
                    bomHoles = "";
                }
                else
                {
                    bomHoles = ", " + noHoles + " Holes";
                }
                bomString = "Radius " + double.Parse(bomUnits[0]) / 2 + "" + bomUnits[1] + ", Pipe Strap " + strapL + " X " + strapT + ", Flat Bar" + bomHoles;
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_PIPE_STRAP"));
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
                double strapT = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_PIPE_STRAP", "STRAP_T")).PropValue;
                double strapL = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_PIPE_STRAP", "STRAP_L")).PropValue;
                double strapO = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_PIPE_STRAP", "STRAP_O")).PropValue;
                double strapW = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_PIPE_STRAP", "STRAP_W")).PropValue;
                double holeSize = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_PIPE_STRAP", "HOLE_SIZE")).PropValue;
                PropertyValueCodelist noholesCodelist = (PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrUtility_PIPE_STRAP", "NO_HOLES");
                double numHoles = noholesCodelist.PropValue;

                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double pipeDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue;
                weight = ((((pipeDiameter / 2 + strapT) * (pipeDiameter / 2 + strapT) * Math.PI * strapL) - (pipeDiameter / 2 * pipeDiameter / 2 * Math.PI * strapL)) / 2 + ((pipeDiameter / 2 - strapO - strapT) * strapT * strapL) * 2 + (strapT * strapL * strapW) * 2 - (Math.PI * (holeSize / 2) * (holeSize / 2) * strapT) * numHoles) * getSteelDensityKGPerM;

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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_PIPE_STRAP"));
                }
            }
        }

        #endregion
    }
}
