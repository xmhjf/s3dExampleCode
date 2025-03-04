//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_END_PLATE_VAR.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_END_PLATE_VAR
//   Author       :  Hema
//   Creation Date:  30.10.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30.10.2012     Hema     CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//   30/Dec/2014    PVK       TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    [VariableOutputs]
    [SymbolVersion("1.0.0.0")]
    public class Utility_END_PLATE_VAR : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_END_PLATE_VAR"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "THICKNESS", "THICKNESS", 1)]
        public InputDouble m_dTHICKNESS;
        [InputDouble(3, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(4, "Length", "Length", 0.999999)]
        public InputDouble m_dLength;
        [InputDouble(5, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("Port2", "Port2")]
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

                Double thickness = m_dTHICKNESS.Value;
                Double W = m_dW.Value;
                Double length = m_dLength.Value;
                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double T, A, H;
                Double alpha = 0.0;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                PropertyValueCodelist thicknessCodeList;
                thicknessCodeList = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrUtility_END_PLATE_VAR", "THICKNESS");
                CodelistItem codeList;
                codeList = thicknessCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)thickness);
                if (codeList != null)
                {
                    if (codeList.Value < 1 || codeList.Value > 12)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidPlateThickness, "Thickness should be between 1 to 12"));
                    }
                }
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidWidth, "Width cannot be zero"));
                    return;
                }
                if (codeList != null)

                    T = Double.Parse(codeList.ShortDisplayName) * 25.4 / 1000;
                else
                    T = 0.009525;

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;


                if (W <= pipeDiameter)
                {
                    alpha = Math.Acos((W / 2) / (pipeDiameter / 2)) * 180 / Math.PI;
                }
                A = pipeDiameter / 2 - (pipeDiameter / 2 * Math.Sin(alpha * Math.PI / 180));
                if (pipeDiameter / 2 < W / 2)
                {
                    A = pipeDiameter / 2;
                    alpha = 0;
                }
                H = length - pipeDiameter / 2;
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Projection3d body;
                Matrix4X4 matrix = new Matrix4X4();
                if (W > pipeDiameter)
                {
                    curveCollection.Add(new Line3d(new Position(0, W / 2, pipeDiameter / 2 + H), new Position(0, -W / 2, pipeDiameter / 2 + H)));
                    curveCollection.Add(new Line3d(new Position(0, -W / 2, pipeDiameter / 2 + H), new Position(0, -W / 2, pipeDiameter / 2 - A)));
                    curveCollection.Add(new Line3d(new Position(0, -W / 2, pipeDiameter / 2 - A), new Position(0, -pipeDiameter / 2, pipeDiameter / 2 - A)));

                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI + alpha), new Vector(0, 0, 1));

                    Arc3d arc2 = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2, Math.PI - 2 * alpha);
                    arc2.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(1, 0, 0));
                    arc2.Transform(matrix);
                    curveCollection.Add(arc2);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                    arc2.Transform(matrix);
                    curveCollection.Add(arc2);

                    curveCollection.Add(new Line3d(new Position(0, pipeDiameter / 2, -pipeDiameter / 2 + A), new Position(0, W / 2, pipeDiameter / 2 - A)));
                    curveCollection.Add(new Line3d(new Position(0, W / 2, pipeDiameter / 2 - A), new Position(0, W / 2, pipeDiameter / 2 + H)));

                    body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), T, true);
                    m_Symbolic.Outputs["BODY"] = body;

                }
                else
                {
                    curveCollection.Add(new Line3d(new Position(0, W / 2, pipeDiameter / 2 + H), new Position(0, -W / 2, pipeDiameter / 2 + H)));
                    curveCollection.Add(new Line3d(new Position(0, -W / 2, pipeDiameter / 2 + H), new Position(0, -W / 2, pipeDiameter / 2 - A)));

                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI + alpha * Math.PI / 180), new Vector(0, 0, 1));

                    Arc3d arc2 = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2, (Math.PI - 2 * alpha * Math.PI / 180));
                    arc2.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(1, 0, 0));
                    arc2.Transform(matrix);
                    curveCollection.Add(arc2);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                    arc2.Transform(matrix);
                    curveCollection.Add(arc2);

                    curveCollection.Add(new Line3d(new Position(0, W / 2, pipeDiameter / 2 - A), new Position(0, W / 2, pipeDiameter / 2 + H)));

                    body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), T, true);
                    m_Symbolic.Outputs["BODY"] = body;

                }
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_END_PLATE_VAR"));
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
                Double T, alpha, A, length, W, pipedia;
                alpha = Math.Acos(2.0 / 3.0);

                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                length = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                W = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_END_PLATE_VAR", "W")).PropValue;
                pipedia = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue;

                PropertyValueCodelist thicknessCodeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_END_PLATE_VAR", "THICKNESS");
                CodelistItem codeList = thicknessCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(thicknessCodeList.PropValue);

                if (codeList != null)
                {
                    if (codeList.Value < 1 || codeList.Value > 12)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidPlateThickness, "Thickness should be between 1 to 12"));
                    }
                }

                if (codeList != null)

                    T = Double.Parse(codeList.ShortDisplayName) * 25.4 / 1000;
                else
                    T = 0.009525;

                if (W <= pipedia)
                {
                    alpha = Math.Acos((W / 2) / (pipedia / 2)) * (180 / Math.PI);
                }
                A = pipedia / 2 - (pipedia / 2 * Math.Sin(alpha * (Math.PI / 180)));
                if (pipedia / 2 < W / 2)
                {
                    A = pipedia / 2;
                    alpha = 0;
                }
                bomDescription = "End Plate - Pipe-Base L = " + Microsoft.VisualBasic.Conversion.Str(length - pipedia / 2) + ", Overall L = " + Microsoft.VisualBasic.Conversion.Str((length - pipedia / 2 + A)) + ", " + Microsoft.VisualBasic.Conversion.Str(W) + " X " + Microsoft.VisualBasic.Conversion.Str(T) + " Radius " + Microsoft.VisualBasic.Conversion.Str(pipedia / 2);
                return bomDescription;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_END_PLATE_VAR"));
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
                Double weight, cogX, cogY, cogZ, alpha = 0, arcLength, segmentArea, A;
                const int getSteelDensityKGPerM = 7900;

                double W = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_END_PLATE_VAR", "W")).PropValue;
                PropertyValueCodelist thicknesslist = (PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrUtility_END_PLATE_VAR", "THICKNESS");
                double T = double.Parse(thicknesslist.PropertyInfo.CodeListInfo.GetCodelistItem(thicknesslist.PropValue).DisplayName) * 25.4 / 1000;
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double pipedia = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue;
                double H = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue - pipedia / 2;

                if (W <= pipedia)
                {
                    alpha = Math.Acos((W / 2) / (pipedia / 2));
                }
                A = pipedia / 2 - (pipedia / 2 * Math.Sin(alpha));
                if (pipedia / 2 < W / 2)
                {
                    A = pipedia / 2;
                    alpha = 0;
                }
                arcLength = pipedia / 2 * ((180) - 2 * (alpha * 180 / Math.PI)) / 180 * Math.PI;
                segmentArea = 0.5 * (pipedia / 2 * arcLength - (W * (pipedia / 2 - A)));
                weight = ((H + A) * W - segmentArea) * T * getSteelDensityKGPerM;

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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_END_PLATE_VAR"));
                }
            }
        }
    }
        #endregion
}

