//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bline_VertHanger_9_122X.cs
//   Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_VertHanger_9_122X
//   Author       :  Vijaya
//   Creation Date:  30.July.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30.July.2012    Vijaya   CR-CP-219113 Converted HS_B-Line VB project to C#.Net Project 
//   27.Nov.2012     Hema     CR-CP-219113 Modified the code with SymbolGeomHelper
//   26.Mar.2013     Vijaya     DI-CP-228142  Modify the error handling for delivered H&S symbols 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Bline_VertHanger_9_122X : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_VertHanger_9_122X"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "A", "Width of Hanger", 0)]
        public InputDouble m_dHangerWidth;
        [InputDouble(3, "HH", "Height of Hanger", 0)]
        public InputDouble m_dHangerHeight;
        [InputDouble(4, "TH", "Thickness of Hanger", 0)]
        public InputDouble m_dHangerThickness;
        [InputDouble(5, "LP", "Length of Plate", 0)]
        public InputDouble m_dPlateLength;
        [InputDouble(6, "WP", "Width of Chamfer", 0)]
        public InputDouble m_dChamferWidth;
        [InputDouble(7, "LS", "Distance from edge to hole", 0)]
        public InputDouble m_dLS;
        [InputDouble(8, "HT", "Cable Tray Height", 0)]
        public InputDouble m_dClampHeight;
        [InputDouble(9, "TrayWT", "Tray Web Thickness", 0)]
        public InputDouble m_dTrayWT;
        [InputDouble(10, "WT", "IJUAHgrBlineTrayWidth", 0)]
        public InputDouble m_dTrayWidth;

        #endregion

        #region "Definitions of Aspects and their outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("L_HANGER", "L_HANGER")]
        [SymbolOutput("L_PLATE", "L_PLATE")]
        [SymbolOutput("R_HANGER", "R_HANGER")]
        [SymbolOutput("R_PLATE", "R_PLATE")]

        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;


                double A = m_dHangerWidth.Value;
                double HH = m_dHangerHeight.Value;
                double TH = m_dHangerThickness.Value;
                double LP = m_dPlateLength.Value;
                double WP = m_dChamferWidth.Value;
                double LS = m_dLS.Value;
                double TW = m_dTrayWT.Value;
                double WT = m_dTrayWidth.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                if (TH <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidTH, "Hanger Thickness value should  be greater than 0."));
                if (A <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidA, "Hanger Width  value should  be greater than 0."));
                if (HH <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidHH, "Hanger Height value should  be greater than 0."));
                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "TrayPort", new Position(0, 0, 0), new Vector(0, 0, 1), new Vector(1, 0, 0));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "RodHole1", new Position(0, (0.5 * WT + TW + LS), HH / 2 + TH / 2 - 0.008), new Vector(0, 0, 1), new Vector(1, 0, 0));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "RodHole2", new Position(0, -(0.5 * WT + TW + LS), HH / 2 + TH / 2 - 0.008), new Vector(0, 0, 1), new Vector(1, 0, 0));
                m_Symbolic.Outputs["Port3"] = port3;

                Collection<Position> pointCollection = new Collection<Position>();
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Vector projectionVector = new Vector();

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 + TW, TH / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d leftPlate = (Projection3d)symbolGeometryHelper.CreateBox(null, TH, A, HH - TH);
                m_Symbolic.Outputs["L_PLATE"] = leftPlate;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -(WT / 2 + TW), TH / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(1, 0, 0));
                Projection3d rightPlate = (Projection3d)symbolGeometryHelper.CreateBox(null, TH, A, HH - TH);
                m_Symbolic.Outputs["R_PLATE"] = rightPlate;

                pointCollection.Add(new Position(A / 2, WT / 2 + TW, HH / 2));
                pointCollection.Add(new Position(WP / 2, WT / 2 + TW + LP, HH / 2));
                pointCollection.Add(new Position(-WP / 2, WT / 2 + TW + LP, HH / 2));
                pointCollection.Add(new Position(-A / 2, WT / 2 + TW, HH / 2));
                pointCollection.Add(new Position(A / 2, WT / 2 + TW, HH / 2));

                projectionVector.Set(0, 0, TH);
                Projection3d leftHanger = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_Symbolic.Outputs["L_HANGER"] = leftHanger;
                pointCollection.Clear();
                curveCollection.Clear();

                pointCollection.Add(new Position(A / 2, -(WT / 2 + TW), HH / 2));
                pointCollection.Add(new Position(WP / 2, -(WT / 2 + TW + LP), HH / 2));
                pointCollection.Add(new Position(-WP / 2, -(WT / 2 + TW + LP), HH / 2));
                pointCollection.Add(new Position(-A / 2, -(WT / 2 + TW), HH / 2));
                pointCollection.Add(new Position(A / 2, -(WT / 2 + TW), HH / 2));

                projectionVector.Set(0, 0, TH);
                Projection3d rightHanger = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_Symbolic.Outputs["R_HANGER"] = rightHanger;
            }
            catch //General Unhandled exception 
            {

                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Bline_VertHanger_9_122X."));
                    return;
                }
            }
        }
        #endregion
    }
}




