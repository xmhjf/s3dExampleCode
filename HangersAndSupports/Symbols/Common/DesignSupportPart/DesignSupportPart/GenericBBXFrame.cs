//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   GenericBBXFrame.cs
//   DesignSupportPart,Ingr.SP3D.Support.Content.Symbols.GenericBBXFrame
//   Author       :  Pavan
//   Creation Date:  28.Sep.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   08.Nov.2012    Pavan    DI-CP-220480  Check items in from Shelfsets
//   25.Nov.2013    BS       DI-CP-241803  Checked back .Net progid to DesignSupport part
//   11-Dec-2014     PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//   23-03-2015     Chethan  TR-CP-268570  Namespace inconsistency in .NET content for few H&S project  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Support.Content.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    [CacheOption(CacheOptionType.NonCached)]
    public class GenericBBXFrame : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "DesignSupportPart,Ingr.SP3D.Support.Content.Symbols.GenericBBXFrame"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Width", "Width", 2)]
        public InputDouble m_Width;
        [InputDouble(3, "Height", "Height", 2)]
        public InputDouble m_Height;
        [InputDouble(4, "DesignSupportType", "DesignSupportType", 1, true)]
        public InputDouble m_DesignSupportType;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("SimplePhysical", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("StructPort", "StructPort")]
        [SymbolOutput("PipePort", "PipePort")]
        [SymbolOutput("Frame1", "Frame1")]
        [SymbolOutput("Frame2", "Frame2")]
        [SymbolOutput("Frame3", "Frame3")]
        [SymbolOutput("Frame4", "Frame4")]
        [SymbolOutput("Frame5", "Frame5")]
        [SymbolOutput("Frame6", "Frame6")]
        [SymbolOutput("Frame7", "Frame7")]
        [SymbolOutput("Frame8", "Frame8")]
        [SymbolOutput("Frame9", "Frame9")]
        [SymbolOutput("Frame10", "Frame10")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                SP3DConnection connection = OccurrenceConnection;

                Part part = (Part)m_PartInput.Value;
                double width = m_Width.Value;
                double height = m_Height.Value;
                int designSupportType1;
                double insulationThickness=0, routePosX=0, routePosY=0, widthInsulation, heightInsulation;
                SupportHelper supportHelper;
                SupportedHelper supportedHelper;

                //Line3d line1, line2, line3, line4, line5 = null, line6 = null, line7 = null, line8 = null;
                Line3d[] line = new Line3d[9]; 
                try
                {
                    if (Occurrence.SupportsInterface("IJOAhsDesignSupportType"))
                    {
                        try
                        {
                            int designSupportType = (int)((PropertyValueCodelist)Occurrence.GetPropertyValue("IJOAhsDesignSupportType", "DesignSupportType")).PropValue;
                            designSupportType1 = (int)m_DesignSupportType.Value;
                        }
                        catch
                        {
                            designSupportType1 = 1;
                        }
                    }
                    else
                        designSupportType1 = 1;
                    //Determine if this is a Place By Structure Assembly
                    RelationCollection hgrRelation = Occurrence.GetRelationship("SupportHasComponents", "Support");
                    BusinessObject support = hgrRelation.TargetObjects[0];
                    supportHelper = new SupportHelper((Ingr.SP3D.Support.Middle.Support)support);
                    RefPortHelper refportHelper = new RefPortHelper((Ingr.SP3D.Support.Middle.Support)support);
                    Boolean isPlaceByStruct = false;
                    if (supportHelper.PlacementType == PlacementType.PlaceByStruct)
                        isPlaceByStruct = true;
                    BoundingBoxHelper boundingBoxHelper = new BoundingBoxHelper((Ingr.SP3D.Support.Middle.Support)support);
                    //When a Design Support is copy-pasted with deleting optional references then there is a chance that
                    //Bounding box Calculation will fail. for the reason that Support might have the Supported objects in
                    //relation, that will be made after the symbol compute.
                    boundingBoxHelper.CreateStandardBoundingBoxes(false);
                    // Get required information about the Bounding Box Surrounding the Pipe.
                    // The Bounding Box used depends on the command.
                    string boundingBox = string.Empty, boundingBoxName = string.Empty;
                    if (isPlaceByStruct)
                    {
                        // Use the Structure Bounding Box
                        boundingBox = "BBSR";
                        boundingBoxName = "BBSR_Low";
                    }
                    else
                    {
                        // Use the Route Bouding Box
                        boundingBox = "BBR";
                        boundingBoxName = "BBR_Low";
                    }

                    if (boundingBoxHelper.GetBoundingBox(boundingBox) != null)
                    {
                        for (int i = 1; i < supportHelper.SupportedObjects.Count + 1; i++)
                        {
                            double tempInsultaion;
                            supportedHelper = new SupportedHelper((Ingr.SP3D.Support.Middle.Support)support);
                            SupportedObjectInfo supportedObjectInfo = supportedHelper.SupportedObjectInfo(i);
                            if (supportedObjectInfo.GetType() == typeof(PipeObjectInfo))
                                tempInsultaion = ((PipeObjectInfo)supportedObjectInfo).InsulationThickness;
                            else if (supportedObjectInfo.GetType() == typeof(DuctObjectInfo))
                                tempInsultaion = ((DuctObjectInfo)supportedObjectInfo).InsulationThickness;
                            else
                                tempInsultaion = 0;
                            if (tempInsultaion > insulationThickness)
                                insulationThickness = tempInsultaion;
                        }
                        width = boundingBoxHelper.GetBoundingBox(boundingBox).Width;
                        height = boundingBoxHelper.GetBoundingBox(boundingBox).Height;

                        widthInsulation = width + (2 * insulationThickness);
                        heightInsulation = height + (2 * insulationThickness);

                        String boundingBoxLow = boundingBoxHelper.GetBoundingBox(boundingBox).LowReferencePortName;
                        Position primaryRoutePosition = boundingBoxHelper.GetBoundingBox(boundingBox).GetRelativeRouteCenterPosition(1);
                        if (primaryRoutePosition != null)
                        {
                            routePosX = primaryRoutePosition.X;
                            routePosY = primaryRoutePosition.Y;
                        }
                    }
                    else
                    {
                        width = 0.0;
                        height = 0.0;
                        widthInsulation = 0;
                        heightInsulation = 0;

                        routePosX = 0;
                        routePosY = 0;
                        insulationThickness = 0;
                        //149933 ; set the route comp port at correct location, if possible
                        
                        
                    }
                    //Construct in the x/y plane with origin at the lower left hand
                    //corner of the profile.
                    //The z-axis will serve as the "Thickness" direction of the Plate.
                    //========================================
                    // Error Check and Adject input values
                    //========================================
                    // Correct Height if necessary
                    if (height <= 0)
                        height = 0.1;

                    // Correct Width if necessary
                    if (width <= 0)
                        width = 0.1;
                    if (designSupportType1 == 1 || designSupportType1 == 3)
                    {
                        line[1] = new Line3d(connection, new Position(0, 0, 0), new Position(width, 0, 0));
                        line[2] = new Line3d(connection, new Position(width, 0, 0), new Position(width, height, 0));
                        line[3] = new Line3d(connection, new Position(width, height, 0), new Position(0, height, 0));
                        line[4] = new Line3d(connection, new Position(0, height, 0), new Position(0, 0, 0));
                    }
                    else
                    {
                        line[1] = new Line3d(connection, new Position(-insulationThickness, -insulationThickness, 0), new Position(widthInsulation - insulationThickness, -insulationThickness, 0));
                        line[2] = new Line3d(connection, new Position(widthInsulation - insulationThickness, -insulationThickness, 0), new Position(widthInsulation - insulationThickness, heightInsulation - insulationThickness, 0));
                        line[3] = new Line3d(connection, new Position(widthInsulation - insulationThickness, heightInsulation - insulationThickness, 0), new Position(-insulationThickness, heightInsulation - insulationThickness, 0));
                        line[4] = new Line3d(connection, new Position(-insulationThickness, heightInsulation - insulationThickness, 0), new Position(-insulationThickness, -insulationThickness, 0));
                    }


                    if (designSupportType1 == 3 && insulationThickness > 0)
                    {
                        line[5] = new Line3d(connection, new Position(-insulationThickness, -insulationThickness, 0), new Position(widthInsulation - insulationThickness, -insulationThickness, 0));
                        line[6] = new Line3d(connection, new Position(widthInsulation - insulationThickness, -insulationThickness, 0), new Position(widthInsulation - insulationThickness, heightInsulation - insulationThickness, 0));
                        line[7] = new Line3d(connection, new Position(widthInsulation - insulationThickness, heightInsulation - insulationThickness, 0), new Position(-insulationThickness, heightInsulation - insulationThickness, 0));
                        line[8] = new Line3d(connection, new Position(-insulationThickness, heightInsulation - insulationThickness, 0), new Position(-insulationThickness, -insulationThickness, 0));
                    }
                    //If it is only one
                    if (supportHelper.SupportedObjects.Count >0)
                    {
                        //get angle only if only one supproted object, get its orientation and use it to rotate BBox accordingly.
                        // this applicable only to duct and 
                        if (supportHelper.SupportedObjects.Count == 1)
                        {
                            double routeAngle = 0;
                            if (boundingBoxHelper.GetBoundingBox(boundingBox) != null)
                            {
                                supportedHelper = new SupportedHelper((Ingr.SP3D.Support.Middle.Support)support);
                                routeAngle = supportedHelper.SupportedObjectInfo(1).OrientationAngle; ////////////////////////////////////////////////////////////////////////////////Get Equivalent function here
                            }
                            else
                                routeAngle = 0;
                            if (Ingr.SP3D.Content.Support.Symbols.HgrCompareDoubleService.cmpdbl(routeAngle , 0)==false)
                            {
                                Matrix4X4 oMatrix = new Matrix4X4();
                                Vector oVector = new Vector(0,0,1);
                                double xTrans, yTrans, dAngle, dDiagonal;

                                dAngle = Math.Atan(height / width);
                                dDiagonal = (width / 2) * Math.Cos(dAngle) + (height / 2) * Math.Sin(dAngle);

                                xTrans = dDiagonal * Math.Cos(dAngle) - dDiagonal * Math.Cos(dAngle + routeAngle);
                                yTrans = dDiagonal * Math.Sin(dAngle) - dDiagonal * Math.Sin(dAngle + routeAngle);

                                oMatrix.SetIdentity();
                                oMatrix.Rotate(routeAngle, oVector);

                                line[1].Transform(oMatrix);
                                line[2].Transform(oMatrix);
                                line[3].Transform(oMatrix);
                                line[4].Transform(oMatrix);

                                if (designSupportType1 == 3 && insulationThickness > 0)
                                {
                                    line[5].Transform(oMatrix);
                                    line[6].Transform(oMatrix);
                                    line[7].Transform(oMatrix);
                                    line[8].Transform(oMatrix);
                                }
                                //after rotating this, we need to translate this box as rotation axis
                                // for duct/cable way is about center and for Bbox is about corner.
                                oMatrix.SetIdentity();
                                oVector.Set(xTrans, yTrans, 0);
                                oMatrix.Translate(oVector);

                                line[1].Transform(oMatrix);
                                line[2].Transform(oMatrix);
                                line[3].Transform(oMatrix);
                                line[4].Transform(oMatrix);

                                if (designSupportType1 == 3 && insulationThickness > 0)
                                {
                                    line[5].Transform(oMatrix);
                                    line[6].Transform(oMatrix);
                                    line[7].Transform(oMatrix);
                                    line[8].Transform(oMatrix);
                                }
                            }
                        }
                    }
                    for (int i = 1; i <= 4; i++)
                    {
                        m_PhysicalAspect.Outputs["Frame" + i] = line[i];
                    }
                    if (designSupportType1 == 3 && insulationThickness > 0)
                    {
                        for (int i = 5; i <= 8; i++)
                        {
                            m_PhysicalAspect.Outputs["Frame" + i] = line[i];
                        }
                    }
                }
                catch
                {
                    designSupportType1 = 1;
                    width = 0.0;
                    height = 0.0;
                    widthInsulation = 0;
                    heightInsulation = 0;
                    routePosX = 0;
                    routePosY = 0;
                    insulationThickness = 0;
                    //149933 ; set the route comp port at correct location, if possible


                    //Construct in the x/y plane with origin at the lower left hand
                    //corner of the profile.
                    //The z-axis will serve as the "Thickness" direction of the Plate.
                    //========================================
                    // Error Check and Adject input values
                    //========================================
                    // Correct Height if necessary
                    if (height <= 0)
                        height = 0.1;

                    // Correct Width if necessary
                    if (width <= 0)
                        width = 0.1;

                    if (designSupportType1 == 1 || designSupportType1 == 3)
                    {
                        line[1] = new Line3d(connection, new Position(0, 0, 0), new Position(width, 0, 0));
                        line[2] = new Line3d(connection, new Position(width, 0, 0), new Position(width, height, 0));
                        line[3] = new Line3d(connection, new Position(width, height, 0), new Position(0, height, 0));
                        line[4] = new Line3d(connection, new Position(0, height, 0), new Position(0, 0, 0));
                    }
                    else
                    {
                        line[1] = new Line3d(connection, new Position(-insulationThickness, -insulationThickness, 0), new Position(widthInsulation - insulationThickness, -insulationThickness, 0));
                        line[2] = new Line3d(connection, new Position(widthInsulation - insulationThickness, -insulationThickness, 0), new Position(widthInsulation - insulationThickness, heightInsulation - insulationThickness, 0));
                        line[3] = new Line3d(connection, new Position(widthInsulation - insulationThickness, heightInsulation - insulationThickness, 0), new Position(-insulationThickness, heightInsulation - insulationThickness, 0));
                        line[4] = new Line3d(connection, new Position(-insulationThickness, heightInsulation - insulationThickness, 0), new Position(-insulationThickness, -insulationThickness, 0));
                    }


                    if (designSupportType1 == 3 && insulationThickness > 0)
                    {
                        line[5] = new Line3d(connection, new Position(-insulationThickness, -insulationThickness, 0), new Position(widthInsulation - insulationThickness, -insulationThickness, 0));
                        line[6] = new Line3d(connection, new Position(widthInsulation - insulationThickness, -insulationThickness, 0), new Position(widthInsulation - insulationThickness, heightInsulation - insulationThickness, 0));
                        line[7] = new Line3d(connection, new Position(widthInsulation - insulationThickness, heightInsulation - insulationThickness, 0), new Position(-insulationThickness, heightInsulation - insulationThickness, 0));
                        line[8] = new Line3d(connection, new Position(-insulationThickness, heightInsulation - insulationThickness, 0), new Position(-insulationThickness, -insulationThickness, 0));
                    }

                    for (int i = 1; i <= 4; i++)
                    {
                        m_PhysicalAspect.Outputs["Frame" + i] = line[i];
                    }
                    if (designSupportType1 == 3 && insulationThickness > 0)
                    {
                        for (int i = 5; i <= 8; i++)
                        {
                            m_PhysicalAspect.Outputs["Frame" + i] = line[i];
                        }
                    }
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                //Add Ports
                Port structPort = new Port(connection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["StructPort"] = structPort;

                Port routePort = new Port(connection, part, "Route",new Position(routePosX, routePosY, 0),new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["PipePort"] = routePort;

            }
            catch //General Unhandled exception
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "Error in ConstructOutputs of GenericBBXFrame.");
                    return;
                }
            }
        }
        #endregion
    }
}






