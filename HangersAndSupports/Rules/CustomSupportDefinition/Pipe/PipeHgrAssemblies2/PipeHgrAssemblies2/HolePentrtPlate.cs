//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PentrtPlate.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.PentrtPlate
//   Author       : Vinay
//   Creation Date: 09-Sep-2015
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    public class HolePentrtPlate : CustomSupportDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.HolePentrtPlate"
        //----------------------------------------------------------------------------------
        //Constants
        private const string HGRSUPHOLEPENTRTPLATE = "HgrSupHolePentrtPlate";

        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    //Gets SupportHelper
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

                    BoundingBoxHelper.CreateStandardBoundingBoxes(false);

                    //Create the list of part classes required by the type
                    PartClass hgrSupPentrtPlatePartClass = (PartClass)catalogBaseHelper.GetPartClass("HgrSupPentrtPltHole");

                    //Use the default selection rule to get a catalog part for each part class
                    string partselection = hgrSupPentrtPlatePartClass.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();
                    parts.Add(new PartInfo(HGRSUPHOLEPENTRTPLATE, "HgrSupPentrtPltHole", partselection));

                    // Return the collection of Catalog Parts
                    return parts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                // Don't support the Route Connection Button
                return 0;
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                //Get IJHgrInputConfig Hlpr Interface off of passed Helper
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;


                //Have to set the occurrence attributes on the Penetration Plate
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                Boolean excludeNotes;
                //Check for "ExcludeNotes" attribute (for migrated DB)
                if (support.SupportsInterface("IJUAhsExcludeNotes"))
                    excludeNotes = (Boolean)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsExcludeNotes", "ExcludeNotes")).PropValue;
                else
                    excludeNotes = false;

                if (excludeNotes == false)
                {
                    //give the name of the note as input of the function CreateNote()
                    Note note1 = CreateNote("Note AIR1");

                    PropertyValueCodelist note1PropertyValueCL = (PropertyValueCodelist)note1.GetPropertyValue("IJGeneralNote", "Purpose");
                    CodelistItem codeList = note1PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);

                    //modify the properties of the note
                    note1.SetPropertyValue("This note is created from AIR", "IJGeneralNote", "Text");

                    //codeList value 3 means fabrication
                    note1.SetPropertyValue(codeList, "IJGeneralNote", "Purpose");
                    note1.SetPropertyValue(false, "IJGeneralNote", "Dimensioned");
                }
                else
                    DeleteNoteIfExists("Note AIR1");

                Note note2 = null;
                //similarly create other note for support assembly
                if (excludeNotes == false)
                {
                    //give the name of the note as input of the function CreateNote()
                    note2 = CreateNote("Note AIR2");

                    PropertyValueCodelist note2PropertyValueCL = (PropertyValueCodelist)note2.GetPropertyValue("IJGeneralNote", "Purpose");
                    CodelistItem codeList = note2PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)2);

                    note2.SetPropertyValue("This is second note created from AIR", "IJGeneralNote", "Text");
                    note2.SetPropertyValue(codeList, "IJGeneralNote", "Purpose");
                    note2.SetPropertyValue(false, "IJGeneralNote", "Dimensioned");
                }
                else
                    DeleteNoteIfExists("Note AIR2");

                //Remove the note using DeleteNote() funtion on InputConfigHlpr
                //This function takes the input as name of the note and returns the boolean value
                //ie, boolean value true means the note is deleted, false means given name of the note is not existing for this support
                //bIsDeleted = my_IJHgrInputConfigHlpr.DeleteNote("Note AIR2")
                if (note2 != null)
                    note2.Delete();

                //Get the Penetration Plate from the collection.
                BusinessObject pentrtPlatepart = (componentDictionary[HGRSUPHOLEPENTRTPLATE]).GetRelationship("madeFrom", "part").TargetObjects[0];

                //Determine if this is a Place By Structure Assembly
                //Get required information about the Bounding Box Surrounding the Pipe.
                //The Bounding Box used depends on the command.
                BoundingBox boundingBox;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    //Use the Structure Bounding Box
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supporting);
                }
                else
                {
                    //Use the Route Bouding Box
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                }

                Double width, height;
                width = boundingBox.Width;
                height = boundingBox.Height;

                String boundingBoxLow = boundingBox.LowReferencePortName;

                // String BBXWidth = BBX.;
                (componentDictionary[HGRSUPHOLEPENTRTPLATE]).SetPropertyValue(width, "IJUAHgrOccGeometry", "Width");
                (componentDictionary[HGRSUPHOLEPENTRTPLATE]).SetPropertyValue(height, "IJUAHgrOccGeometry", "Height");

                PropertyValueCodelist PartType = (PropertyValueCodelist)pentrtPlatepart.GetPropertyValue("IJExternalWeldedPipePart", "PartType");
                CodelistItem PartTypecodeList = PartType.PropertyInfo.CodeListInfo.GetCodelistItem((int)5);
                Matrix4X4 lcsBBox_Low = RefPortHelper.PortLCS(boundingBoxLow);

                (componentDictionary[HGRSUPHOLEPENTRTPLATE]).SetPropertyValue(lcsBBox_Low.GetIndexValue(12), "IJUAHgrOccPenPlate", "PipePort_X");
                (componentDictionary[HGRSUPHOLEPENTRTPLATE]).SetPropertyValue(lcsBBox_Low.GetIndexValue(13), "IJUAHgrOccPenPlate", "PipePort_Y");


                Double aspThick = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsHoleAspThickness", "Aspthickness")).PropValue;
                Double aspGap = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsHoleAspGap", "HoleAspGap")).PropValue;
                int aspDir = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAhsHoleAspDir", "AspDirection")).PropValue;

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDia = pipeInfo.OutsideDiameter;

                if (aspThick < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", PipeHgrAssemblies2Localizer.GetString(PipeHgrAssemblies2ResourceIDs.ErrInvalidThickness, "Invalid Thickness. Value should be greater than or equal to zero."), "", "HolePentrtPlate", 1);
                    return;
                }
                if (aspGap < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", PipeHgrAssemblies2Localizer.GetString(PipeHgrAssemblies2ResourceIDs.ErrInvalidGapvalue, "Invalid Gap. Value should be greater than or equal to zero."), "", "HolePentrtPlate", 1);
                    return;
                }


                (componentDictionary[HGRSUPHOLEPENTRTPLATE]).SetPropertyValue(aspThick, "IJUAHgrHoleAspectThk", "HoleAspThick");
                (componentDictionary[HGRSUPHOLEPENTRTPLATE]).SetPropertyValue(aspGap, "IJUAHgrHoleAspectThk", "HoleAspGap");
                (componentDictionary[HGRSUPHOLEPENTRTPLATE]).SetPropertyValue(aspDir, "IJUAHgrHoleAspectDir", "AspectDirect");

                support.SetPropertyValue(aspThick, "IJUAhsHoleAspThickness", "Aspthickness");
                support.SetPropertyValue(aspGap, "IJUAhsHoleAspGap", "HoleAspGap");


                //Set the part type on the penetration plate
                (componentDictionary[HGRSUPHOLEPENTRTPLATE]).SetPropertyValue(PartTypecodeList, "IJExternalWeldedPipePart", "PartType");

                // ====== ======
                //Create Joints
                //====== ======

                //----------------------------------------------------
                //Create the Joint between the Bounding Box Low and Struct Reference Ports
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreateRigidJoint("-1", "BBS_Low", HGRSUPHOLEPENTRTPLATE, "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                else
                    JointHelper.CreateRigidJoint("-1", "BBR_Low", HGRSUPHOLEPENTRTPLATE, "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Route Connections
        //----------------------------------------------------------------------------------- 
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();

                    //Gets the count of Routes
                    int numberOfRoutes = SupportHelper.SupportedObjects.Count;

                    for (int index = 1; index <= numberOfRoutes; index++)
                    {
                        //first Part return by GetAssemblyCatalogParts() is G4G_7878_01
                        //Connects to Route Input (i.e. the pipe)
                        routeConnections.Add(new ConnectionInfo(HGRSUPHOLEPENTRTPLATE, index));
                    }
                    //Return the collection of Route connection information.
                    return routeConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Struct Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportingConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Structure connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        //First Part return by GetAssemblyCatalogParts() is HgrSupPentrtPlate
                        //Connects to First Structure Input (i.e. the beam or plate)
                        //Add the PipeClampConnections to the Route Collection of Connections.
                        structConnections.Add(new ConnectionInfo(HGRSUPHOLEPENTRTPLATE, 1));
                    }

                    //Return the collection of Structure connection information.
                    return structConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        public override SupportingPortSelectionType SupportingPortType
        {
            get
            {
                return SupportingPortSelectionType.PortBeforeCut;
            }
        }
    }
}


