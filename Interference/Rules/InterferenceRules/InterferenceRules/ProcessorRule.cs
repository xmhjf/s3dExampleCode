using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Route.Middle;
using System.Collections;
using System.Collections.ObjectModel;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Interference
{
    /// <summary>
    /// This code is delivered to end user so that he can modify the same to change the default behavior of the Rules.
    /// ProcessorRule class inherits from InterferenceRule class and overrides methods to provide custom behavior for the rules
    /// </summary>
    public class ProcessorRule : InterferenceRule
    {
        /// <summary>
        /// Determines whether the reference file should be processed. 
        /// </summary>
        /// <param name="referenceType">The reference file type </param>
        /// <param name="referencePath">The reference file path as string </param>
        /// <returns> True for the reference file to be processed else returns false</returns>
        public override bool ProcessReference(Reference.SP3DReferenceFileType referenceType, string referencePath)
        {
            bool bProcessFile = true;
            if (referencePath == null)
            {
                throw new ArgumentNullException("reference path");
            }           
            //if (bProcessFile)
            //{
            //    int iReturn = referencePath.IndexOf("ifc_hvac1", StringComparison.OrdinalIgnoreCase);
            //    int iReturn2 = referencePath.IndexOf("ifc_pip1", StringComparison.OrdinalIgnoreCase);
            //    if ((iReturn >= 0) || (iReturn2 >= 0))
            //    {
            //        MiddleServiceProvider.ErrorLogger.Log("Ignore all elements in files ifc_hvac1 and ifc_pip1", 0);
            //        return false;
            //    }
            //    if (referenceType == Reference.SP3DReferenceFileType.AutoCAD)
            //    {
            //        return false;
            //    }
            //    if (referenceType == Reference.SP3DReferenceFileType.MicroStation)
            //    {
            //        return false;
            //    }
            //    if (referenceType == Reference.SP3DReferenceFileType.PDS)
            //    {
            //        return false;
            //    }
            //    if (referenceType == Reference.SP3DReferenceFileType.R3D)
            //    {
            //        return false;
            //    }
            //}
            return bProcessFile;

        }

        /// <summary>
        /// Determines whether the external reference object should be processed. 
        /// </summary>
        /// <param name="referenceObject">The InterferingObjectInfo object which should be processed </param>
        /// <returns> True for object to be processed and false for Object not to be processed</returns>
        public override bool ProcessReferenceObject(InterferingObjectInfo referenceObject)
        {
            if (referenceObject == null)
            {
                throw new ArgumentNullException("referenceObject");
            }
            bool bProcessReferenceObject = true;
            //if (bProcessReferenceObject)
            //{
            //    BOCInformation bocInfo = referenceObject.BOCinfo;
            //    string strObjectType = bocInfo.Name;
            //    if (strObjectType == "Reference3DStair")
            //    {
            //        return false;
            //    }
            //}

            return bProcessReferenceObject;
        }

        /// <summary>
        /// Checks to determine if the interference should be considered valid. This rule is called when a new interference
        /// is detected and when an existing interference is being modified (modify).
        /// 
        ///   1)  Create Interference (input existingClashBeingModified is false):
        ///       The rule gets triggered just after Interference detection detects an Interference and just before
        ///       persisting the same to Database. The main purpose of this rule is to decide whether user really
        ///       needs this Interference to be persisted in Database. For example, industry practice is not to report the
        ///       Insulation and insulation interferences when 2 pipes are connected by an elbow. This rule will also 
        ///       allow to put the user notes and user can choose his own status for the interference.
        ///       Once the user chooses this particular interference to be recorded in the database, he/she may also need to
        ///       assign the proper permission group for the same.  A means to control PGs is shown in the code below using the
        ///       m_strPermissionGroups variable.  This variable contains a list of permission groups in order of rank. 
        ///       While processing an interference, it is assigned to the lowest PG rank of interfering objects.
        ///       there by people who are incharge of the same will correct them accordingly.
        ///
        ///   2)  Modify Interference(input existingClashBeingModified is true):
        ///       The rule triggered when Interference detection process is trying to modify an interference,
        ///       because of modification of the parts participating in the collision. Default implementation
        ///       just changes the status and Notes of an interference.  User can see their properties and 
        ///       change / add /remove Notes and status
        ///       of the interference. If after processing the function decides that this interference is not
        ///       valid, this function should return False to tell the system to remove the existing interference.
        ///
        /// </summary>
        /// <param name="interferenceObject">The interference object whose validity is computed. </param>
        /// <param name="interferringObjectA">The InterferingObjectInfo object participating in the interference. </param>
        /// <param name="interferringObjectB">The InterferingObjectInfo object participating in the interference. </param>
        /// <param name="existingClashBeingModified">A boolean value to determine existing clash being modified </param>
        /// <returns> True for object to be processed and false for not to be processed</returns>
        public override bool IsValidInterference(Ingr.SP3D.Common.Middle.Interference interferenceObject, InterferingObjectInfo interferringObjectA,
            InterferingObjectInfo interferringObjectB, bool existingClashBeingModified)
        {
            if (interferenceObject == null)
            {
                throw new ArgumentNullException("interferenceObject");
            }
            if (interferringObjectA == null)
            {
                throw new ArgumentNullException("interferringObjectA");
            }
            if (interferringObjectB == null)
            {
                throw new ArgumentNullException("interferringObjectB");
            }

            bool bValidateMemberToMemberWithinTolerance = false;
            bool bValidateEntityThroughGratingSlab = false;
            bool bValidateGhostPermissionGroup = false;
            bool bValidateConnByIntermediateDuctConduitCableway = true;
            bool bValidateImportedToImported = true;
            bool bConnectByIntermediatePipeObject = true;
            bool bUserWantsToAddComment = false;

            BOCInformation interferenceObjectABOC = interferringObjectA.BOCinfo;
            BOCInformation interferenceObjectBBOC = interferringObjectB.BOCinfo;

            if (interferenceObjectABOC == null)
            {
                throw new ArgumentNullException("Failed to get interferenceObjectBBOC");
            }
            if (interferenceObjectBBOC == null)
            {
                throw new ArgumentNullException("Failed to get interferenceObjectBBOC");
            }

            string interferenceObjectABOCName = interferenceObjectABOC.Name;
            string interferenceObjectBBOCName = interferenceObjectBBOC.Name;
            
            if(interferenceObject.Type == InterferenceType.BadObject)
            {
                MiddleServiceProvider.ErrorLogger.Log("Processing did not happen as interference object type is bad : " + interferenceObject,0);
                return false;
            }

            if (bUserWantsToAddComment)
            { 
                string appendNotes = "";//If user adds any string here that will be appended to the IFC object. 
                if(interferenceObject.Type == InterferenceType.Severe)
                {
                    string existingRemark = interferenceObject.Remark;
                    interferenceObject.Remark = existingRemark + appendNotes;

                }
            }
            //Check if either object belongs to a permission group we want to avoid having interferences with
            if (bValidateGhostPermissionGroup)
            {
                //make sure first argument to objectBelongToPG is all UPPPERCASE
                if (ObjectBelongsToPG("GHOST", interferringObjectA, interferringObjectB))
                {
                    return false;
                }
            }

            //Check if two pipes are connect by a third piping related object; if so, no interference
            if (bConnectByIntermediatePipeObject)
            {
                if ((interferenceObjectABOCName == interferenceObjectBBOCName) && (interferenceObjectABOCName == "Pipes"))
                {
                    int iAspect;
                    int iAspect2;
                    interferenceObject.GetInterferingAspects(out iAspect, out iAspect2);
                    AspectID PipeObjectAspectId = AspectID.Insulation;
                    if (((iAspect == iAspect2) && (iAspect == (int)PipeObjectAspectId)) && AreObjectsConnectedByIntermediateObject(interferringObjectA, interferringObjectB))
                    {
                        return false;
                    }
                }
            }

            //Check if two Ducts/two Conduits/two CableTrays are connect by a common related object; if so, no interference
            if (bValidateConnByIntermediateDuctConduitCableway)
            {
                if (((interferenceObjectABOCName == interferenceObjectBBOCName) & 
                    (((interferenceObjectABOCName == "CablewayStraight") ||
                    (interferenceObjectABOCName == "CableTrays") ||
                    (interferenceObjectABOCName == "Ducts")) || 
                    (interferenceObjectABOCName == "Conduits"))) && 
                    AreObjectsConnectedByIntermediateObject(interferringObjectA, interferringObjectB))
                {
                    return false;
                }
            }

            //Check if a valid entity passes through a grating slab; if so, no interference
            if (bValidateEntityThroughGratingSlab)
            {
                if (IsGratingSlab(interferringObjectA))
                {
                    if (IsValidEntityThroughGratingSlab(interferringObjectA, interferringObjectB))
                    {
                        return false;
                    }
                }
                if (IsGratingSlab(interferringObjectB))
                {
                    if (IsValidEntityThroughGratingSlab(interferringObjectB, interferringObjectA))
                    {
                        return false;
                    }
                }
            }

            //check if beam to another object clash; if so and the interference is less than
            //tolerance" distance from the end of the beam, don't create the interference
            if (bValidateMemberToMemberWithinTolerance)
            {
                if ((interferenceObjectABOCName == "MemberPartLinear") || (interferenceObjectABOCName == "MemberPartCurve")) 
                {
                    if(IsMemberWithinTolerance(interferenceObject, interferringObjectA))
                    {
                        return false;
                    }
                }
                
                if ((interferenceObjectBBOCName == "MemberPartLinear") || (interferenceObjectBBOCName == "MemberPartCurve"))
                {
                    if(IsMemberWithinTolerance(interferenceObject, interferringObjectB))
                    {
                        return false;
                    }
                }
            }

            //if both objects involved are third party objects, don't create interference
            if (bValidateImportedToImported && AreBothImportedObjects(interferringObjectA, interferringObjectB))
            {
                return false;
            }

            if (!existingClashBeingModified)
            {
                if (!interferenceObject.IsLocallyDetected)
                {
                    switch (interferenceObject.Type)
                    {
                        case InterferenceType.Severe:
                            interferenceObject.Status = InterferenceStatus.MustResolve;
                            break;

                        case InterferenceType.Optional:
                            interferenceObject.Status = InterferenceStatus.NotReviewed;
                            break;

                        case InterferenceType.Clearance:
                            interferenceObject.Status = InterferenceStatus.Ignored;
                            break;

                        case InterferenceType.BadObject:
                            interferenceObject.Status = InterferenceStatus.MustResolve;
                            break;
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Is used to change the default behaviour of a class by making it use simple physical aspect 
        /// </summary>
        /// <param name="classInfo">BOC class information </param>
        /// <returns> True to use simple physical aspect and false for detailed physical aspect</returns>
        /// <remarks>Called when a class supports both Simple Physical and Detailed Physical aspects.  
        /// By default, Detailed Physical aspect will always be used for interference testing 
        /// but this method allows override of this behavior at the class level.</remarks>
        public override bool UseSimplePhysicalAspect(BOCInformation classInfo)
        {
            if (classInfo == null)
            {
                throw new ArgumentNullException("classInfo");
            }
            return false;
        }

        /// <summary>
        /// Is used to change the default behaviour of a class by making it use simple physical aspect 
        /// </summary>
        /// <param name="classInfo">BOC class information </param>
        /// <returns> True to use simple physical aspect and false for detailed physical aspect</returns>
        /// <remarks>Called when a class supports both Simple Physical and Detailed Physical aspects.  
        /// By default, Detailed Physical aspect will always be used for interference testing 
        /// but this method allows override of this behavior at the class level.
        /// Enabling Simple and Detailed physical aspect for all objects (or certain object types) will
        /// have a negative impact on IFC performance. By default this rule is disabled.</remarks>
        public override bool UseSimpleAndDetailedPhysicalAspect(BOCInformation classInfo)
        {
            if (classInfo == null)
            {
                throw new ArgumentNullException("classInfo");
            }
            //if (classInfo.Name == "SmartEquipment")
                //return true;

            return false;
        }

        /// <summary>
        /// Checks whether the s3dobject should be processed for interference. 
        /// This Rule will be called by the Interference engine after it updates the Range in the spatial index. The object 
        /// under processing is sent as an argument and this Rule returns true or false. Based on this the object is either 
        /// ignored or considered further.  
        /// 
        /// There are examples on how to skip objects based on their object type and range extents.
        /// </summary>
        /// <param name="objectA">InterferingObjectInfo object which needs to be processed for interference</param>
        /// <returns> True for processing the object and false for not processing the object</returns>  
        public override bool ProcessObject(InterferingObjectInfo objectA)
        {
            if (objectA == null)
            {
                throw new ArgumentNullException("objectA");
            }

            //Range limit value
            float rangeDiagonalLenLimit = 100; //in meters
            bool bRangeCheckEnabled = false;
            //Sample on how to ignore objects based on the type of object. The code below
            //ignores Pipe objects. Change the BOCinfo.Name to ignore some other
            //objects. The names have to be an exact match

            //if (objectA.BOCinfo.Name == "Pipes")
            //{
            //    return false;
            //}

            //   Sample on how to ignore objects based on their range extents. This should be enabled if you
            //   insert very large SAT files in your model and want IFC to work while you are breaking those large SAT
            //   files into smaller SAT or Designed Equipment placeholders.
            //   This is necessary if you want IFC to complete, but has inserted large objects from third party package
            //   such as large SAT file representing the duct work for the whole plant as a single object.
            //   The reason is that every processed object will have a range that will interfere with this large object,
            //   making IFC processing very ineficient, sometimes preventing it from completing in a reasonable amount of time.
            if (bRangeCheckEnabled)
            {
                BusinessObject s3dBO = objectA.BusinessObject;

                IRange range = s3dBO as IRange;
                RangeBox box = range.Range;
                if (box != null)
                {
                    Position rangeHigh;
                    Position rangeLow;
                    rangeHigh = box.High;
                    rangeLow = box.Low;
                    if (rangeHigh.DistanceToPoint(rangeLow) > rangeDiagonalLenLimit)
                    {
                        MiddleServiceProvider.ErrorLogger.Log("not Processed object due to range diagonal length greater than threshold. Object id is : " + objectA.BusinessObject.ObjectID, 0);
                        return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Checks for valid entity through grating slabs 
        /// </summary>
        /// <param name="objectA">InterferingObjectInfo object which needs to be a valid entity through grating slab</param>
        /// <param name="objectB">InterferingObjectInfo object which needs to be a valid entity through grating slab</param>
        /// <returns> True for valid entity and false for invalid entity</returns>
        private bool IsValidEntityThroughGratingSlab(InterferingObjectInfo objectA, InterferingObjectInfo objectB)
        {
            if (objectA == null) 
            {
                throw new ArgumentNullException("objectA");
            }

            if (objectB == null)
            {
                throw new ArgumentNullException("objectB");
            }
            //this function validates these types of entities clashing with a grating slab
            //1 - pipes less than 3 inches in diameter
            //2 - members of type "column" and "brace"
            //3 - all handrails
            BOCInformation ObjectABOCInfo = objectA.BOCinfo;
            BOCInformation ObjectBBOCInfo = objectB.BOCinfo;
            if ((ObjectABOCInfo != null) && (ObjectBBOCInfo != null))
            {
                if (ObjectBBOCInfo.Name == "Pipes")
                {
                    //pipes less than 3 inches in diameter
                    RoutePart part = objectB.BusinessObject as RoutePart;
                    if (part != null)
                    {
                        ReadOnlyCollection<RouteFeature> features = part.Features;
                        if ((features != null) && (features.Count > 0))
                        {
                            IPipePathFeature feature = features[0] as IPipePathFeature;
                            if (feature != null)
                            {
                                NominalDiameter featureNPD = feature.NPD;
                                if (featureNPD != null)
                                {
                                    if ((featureNPD.Size < 3.0) && (featureNPD.Units == "in"))
                                    {
                                        return true;
                                    }
                                    if ((featureNPD.Size < 75.0) && (featureNPD.Units == "mm"))
                                    {
                                        return true;
                                    }
                                }
                            }
                        }
                    }

                }

                //if (ObjectBBOCInfo.Name == "MemberPart")
                //{
                //    MemberPart objMemPart = objectB.BusinessObject as MemberPart;
                //    if (objMemPart.TypeCategory == 2)
                //    {
                //        return true;
                //    }
                //    if (objMemPart.TypeCategory == 3)
                //    {
                //        return true;
                //    }
                //    return false;
                //}

                //for handrails
                if (ObjectBBOCInfo.Name == "Handrails")
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Checks whether object belongs to a perticular permission group
        /// </summary>
        /// <param name="string">Permission group name</param>
        /// <param name="objectA">first InterferingObjectInfo object whose permission group has to checked</param>
        /// <param name="objectB">second InterferingObjectInfo object whose permission group has to checked</param>
        /// <returns> if either input object belongs to the permission group else returns false</returns>
        private bool ObjectBelongsToPG(string pgname, InterferingObjectInfo objectA, InterferingObjectInfo objectB)
        {
            if (objectA == null)
            {
                throw new ArgumentNullException("objectA");
            }
            if (pgname == null)
            {
                throw new ArgumentNullException("pgname");
            }
            if (objectB == null)
            {
                throw new ArgumentNullException("objectB");
            }
            BusinessObject ObjectABO = objectA.BusinessObject;
            if (ObjectABO != null)
            {
                PermissionGroup ObjectABOPermissionGroup = ObjectABO.PermissionGroup;
                if (ObjectABOPermissionGroup != null)
                {
                    return ((ObjectABOPermissionGroup.Name.Equals(pgname, StringComparison.OrdinalIgnoreCase)) || (objectB.BusinessObject.PermissionGroup.Name.Equals(pgname, StringComparison.OrdinalIgnoreCase)));
                }
            }
            return false;
        }

        /// <summary>
        /// Checks whether the objects are imported from another application but not sp3d
        /// </summary>
        /// <param name="objectA">InterferingObjectInfo object whose origin needs to be evaluated</param>
        /// <param name="objectB">InterferingObjectInfo object whose origin needs to be evaluated</param>
        /// <returns> True for imported objects and false in case of sp3d objects</returns>
        private bool AreBothImportedObjects(InterferingObjectInfo objectA, InterferingObjectInfo objectB)
        {
            if (objectA == null)
            {
                throw new ArgumentNullException("objectA");
            }
            if (objectB == null)
            {
                throw new ArgumentNullException("objectB");
            }

            BusinessObject objectABO = objectA.BusinessObject;
            BusinessObject objectBBO = objectB.BusinessObject;

            if ((objectABO != null) && (objectABO != null))
            {
                if (objectABO.SupportsInterface("IJImportedStructureItem"))
                {
                    if (objectBBO.SupportsInterface("IJImportedStructureItem"))
                    {
                        return true;
                    }
                }

                if (objectABO.SupportsInterface("IJCISData"))
                {
                    if (objectBBO.SupportsInterface("IJCISData"))
                    {
                        //If originAppId is set for both objects then they are originated in another application
                        //for Objects placed in SP3D the OriginAppId is empty

                        PropertyValue objPropertyValue = objectABO.GetPropertyValue("IJCISData", "OriginAppId");
                        PropertyValue objPropertyValue2 = objectBBO.GetPropertyValue("IJCISData", "OriginAppId");

                        PropertyValueString objPropertyValueStringA = (PropertyValueString)objPropertyValue;
                        PropertyValueString objPropertyValueStringB = (PropertyValueString)objPropertyValue;
                        if (objPropertyValueStringA.PropValue != null && objPropertyValueStringB.PropValue != null)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Checks whether interference is within tolerance limit of either end of the input beam.
        /// </summary>
        /// <param name="Interference">Interference object</param>
        /// <param name="objectA">InterferingObjectInfo object participating in the interference</param>
        /// <returns> True for within tolerance limit else returns false</returns>
        private bool IsMemberWithinTolerance(Ingr.SP3D.Common.Middle.Interference interferenceObject, InterferingObjectInfo objectA)
        {
            if (interferenceObject == null)
            {
                throw new ArgumentNullException("interferenceObject");
            }
            if (objectA == null)
            {
                throw new ArgumentNullException("objectA");
            }

            Position positionA;
            Position positionB;
            double toleranceInMeters = 0.254;

            BusinessObject objectABO = objectA.BusinessObject;
            if (objectABO != null)
            {
                MemberPart objectAMemberPart = objectA.BusinessObject as MemberPart;
                if (objectAMemberPart != null)
                {
                    objectAMemberPart.GetEndPoints(out positionA, out positionB);

                    Position positionInterference = interferenceObject.Location;

                    if (((positionA.DistanceToPoint(positionInterference)) < toleranceInMeters) || ((positionB.DistanceToPoint(positionInterference)) < toleranceInMeters))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Checks whether the object is grating slab
        /// </summary>
        /// <param name="interferenceObject">InterferingObjectInfo object which needs to be evaluated</param>
        /// <returns> True for grating slab object and false in case of not a grating slab</returns>
        private bool IsGratingSlab(InterferingObjectInfo interferenceObject)
        {
            BOCInformation interferenceObjectBOCinfo = interferenceObject.BOCinfo;
            if (interferenceObjectBOCinfo != null)
            {
                if (interferenceObjectBOCinfo.Name == "Slab")
                {
                    Slab findGrating = interferenceObject.BusinessObject as Slab;
                    if (findGrating != null)
                    {
                        Part findGratingType = findGrating.Type;
                        if (findGratingType != null)
                        {
                            string sPartNumber = findGratingType.PartNumber;
                            int n = sPartNumber.IndexOf("Grating", StringComparison.OrdinalIgnoreCase);
                            if (n >= 0)
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }
    }
}