using Ingr.SP3D.Common.Middle;
using System;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using System.Collections.Generic;

namespace Ingr.SP3D.Content.Planning
{
    public class DefaultGroupNameRule : NameRuleBase
    {
        public override void ComputeName(BusinessObject oEntity, ReadOnlyCollection<BusinessObject> oParents, BusinessObject oActiveEntity)
        {
            try
            {
                CommonPartsGroup commonPartGroup = oEntity as CommonPartsGroup;
                CommonPartsManager commonPartManager;
                MemberPart memberPart;
                BusinessObject commonPart;
                ReadOnlyCollection<ICommonEntity> folderCollection = null;
                long numOfgroups = 0;
                long processControl;
                long count = 0;              
                string entityName = string.Empty;              
                string xSectionName;  
            
                if (commonPartGroup == null)
                {
                    return;
                }

                //Get the first object under CPG
               processControl = commonPartGroup.ProcessControlType;
               ReadOnlyCollection<BusinessObject> commonParts = commonPartGroup.Parts;

               if (processControl == 2)
               {
                   //need to check count
                   commonPartManager = commonPartGroup.Manager;
                   if (commonPartManager != null)
                   {
                       folderCollection = commonPartManager.CommonEntityChildren;
                   }
               }

                if (commonParts != null)
                {
                    if (commonParts.Count > 0)
                    {
                        commonPart = commonParts[0];

                        if ((commonPartGroup.ProcessPurpose == (int)ProcessPurpose.StandardByXML 
                            || commonPartGroup.ProcessPurpose == (int)ProcessPurpose.StandardByModel) 
                            && !string.IsNullOrEmpty(commonPartGroup.StandardReferenceEntityName))
                        {
                            oEntity.SetPropertyValue(commonPartGroup.StandardReferenceEntityName, "IJNamedItem", "Name");
                        }
                        else if (commonPart is PlatePart && (!(commonPart is CollarPart)) && ((IsBracket(commonPart) == false)))
                        {
                            entityName = GetGroupNameFromHierarchy(commonPart);

                            //if CPG is Manually Created add the index
                            if (processControl == 2)
                            {
                                if (folderCollection != null)
                                {
                                    numOfgroups = GetNumberofGroups(folderCollection.Count, folderCollection);
                                    string tempName = entityName + "_" + "UserDefined_" + numOfgroups;
                                    oEntity.SetPropertyValue(tempName, "IJNamedItem", "Name");
                                }                                
                            }
                            else
                            {
                                oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                            }
                        }
                        else if (commonPart is ProfilePart  && (typeof(ProfilePart) == commonPart.GetType()) )
                        {
                            //Using the same logic as of the PlatePart Group
                            entityName = GetGroupNameFromHierarchy(commonPart);

                            //if CPG is Manually Created add the index
                            if (processControl == 2)
                            {
                                if (folderCollection != null)
                                {
                                    numOfgroups = GetNumberofGroups(folderCollection.Count, folderCollection);
                                    string tempName = entityName + "_" + "UserDefined_" + numOfgroups;
                                    oEntity.SetPropertyValue(tempName, "IJNamedItem", "Name");
                                }
                            }
                            else
                            {
                                oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                            }
                        }
                        else if (commonPart is CollarPart)
                        {
                            if (IsStandardCollar(commonPart) == true)
                            {
                                entityName = GetCollarGroupName(commonPart);
                            }
                             else
                            {
                                entityName = GetGroupNameFromHierarchy(commonPart);
                            }

                            //if CPG is Manually Created add the index
                            if (processControl == 2)
                            {
                                if (folderCollection != null)
                                {
                                    numOfgroups = GetNumberofGroups(folderCollection.Count, folderCollection);
                                    string tempName = entityName + "_" + "UserDefined_" + numOfgroups;
                                    oEntity.SetPropertyValue(tempName, "IJNamedItem", "Name");
                                }
                            }
                            else
                            {
                                oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                            }
                        }
                        else if (IsBracket(commonPart) ==true)
                        {
                            if (IsStandardBracket(commonPart) ==true)
                            {
                                entityName = GetBracketGroupName(commonPart);
                            }
                            else
                            {
                                entityName = GetGroupNameFromHierarchy(commonPart);
                            }

                            //if CPG is Manually Created add the index
                            if (processControl == 2)
                            {
                                if (folderCollection != null)
                                {
                                    numOfgroups = GetNumberofGroups(folderCollection.Count, folderCollection);
                                    string tempName = entityName + "_" + "UserDefined_" + numOfgroups;
                                    oEntity.SetPropertyValue(tempName, "IJNamedItem", "Name");
                                }
                            }
                            else
                            {
                                oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                            }
                        }
                        else if (commonPart is MemberPart && (typeof(MemberPart) == commonPart.GetType()))
                        {
                            memberPart = commonPart as MemberPart;
                            CrossSection crossSection = memberPart.CrossSection;
                            xSectionName = crossSection.Name;
                            count = GetNameCount(commonPart);
                            entityName = xSectionName + "-0" + count;

                            //if CPG is Manually Created add the index
                            if (processControl == 2)
                            {
                                if (folderCollection != null)
                                {
                                    numOfgroups = GetNumberofGroups(folderCollection.Count, folderCollection);
                                    string tempName = entityName + "_" + "UserDefined_" + numOfgroups;
                                    oEntity.SetPropertyValue(tempName, "IJNamedItem", "Name");
                                }
                            }
                            else
                            {
                                oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                            }
                        }
                        else if (commonPart is AssemblyBase)
                        {
                            AssemblyBase assemblyBase = (AssemblyBase)commonPart;
                            count =GetNameCount(commonPart);
                            entityName = assemblyBase.Name + "-0" + count;

                            //if CPG is Manually Created add the index
                            if (processControl == 2)
                            {
                                if (folderCollection != null)
                                {
                                    numOfgroups = GetNumberofGroups(folderCollection.Count, folderCollection);
                                    string tempName = entityName + "_" + "UserDefined_" + numOfgroups;
                                    oEntity.SetPropertyValue(tempName, "IJNamedItem", "Name");
                                }
                            }
                            else
                            {
                                oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                            }
                        }

                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.ComputeName: Error encountered (" + e.Message + ")");
            }   
        }

        public override Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            return new Collection<BusinessObject>();
        }

        private string GetGroupNameFromHierarchy(BusinessObject commonPart)
        {
            string groupName = null;
            try
            {
                IAssemblyChild assemblyChild;
                IAssembly[] parentAssy = new IAssembly[3];
                int i;
                long count = 0;

                assemblyChild = (IAssemblyChild)commonPart;

                if (assemblyChild == null)
                {
                    return groupName;
                }

                for (i = 0; i <= 3; i++)
                {
                    if (assemblyChild.AssemblyParent != null)
                    {
                        if (assemblyChild.AssemblyParent is AssemblyBase)
                        {
                            parentAssy[i] = assemblyChild.AssemblyParent;
                        }
                        else
                        {
                            //parent is config root.
                            break;
                        }

                        if (parentAssy[i] is Block)
                        {
                            break;
                        }

                        if (i == 3)
                        {
                            break;
                        }
                    }
                    assemblyChild = (IAssemblyChild)assemblyChild.AssemblyParent;
                }

                for (int j = i; j >= 0; j--)
                {
                    if (parentAssy[j] != null)
                    {
                        if (groupName == null)
                        {
                            groupName = Convert.ToString(((BusinessObject)parentAssy[j]).GetPropertyValue("IJNamedItem", "Name"));
                        }
                        else
                        {
                            groupName = groupName + "-" +Convert.ToString(((BusinessObject)parentAssy[j]).GetPropertyValue("IJNamedItem", "Name"));
                        }
                    }
                }

                for (int j = 1; j <= (3 - i); j++)
                {
                    groupName += "-0";
                }

                count = GetNameCount(commonPart);
                groupName += "-0" + count;

            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.GetGroupNameFromHierarchy: Error encountered (" + e.Message + ")");
            }   
            return groupName;
        }

        private long GetNameCount(BusinessObject srcObject)
        {
            long count = 0;
            try
            {
                string locationId;
                string[] delimiter = { " " };
                string srcObjectName = GetTypeString(srcObject);                 
                string partName = string.Join(" ", srcObjectName.Split(delimiter, StringSplitOptions.None));                

                GetCountAndLocationID(partName, out count, out locationId);
                count += 1;
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.GetNameCount: Error encountered (" + e.Message + ")");
            }   

            return count;
        }

        private string GetBracketGroupName(BusinessObject bracketpart)
        {
            string bracketGrupName = string.Empty;

            try
            {
                CustomPlatePart customPlatePart = bracketpart as CustomPlatePart;

                //for structDetailing bracket
                if (customPlatePart != null && customPlatePart.CustomPlateType == CustomPlateType.Bracket && customPlatePart.Type == PlateType.BracketPlate)
                {
                    bracketGrupName = customPlatePart.PartName;
                }
                else  //for moldedForm bracket
                {
                    bracketGrupName = GetBracketItemName(bracketpart);
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.GetBracketGroupName: Error encountered (" + e.Message + ")");
            }
            return bracketGrupName;
        }

        private bool IsBracket(BusinessObject entity)
        {          
            try
            {
                CustomPlatePart custmPlatePart = entity as CustomPlatePart;
                BracketPlateSystem bracketSystem = null;

                if (entity is PlatePart) /// for MoldedForm bracket
                {
                    PlatePart  tempPlatePart = (PlatePart)entity;                    
                    bracketSystem =(BracketPlateSystem) tempPlatePart.RootPlateSystem ;

                    if (bracketSystem != null)
                    {
                        if (bracketSystem.PlaneDefinition != null || bracketSystem.PlaneDefinition == null)
                        {
                            return true;
                        }
                    }
                }
                else if (custmPlatePart != null) // for structDetailing bracket
                {
                    if (custmPlatePart.Type ==  PlateType.BracketPlate && custmPlatePart.CustomPlateType == CustomPlateType.Bracket )
                    {
                        return true;
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.Isbracket: Error encountered (" + e.Message + ")");
            }
            return false;

        }

        private string GetBracketItemName(BusinessObject entity)
        {
            string bracketItemName = string.Empty;

            try
            {
                PlatePart platepart = entity as PlatePart; // for moldedForm Bracket
                if (platepart != null)
                {
                    BracketPlateSystem bracketPlateSys = platepart.RootPlateSystem as BracketPlateSystem;
                    if (bracketPlateSys != null)
                    {
                        bracketItemName = bracketPlateSys.PartName;
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.GetBracketItemName: Error encountered (" + e.Message + ")");
            }

            return bracketItemName;
        }      

        private void GetAttribute(BusinessObject entity, string AttributeName, out bool attrfound, out string longStringValue)
        {
            longStringValue = null;
            attrfound = false;         
            try
            {
                ReadOnlyDictionary<InterfaceInformation> interfacesInfo = entity.ClassInfo.Interfaces;
                if (interfacesInfo != null)
                {
                    foreach (var keyValuePair in interfacesInfo)
                    {
                        InterfaceInformation interfaceInfo = keyValuePair.Value;

                        if (interfaceInfo != null)
                        {
                            PropertyInformation propertyInfo = interfaceInfo.GetPropertyInfo(AttributeName);

                            if (propertyInfo != null)
                            {
                                PropertyValue  propertyValue = entity.GetPropertyValue(propertyInfo);                              
                                List<CodelistItem> codeListMembers = propertyInfo.CodeListInfo.CodelistMembers;                              

                                if (codeListMembers != null)
                                {
                                    foreach (var item in codeListMembers)
                                    {
                                        string PropertyName=Convert.ToString(propertyValue);
                                        if (string.Compare(item.ShortDisplayName,PropertyName) == 0)
                                        {
                                            CodelistItem requiredCodelistItem = item;                                            
                                            longStringValue = requiredCodelistItem.DisplayName;
                                            attrfound = true;
                                            return ;
                                        }
                                    }
                                } 
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.GetAttribute: Error encountered (" + e.Message + ")");
            }
            
        }

        private string GetBracketClassName(BusinessObject entity)
        {
            string bracketClassname = string.Empty;
            try
            {
               bool attributeFound = false;
               string AttributeName = "BracketByPlaneType";
               string longStringValue = string.Empty;
               PlatePart platePart = entity as PlatePart;

               if (platePart != null)
               {
                   BracketPlateSystem bracketPlateSys = platePart.RootPlateSystem as BracketPlateSystem;  //for MoldedForm Bracket only.

                   if (bracketPlateSys != null)
                        GetAttribute(bracketPlateSys, AttributeName, out attributeFound, out longStringValue);

                   if (attributeFound)
                   {
                       bracketClassname = longStringValue;
                   }
               }
                
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.GetBracketClassName: Error encountered (" + e.Message + ")");
            }

            return bracketClassname;
        }

        private bool IsStandardBracket(BusinessObject BracketPart)
        {
            bool standardBracket = false;

            try
            {
                CustomPlatePart customplatepart = BracketPart as CustomPlatePart;             
                PlatePart  platePart = BracketPart as PlatePart;

                if (customplatepart != null)  //StructDetailing bracket
                {
                    PartClass partClass = customplatepart.Part.PartClass as PartClass;
                    if (partClass != null)
                    {
                        string bracketClassName = partClass.PartClassName;

                        if (string.Compare(bracketClassName, "2SBracketLinear") == 0)
                        {
                            standardBracket = true;
                        }
                        else
                        {
                            standardBracket = false;
                        }
                    }
                }
                else
                {
                    string bracketClassName = GetBracketClassName(BracketPart); //moldedForm bracket
                    if (string.Compare(bracketClassName, "2SLT") == 0)
                    {
                        standardBracket = true;
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.IsStandardBracket: Error encountered (" + e.Message + ")");
            }

            return standardBracket;
        }

        private bool IsWebBaseAngle90(BusinessObject collar)
        {
            bool WebBaseAngle90 = false;

            try
            {
                CustomPlatePart customPlatepart = collar as CustomPlatePart;              
                CollarPart collartPart = collar as CollarPart;
                Vector basePlateNormal = null;
                BusinessObject penetratingObj = null, penetratedObj = null;

                if (collartPart != null)
                {
                    Position penetrationLocation = collartPart.Slot.Position;
                    TopologyPort basePlatePort = collartPart.BasePlatePort;
                    basePlateNormal = basePlatePort.Normal;
                    Feature slotFeature = collartPart.Slot;
                    slotFeature.GetInputs(out penetratedObj, out penetratingObj);
                }          

                TopologyPort penetratedPort = null, penetratingPort = null;
                Vector penetratedNormal=null, webNormal = null;

                if (penetratedObj != null)
                {
                    if (penetratedObj is PlatePart)
                    {
                        PlatePartBase platePartbase = penetratedObj as PlatePartBase;

                        if (platePartbase != null)
                        {
                            if (platePartbase.MoldedSide == ContextTypes.Base)
                            {
                                penetratedPort = platePartbase.GetPort(TopologyGeometryType.Face, ContextTypes.Base);
                            }
                            else if (platePartbase.MoldedSide == ContextTypes.Offset)
                            {
                                penetratedPort = platePartbase.GetPort(TopologyGeometryType.Face, ContextTypes.Offset);
                            }
                        }
                    }
                    else
                    {
                        ProfilePart tempProfilePart = penetratedObj as ProfilePart;

                        if (tempProfilePart != null)
                        {
                            ReadOnlyCollection<IPort> tempPorts = tempProfilePart.GetPorts(PortType.All);

                            foreach (IPort port in tempPorts)
                            {
                                TopologyPort tempPort = port as TopologyPort;

                                if (tempPort != null)
                                {
                                    if (tempPort.SectionId == (int)SectionFaceType.Web_Left)
                                    {
                                        penetratedPort = tempPort;
                                    }
                                }
                            }
                        }
                    }
                }

                if (penetratingObj != null)
                {
                    ProfilePart profilePart = penetratingObj as ProfilePart;

                    if (profilePart != null)
                    {
                        ReadOnlyCollection<IPort> ports = profilePart.GetPorts(PortType.All);

                        foreach (IPort port in ports)
                        {
                            TopologyPort tempPort = port as TopologyPort;

                            if (tempPort != null)
                            {
                                if (tempPort.SectionId == (int)SectionFaceType.Web_Right)
                                {
                                    penetratingPort = tempPort;
                                    break;
                                }
                            }
                        }
                    }
                }

                if (penetratedPort != null)
                {
                    penetratedNormal = penetratedPort.Normal;
                }

                if (penetratingPort != null)
                {
                    webNormal = penetratingPort.Normal;
                }

                double dotResult = 0.0;

                if (penetratedNormal != null && basePlateNormal != null && webNormal != null)
                {
                    Vector penetratedCrossBase = penetratedNormal.Cross(basePlateNormal);
                    Vector penetratedCrossWeb = penetratedNormal.Cross(webNormal);

                    penetratedCrossBase.Length = 1;
                    penetratedCrossWeb.Length = 1;

                    dotResult = penetratedCrossBase.Dot(penetratedCrossWeb);
                }       

                if (dotResult < -1)
                {
                    dotResult = -1;
                }
                else if (dotResult > 1)
                {
                    dotResult = 1;
                }

                if (Math.Abs(dotResult ) < 0.000001)
                {
                    WebBaseAngle90 = true;
                }
                else
                {
                    WebBaseAngle90 = false;
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.IsWebBaseAngle90: Error encountered (" + e.Message + ")");
            }

            return WebBaseAngle90;
        }

        private ReadOnlyCollection<IConnection> GetConnectionsData(IConnectable connectable, Type connectionType)
        {
            Collection<IConnection> connections = new Collection<IConnection>();           
            try
            {
                if (connectable == null)
                {
                    throw new ArgumentNullException("connectable");
                }

                if (!(connectionType == typeof(AssemblyConnection) || connectionType == typeof(PhysicalConnection) || connectionType == typeof(LogicalConnection)))
                {
                    return new ReadOnlyCollection<IConnection>(connections); ;
                }

                ReadOnlyCollection<IPort> connectedPorts = connectable.GetConnectedPorts(Ingr.SP3D.Common.Middle.PortType.All);

                foreach (IPort connectedPort in connectedPorts)
                {
                    ReadOnlyCollection<IConnection> TempConnections = connectedPort.Connections;
                    if (TempConnections != null)
                    {
                        foreach (IConnection connection in connections)
                        {
                            connections.Add(connection);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.GetConnectionsData: Error encountered (" + e.Message + ")");
            }
            return new ReadOnlyCollection<IConnection>(connections);
        }

        private string GetCollarGroupName(BusinessObject collarPart)
        {
            string collargroupname = string.Empty;
            try
            {
                string commonPartname = string.Empty;
                dynamic platePartbase;
                CollarPart tempcollarpart;
                string rootClassName = string.Empty;
                BusinessObject penetratingObject;
                string grade;

                if (collarPart is CollarPart)
                {
                    tempcollarpart = collarPart as CollarPart;
                    commonPartname = tempcollarpart.PartName;
                    platePartbase = collarPart as PlatePartBase;

                    if (platePartbase != null)
                    {
                        commonPartname = commonPartname + "_" + Convert.ToString(Math.Round(platePartbase.Thickness * 1000, 1));
                    }

                    grade = tempcollarpart.MaterialGrade;

                    if (string.Compare(grade, "A") != 0)
                    {
                        commonPartname = commonPartname + "'" + tempcollarpart.MaterialGrade + "'";
                    }

                    penetratingObject = tempcollarpart.PenetratingObject;
                    commonPartname = commonPartname + "(" + Convert.ToString(penetratingObject.GetPropertyValue("IJNamedItem", "Name")) + ")";
                    collargroupname = commonPartname;
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.GetCollarGroupName: Error encountered (" + e.Message + ")");
            }

            return collargroupname;
        }

        private long GetNumberofGroups(long folderCount, ReadOnlyCollection<ICommonEntity> folderCollection)
        {
            long numOfGrups = 0;

            try
            {
                foreach (var folder in folderCollection)
                {
                    CommonPartsFolder cmnPartFolder = folder as CommonPartsFolder;

                    if (cmnPartFolder != null)
                    {
                        ReadOnlyCollection<ICommonEntity> commonEntityChildren = cmnPartFolder.CommonEntityChildren;

                        if (commonEntityChildren != null && commonEntityChildren.Count > 0)
                        {
                            numOfGrups = numOfGrups + commonEntityChildren.Count;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.NumberofGroups: Error encountered (" + e.Message + ")");
            }

            return numOfGrups;
        }

        private void UpdateAlphabet(string Bracketname, string LastAlphabet)
        {
            string lastChar = Bracketname.Substring(Bracketname.Length - 1, 1);

            int asciiOfLastChar = Convert.ToInt32(lastChar);

            if (asciiOfLastChar >= 65 && asciiOfLastChar <= 90)
            {
                int asciiOfLastAlphabet = Convert.ToInt32(LastAlphabet);

                if (asciiOfLastChar >= asciiOfLastAlphabet)
                {
                    LastAlphabet = ((char)(asciiOfLastChar + 1)).ToString();
                }
            }

        }

        private ProfilePart GetConnectedProfilePart(BusinessObject Bracket, bool smartPlate)
        {
            ProfilePart connectedProfilePart = null;
            try
            {
                System.Collections.Generic.List<BusinessObject> profilepartColl1 = new System.Collections.Generic.List<BusinessObject>();
                System.Collections.Generic.List<BusinessObject> profilepartColl2 = new System.Collections.Generic.List<BusinessObject>();

                if (smartPlate)
                {
                    CustomPlatePart customPlatePart = Bracket as CustomPlatePart;

                    if (customPlatePart != null && customPlatePart.CustomPlateType == CustomPlateType.Bracket)
                    {
                        System.Collections.Generic.List<BusinessObject> bracketSupports = customPlatePart.Supports;

                        if (bracketSupports[0] is ProfilePart)
                        {
                            ProfilePart profilePart = bracketSupports[0] as ProfilePart;
                            if (profilePart != null)
                            {
                                ReadOnlyCollection<IPort> profilePartConnectables = profilePart.GetConnectablePorts(PortType.All);
                                BusinessObject tempConnetable = (BusinessObject)profilePartConnectables[0];
                                profilepartColl1.Add(tempConnetable);
                            }                            
                        }
                        else if (bracketSupports[1] is ProfilePart)
                        {
                            ProfilePart profilePart = bracketSupports[1] as ProfilePart;
                            if (profilePart != null)
                            {
                                ReadOnlyCollection<IPort> profilePartConnectables = profilePart.GetConnectablePorts(PortType.All);
                                BusinessObject tempConnetable = (BusinessObject)profilePartConnectables[0];
                                profilepartColl1.Add(tempConnetable);
                            }
                        }
                    }
                    else
                    {
                        PlatePart bracketPlatePart;
                        BracketPlateSystem bracketPlateSys;
                        ISystem profileSystem = null;

                        bracketPlatePart = Bracket as PlatePart;

                        if (bracketPlatePart != null)
                        {
                            bracketPlateSys = bracketPlatePart.SystemParent as BracketPlateSystem;

                            //can be know from the PlaneDefinition property on BracketPlateSystem. 
                            //If the PD is not  null then the bracket is bracket by plane
                            if (bracketPlateSys != null && bracketPlateSys.PlaneDefinition != null)
                            {
                                BracketSupportDefinition brcktTempSuportDefinition = bracketPlateSys.SupportsOrientationDefinition.FirstSupportDefinition;
                                if (brcktTempSuportDefinition.Support is Profile)
                                {
                                    profileSystem = (ISystem)brcktTempSuportDefinition.Support;
                                }

                                brcktTempSuportDefinition = bracketPlateSys.SupportsOrientationDefinition.SecondSupportDefinition;
                                if (brcktTempSuportDefinition.Support is Profile)
                                {
                                    profileSystem = (ISystem)brcktTempSuportDefinition.Support;
                                }

                                brcktTempSuportDefinition = bracketPlateSys.SupportsOrientationDefinition.ThirdSupportDefinition;
                                if (brcktTempSuportDefinition.Support is Profile)
                                {
                                    profileSystem = (ISystem)brcktTempSuportDefinition.Support;
                                }

                                brcktTempSuportDefinition = bracketPlateSys.SupportsOrientationDefinition.FourthSupportDefinition;
                                if (brcktTempSuportDefinition.Support is Profile)
                                {
                                    profileSystem = (ISystem)brcktTempSuportDefinition.Support;
                                }

                                brcktTempSuportDefinition = bracketPlateSys.SupportsOrientationDefinition.FifthSupportDefinition;
                                if (brcktTempSuportDefinition.Support is Profile)
                                {
                                    profileSystem = (ISystem)brcktTempSuportDefinition.Support;
                                }

                            }
                            else if (bracketPlateSys != null && bracketPlateSys.PlaneDefinition == null)
                            {
                                BracketSupportDefinition brcktTempSuportDefinition = bracketPlateSys.SupportsOrientationDefinition.FirstSupportDefinition;
                                if (brcktTempSuportDefinition.Support is Profile)
                                {
                                    profileSystem = (ISystem)brcktTempSuportDefinition.Support;
                                }

                                brcktTempSuportDefinition = bracketPlateSys.SupportsOrientationDefinition.SecondSupportDefinition;
                                if (brcktTempSuportDefinition.Support is Profile)
                                {
                                    profileSystem = (ISystem)brcktTempSuportDefinition.Support;
                                }

                                brcktTempSuportDefinition = bracketPlateSys.SupportsOrientationDefinition.ThirdSupportDefinition;
                                if (brcktTempSuportDefinition.Support is Profile)
                                {
                                    profileSystem = (ISystem)brcktTempSuportDefinition.Support;
                                }
                            }

                            if (profileSystem != null)
                            {
                                ReadOnlyCollection<ISystemChild> systemChilds = profileSystem.SystemChildren;

                                foreach (ISystemChild child in systemChilds)
                                {
                                    if (child is Profile)
                                    {
                                        profileSystem = (ISystem)child;
                                        break;
                                    }
                                }
                                systemChilds = profileSystem.SystemChildren;
                                foreach (ISystemChild child in systemChilds)
                                {
                                    if (child is ProfilePart)
                                    {
                                        BusinessObject tempProfilePart = (ProfilePart)child;
                                        profilepartColl1.Add(tempProfilePart);
                                    }
                                }
                            }
                        }
                    }
                }

                ReadOnlyCollection<IConnection> connections = GetConnectionsData((IConnectable)Bracket, typeof(PhysicalConnection));

                foreach (IConnection connection in connections)
                {
                    PhysicalConnection tempPhyconn = connection as PhysicalConnection;

                    if (tempPhyconn != null)
                    {
                        if (tempPhyconn.BoundedObject is ProfilePart)
                        {
                            profilepartColl2.Add(tempPhyconn.BoundedObject);
                        }
                        else if (tempPhyconn.BoundingObject is ProfilePart)
                        {
                            profilepartColl2.Add(tempPhyconn.BoundingObject);
                        }
                    }
                }

                foreach (BusinessObject BusinesObj in profilepartColl1)
                {
                    ProfilePart profilePart1 = BusinesObj as ProfilePart;

                    foreach (BusinessObject BusinesObj1 in profilepartColl2)
                    {
                        ProfilePart profilePart2 = BusinesObj1 as ProfilePart;

                        if ((profilePart1 != null && profilePart2 != null) && (profilePart1 == profilePart2))
                        {
                            connectedProfilePart = profilePart1;
                            break;
                        }
                    }

                    if (connectedProfilePart != null)
                    {
                        break;
                    }
                }

            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.GetConnectedProfilePart: Error encountered (" + e.Message + ")");
            }


            return connectedProfilePart;
        }

        private bool IsStandardCollar(BusinessObject collarPart)
        {
            bool collar = false;
            collar = IsWebBaseAngle90(collarPart);
            return collar;
        }

    }
}
