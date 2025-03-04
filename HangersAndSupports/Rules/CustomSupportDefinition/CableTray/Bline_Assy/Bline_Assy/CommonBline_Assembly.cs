using System;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

namespace Ingr.SP3D.Content.Support.Rules
{
    class CommonBline_Assembly
    {
        /// <summary>
        /// This method returns the PartNumber from Auxiliary table.
        /// </summary>
        /// <param name="PartClass">Parclass Name.</param>
        /// <param name="Interface">Interface Name.</param>
        /// <param name="Atttribute">Attribute Name</param>
        /// <param name="Parameters">Input Parametes</param>
        /// <returns>string</returns>        
        /// <code>
        ///      CommonBline_Assembly cmnBline_Assembly = new CommonBline_Assembly();
        ///        string ChannelPartNumber; 
        ///       ChannelPartNumber = cmnBline_Assembly.GetPartFromTable("BLineAssyChannelsAUX", "IJUAHgrAssyChannelsAUX", "ChannelPartNo", parameters);
        ///</code>
        public string GetPartFromTable(string PartClass, string Interface, string Atttribute, Dictionary<string, string> Parameters)
        {
            try
            {
                CatalogBaseHelper Cataloghelper = new CatalogBaseHelper();

                PartClass AuxTable = (PartClass)Cataloghelper.GetPartClass(PartClass);

                ReadOnlyCollection<BusinessObject> classItems = AuxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                foreach (BusinessObject classItem in classItems)
                {
                    ReadOnlyCollection<PropertyValue> properties = classItem.GetAllProperties();
                    foreach (PropertyValue propVal in properties)
                    {
                        int numberOfParamsMatching = 0;
                        for (int i = 0; i < properties.Count; i++)
                        {
                            if (Parameters.ContainsKey(properties[i].PropertyInfo.Name))
                            {
                                string value;
                                bool temp = Parameters.TryGetValue(properties[i].PropertyInfo.Name, out value);
                                if (value == properties[i].ToString())
                                {
                                    numberOfParamsMatching++;
                                    if (numberOfParamsMatching == Parameters.Count)
                                    {
                                        return classItem.GetPropertyValue(Interface, Atttribute).ToString();
                                    }
                                }
                                else
                                    break;
                            }
                        }
                        break;
                    }
                }
                return null; // Add exception handling
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in CommonBline_Assembly of Bline_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

    }
}
