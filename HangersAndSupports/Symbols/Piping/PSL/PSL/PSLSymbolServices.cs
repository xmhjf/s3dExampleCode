//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSLServiceClass.cs
//   PSL,Ingr.SP3D.Content.Support.Symbols.PSLServiceClass
//   Author       :  BS
//   Creation Date:  21.Aug.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-Aug-2013     BS      CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
//   30-Dec-2014      PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public static class PSLSymbolServices
    {
        private static ReadOnlyCollection<PropertyValue> loadPSLConst;
        private static bool loadFirstTime = true;
        private static IDictionary<KeyValuePair<string, string>, Object> previousPropertyNameValuePair;
        private static bool isSameConditionalSet = false;
        private static bool isSamePartClass = false;
        private static string previousPartClass;
        private static BusinessObject[] parts;
        public enum ComparisionOperator
        {
            /// <summary>
            /// Greater or equal operator.
            /// </summary>
            GREATER_OR_EQUAL,
            /// <summary>
            /// Less than or equal operator.
            /// </summary>
            LESS_OR_EQUAL,
            /// <summary>
            /// Not equal operator.
            /// </summary>
            NOT_EQUAL,
            /// <summary>
            /// Greater than operator.
            /// </summary>
            GREATER,
            /// <summary>
            /// Less than operator.
            /// </summary>
            LESS,
            /// <summary>
            /// Equal operator.
            /// </summary>
            EQUAL,
            /// <summary>
            /// Between the range including limits.
            /// </summary>
            BETWEEN_WITH_LIMITS,
            /// <summary>
            /// Between the range with limits.
            /// </summary>
            BETWEEN_WITHOUT_LIMITS
        }

        /// <summary>
        /// Gets the data on part by condition.
        /// </summary>
        /// <param name="partClassName">The data required part class name.</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="conditionInterface">The condition interface.</param>
        /// <param name="conditionProperty">The condition property.</param>
        /// <param name="referenceValue">The reference value.</param>
        /// <returns>double value</returns>
        /// <example>
        /// <code>
        ///     PropertyValueCodelist loadClassCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass");
        ///     loadClass = loadClassCodeList.PropertyInfo.DisplayName;
        ///     double dis = PSLSymbolServices.GetDataByCondition("PSL_CONSTANT","IJUAFINL_C", "C","IJUAFINL_LoadClass", "LoadClass",loadClass)   
        /// </code>
        /// </example>        
        public static double GetDataByCondition(string partClassName, string dataInterface, string dataProperty, string conditionInterface, string conditionProperty, string referenceValue)
        {
            IEnumerable<BusinessObject> pslParts = null;
            try
            {
                double distance;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass pslPartClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);

                if (pslPartClass.PartClassType.Equals("HgrServiceClass"))
                    pslParts = pslPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    pslParts = pslPartClass.Parts;

                pslParts = pslParts.Where(part => (string)((PropertyValueString)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue == referenceValue);
                if (pslParts.Count() > 0)
                    distance = ((double)((PropertyValueDouble)pslParts.ElementAt(0).GetPropertyValue(dataInterface, dataProperty)).PropValue);
                else
                    distance = 0;
                return distance;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get DataByCondition of PSLServices Class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (pslParts is IDisposable)
                {
                    ((IDisposable)pslParts).Dispose(); // This line will be executed
                }
            }
        }

        /// <summary>
        /// Gets the data on part by condition.
        /// </summary>
        /// <param name="partClassName">The data required part class name.</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="conditionInterface">The condition interface.</param>
        /// <param name="conditionProperty">The condition property.</param>
        /// <param name="referenceValue">The reference value.</param>
        /// <param name="comparisionavalue">The comparision value</param>
        /// <returns>double value</returns>
        /// <example>
        /// <code>
        ///     PropertyValueCodelist loadClassCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass");
        ///     loadClass = loadClassCodeList.PropertyInfo.DisplayName;
        ///     double dis = PSLSymbolServices.GetDataByCondition("PSL_CONSTANT","IJUAFINL_C", "C","IJUAFINL_LoadClass", "LoadClass",loadClass,0.001)   
        /// </code>
        /// </example>        
        public static double GetDataByCondition(string partClassName, string dataInterface, string dataProperty, string conditionInterface, string conditionProperty, double minimumReferencevalue, double maximumReferencevalue)
        {
            IEnumerable<BusinessObject> pslParts = null;
            try
            {
                double distance = 0;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass pslPartClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);

                if (pslPartClass.PartClassType.Equals("HgrServiceClass"))
                    pslParts = pslPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    pslParts = pslPartClass.Parts;

                pslParts = pslParts.Where(part => (double)((PropertyValueDouble)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue > minimumReferencevalue && (double)((PropertyValueDouble)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue < maximumReferencevalue);
                if (pslParts.Count() > 0)
                    distance = ((double)((PropertyValueDouble)pslParts.ElementAt(0).GetPropertyValue(dataInterface, dataProperty)).PropValue);
                return distance;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get DataByCondition of PSLServices Class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (pslParts is IDisposable)
                {
                    ((IDisposable)pslParts).Dispose(); // This line will be executed
                }
            }
        }

        /// <summary>
        /// Gets the data on part by conditions.
        /// </summary>
        /// <param name="partClassName">The data required part class name .</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="propertyNameValuePair">The property name value pair, which is an IDictionary object containing the attribute name and Interface as keyValuepair and value of the other attributes in the table.
        /// These will be used to match with the values of the attributes in the table. For the matching row, the value of the attribute (specified by the 3rd parameter) will be returned.
        /// This method currently assumes that only one row in the table will satisfy the requirement.</param>
        /// <returns> double value</returns>
        /// <example>
        /// <code>
        ///        Dictionary<KeyValuePair<string, string>, Object> parameter = new Dictionary<KeyValuePair<string, string>, Object>();
        ///        parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_CONSTANTS", "SIZE"), size.Trim());
        ///        parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_CONSTANTS", "TOTAL_TRAV"), totalTravel);
        ///
        ///        double E = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "E", parameter);
        ///        parameter = new Dictionary<KeyValuePair<string, string>, Object>();
        ///        'Below code is to get F value from PSL_CONSTANTS service table if 'TOTAL_TRAV' property is lessthan or equal to totalTravel+ tolerance and greaterthan or equal to totalTravel- tolerance
        ///        parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_CONSTANTS", "TOTAL_TRAV"), totalTravel);
        ///        double F = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "F", parameter,0.005,ComparisionOperator.BETWEEN_WITH_LIMITS);
        /// </code>
        /// </example>        
        public static object GetDataByMultipleConditions(string partClassName, string dataInterface, string dataProperty, IDictionary<KeyValuePair<string, string>, Object> propertyNameValuePair, double tolerance = 0, ComparisionOperator operatorType = ComparisionOperator.EQUAL)
        {
            IEnumerable<BusinessObject> pslCmpParts = null;
            object value = null;
            try
            {
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass pslCmpPartClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);
                Object[] referenceValues = new Object[propertyNameValuePair.Values.Count];
                KeyValuePair<string, string>[] propertyInterfaces = new KeyValuePair<string, string>[propertyNameValuePair.Keys.Count];
                propertyNameValuePair.Keys.CopyTo(propertyInterfaces, 0);
                propertyNameValuePair.Values.CopyTo(referenceValues, 0);
                if (!String.IsNullOrEmpty(previousPartClass))
                {
                    if (previousPartClass.Equals(partClassName))
                    {
                        isSamePartClass = true;
                        if (isDictionariesEqual(previousPropertyNameValuePair, propertyNameValuePair))
                            isSameConditionalSet = true;
                        else
                            isSameConditionalSet = false;
                    }
                    else
                    {
                        isSamePartClass = false;
                        isSameConditionalSet = false;
                    }
                }
                if (loadFirstTime)
                {
                    if (!isSamePartClass)
                    {
                        if (pslCmpPartClass.PartClassType.Equals("HgrServiceClass"))
                            pslCmpParts = pslCmpPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                        else
                            pslCmpParts = pslCmpPartClass.Parts;
                        parts = pslCmpParts.ToArray<BusinessObject>();
                        //parts = pslCmpParts;
                    }
                    for (int i = 0; i < parts.Length; i++)
                    {
                        if (CheckCondition(parts[i], propertyNameValuePair, tolerance, operatorType))
                        {
                            loadPSLConst = parts[i].GetAllProperties();
                            previousPartClass = partClassName;
                            previousPropertyNameValuePair = propertyNameValuePair;
                            loadFirstTime = false;
                            break;
                        }
                    }
                }
                else
                {
                    if (!isSameConditionalSet)
                    {
                        if (!isSamePartClass)
                        {
                            if (pslCmpPartClass.PartClassType.Equals("HgrServiceClass"))
                                pslCmpParts = pslCmpPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                            else
                                pslCmpParts = pslCmpPartClass.Parts;
                            parts = pslCmpParts.ToArray<BusinessObject>();
                        }
                        for (int i = 0; i < parts.Length; i++)
                        {
                            if (CheckCondition(parts[i], propertyNameValuePair, tolerance, operatorType))
                            {
                                loadPSLConst = parts[i].GetAllProperties();
                                previousPartClass = partClassName;
                                previousPropertyNameValuePair = propertyNameValuePair;
                                loadFirstTime = false;
                                break;
                            }

                        }
                    }
                }
                SP3DPropType propType;

                if (loadPSLConst != null)
                {
                    foreach (PropertyValue propvalue in loadPSLConst)
                    {
                        if (propvalue.PropertyInfo.InterfaceInfo.Name.Equals(dataInterface))
                        {
                            if (propvalue.PropertyInfo.Name.Equals(dataProperty))
                            {
                                propType = propvalue.PropertyInfo.PropertyType;

                                switch (propType)
                                {
                                    case SP3DPropType.PTUndefined:
                                        break;
                                    case SP3DPropType.PTInteger:
                                        PropertyValueInt propValInt = (PropertyValueInt)propvalue;
                                        value = propValInt.PropValue;
                                        break;
                                    case SP3DPropType.PTString:
                                        PropertyValueString propValString = (PropertyValueString)propvalue;
                                        value = propValString.PropValue;
                                        break;
                                    case SP3DPropType.PTBool:
                                        PropertyValueBoolean propValBoolean = (PropertyValueBoolean)propvalue;
                                        value = propValBoolean.PropValue;
                                        break;
                                    case SP3DPropType.PTDate:
                                        PropertyValueDateTime propValDateTime = (PropertyValueDateTime)propvalue;
                                        value = propValDateTime.PropValue;
                                        break;
                                    case SP3DPropType.PTDouble:
                                        PropertyValueDouble propValDouble = (PropertyValueDouble)propvalue;
                                        value = propValDouble.PropValue;
                                        break;
                                    case SP3DPropType.PTCodelist:
                                        PropertyValueCodelist propValCodelist = (PropertyValueCodelist)propvalue;
                                        value = propvalue.PropertyInfo.CodeListInfo.GetCodelistItem(propValCodelist.PropValue).Name;
                                        break;
                                    case SP3DPropType.PTShort:
                                        PropertyValueShort propValShort = (PropertyValueShort)propvalue;
                                        value = propValShort.PropValue;
                                        break;
                                    case SP3DPropType.PTFloat:
                                        PropertyValueFloat propValFloat = (PropertyValueFloat)propvalue;
                                        value = propValFloat.PropValue;
                                        break;
                                    default:
                                        break;
                                }
                                break;
                            }

                        }
                    }
                }

                return value;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get DataByCondition of PSLSymbolServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (pslCmpParts is IDisposable)
                {
                    ((IDisposable)pslCmpParts).Dispose(); // This line will be executed
                }
            }
        }

        /// <summary>
        /// Determines whether both dictionaries are equal for the specified property name value pairs.
        /// </summary>
        /// <param name="propertyNameValuePair1">First property name value pair, which is an IDictionary object containing the attribute name and Interface as keyValuepair and value of the other attributes in the table.
        /// These will be used to match with the values of the attributes in the table. For the matching row, the value of the attribute (specified by the 3rd parameter) will be returned.
        /// This method currently assumes that only one row in the table will satisfy the requirement..</param>
        /// <param name="propertyNameValuePair2">Second property name value pair.</param>
        /// <returns></returns>
        private static bool isDictionariesEqual(IDictionary<KeyValuePair<string, string>, Object> propertyNameValuePair1, IDictionary<KeyValuePair<string, string>, Object> propertyNameValuePair2)
        {
            bool isEqual = false;
            if (propertyNameValuePair1.Count == propertyNameValuePair2.Count)
            {
                string[] interfaces1 = new string[propertyNameValuePair1.Keys.Count];
                string[] properties1 = new string[propertyNameValuePair1.Keys.Count];
                Object[] referenceValues1 = new Object[propertyNameValuePair1.Values.Count];
                KeyValuePair<string, string>[] propertyInterfaces1 = new KeyValuePair<string, string>[propertyNameValuePair1.Keys.Count];
                propertyNameValuePair1.Keys.CopyTo(propertyInterfaces1, 0);
                for (int i = 0; i < propertyInterfaces1.Length; i++)
                {
                    interfaces1[i] = propertyInterfaces1[i].Key;
                    properties1[i] = propertyInterfaces1[i].Value;
                }
                propertyNameValuePair1.Values.CopyTo(referenceValues1, 0);
                string[] interfaces2 = new string[propertyNameValuePair1.Keys.Count];
                string[] properties2 = new string[propertyNameValuePair1.Keys.Count];
                Object[] referenceValues2 = new Object[propertyNameValuePair1.Values.Count];
                KeyValuePair<string, string>[] propertyInterfaces2 = new KeyValuePair<string, string>[propertyNameValuePair2.Keys.Count];
                propertyNameValuePair2.Keys.CopyTo(propertyInterfaces2, 0);
                propertyNameValuePair2.Values.CopyTo(referenceValues2, 0);
                for (int i = 0; i < propertyInterfaces2.Length; i++)
                {
                    interfaces2[i] = propertyInterfaces2[i].Key;
                    properties2[i] = propertyInterfaces2[i].Value;
                }
                bool interfacesEqual = ArraysEqual(interfaces1, interfaces2);
                bool propertiesEqual = ArraysEqual(properties1, properties2);
                bool valuesEqual = ArraysEqual(referenceValues1, referenceValues2);
                if (interfacesEqual &&  valuesEqual)
                    isEqual = true;
                else
                    isEqual = false;
            }
            return isEqual;
        }
        /// <summary>
        /// Checks if arrays are equal.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="a1">The Array 1.</param>
        /// <param name="a2">The Array 2.</param>
        /// <returns></returns>
        private static bool ArraysEqual<T>(T[] a1, T[] a2)
        {
            if (ReferenceEquals(a1, a2))
                return true;

            if (a1 == null || a2 == null)
                return false;

            if (a1.Length != a2.Length)
                return false;

            EqualityComparer<T> comparer = EqualityComparer<T>.Default;
            for (int i = 0; i < a1.Length; i++)
            {
                if (!comparer.Equals(a1[i], a2[i])) return false;
            }
            return true;
        }
        /// <summary>
        /// Checks the if condition for the given Comparision Operator type .
        /// </summary>
        /// <param name="part">The part on which proerties are to be check.</param>
        /// <param name="propertyNameValuePair">The property name value pair.</param>
        /// <param name="tolerance">The tolerance.</param>
        /// <param name="operatorType">ComparisionOperator type: Example EQUAL,NOT_EQUAL</param>
        /// <returns></returns>
        private static bool CheckCondition(BusinessObject part, IDictionary<KeyValuePair<string, string>, Object> propertyNameValuePair, double tolerance, ComparisionOperator operatorType = ComparisionOperator.EQUAL)
        {
            //string[] interfaces1 = new string[propertyNameValuePair1.Keys.Count];
            //string[] properties1 = new string[propertyNameValuePair1.Keys.Count];
            Object[] referenceValues = new Object[propertyNameValuePair.Values.Count];
            KeyValuePair<string, string>[] propertyInterfaces = new KeyValuePair<string, string>[propertyNameValuePair.Keys.Count];
            propertyNameValuePair.Keys.CopyTo(propertyInterfaces, 0);
            propertyNameValuePair.Values.CopyTo(referenceValues, 0);
            bool result = false;
            SP3DPropType propType;
            for (int i = 0; i < propertyNameValuePair.Keys.Count; i++)
            {
                PropertyValue propvalue = part.GetPropertyValue(propertyInterfaces[i].Key, propertyInterfaces[i].Value);
                if (propvalue != null)
                {
                    propType = propvalue.PropertyInfo.PropertyType;

                    switch (propType)
                    {
                        case SP3DPropType.PTUndefined:
                            break;
                        case SP3DPropType.PTInteger:
                            PropertyValueInt propValInt = (PropertyValueInt)propvalue;
                            result = GetResultForCondition<int>((int)propValInt.PropValue, (int)referenceValues[i], tolerance, operatorType);
                            break;
                        case SP3DPropType.PTString:
                            PropertyValueString propValString = (PropertyValueString)propvalue;
                            result = GetResultForCondition<string>((string)propValString.PropValue, (string)referenceValues[i], tolerance, operatorType);
                            break;
                        case SP3DPropType.PTBool:
                            PropertyValueBoolean propValBoolean = (PropertyValueBoolean)propvalue;
                            result = GetResultForCondition<bool>((bool)propValBoolean.PropValue, (bool)referenceValues[i], tolerance, operatorType);
                            break;
                        case SP3DPropType.PTDate:
                            PropertyValueDateTime propValDateTime = (PropertyValueDateTime)propvalue;
                            result = GetResultForCondition<DateTime>((DateTime)propValDateTime.PropValue, (DateTime)referenceValues[i], tolerance, operatorType);
                            break;
                        case SP3DPropType.PTDouble:
                            PropertyValueDouble propValDouble = (PropertyValueDouble)propvalue;
                            result = GetResultForCondition<double>((double)propValDouble.PropValue, (double)referenceValues[i], tolerance, operatorType);
                            break;
                        case SP3DPropType.PTCodelist:
                            PropertyValueCodelist propValCodelist = (PropertyValueCodelist)propvalue;
                            result = GetResultForCondition<int>((int)propValCodelist.PropValue, (int)referenceValues[i], tolerance, operatorType);
                            break;
                        case SP3DPropType.PTShort:
                            PropertyValueShort propValShort = (PropertyValueShort)propvalue;
                            result = GetResultForCondition<short>((short)propValShort.PropValue, (short)referenceValues[i], tolerance, operatorType);
                            break;
                        case SP3DPropType.PTFloat:
                            PropertyValueFloat propValFloat = (PropertyValueFloat)propvalue;
                            result = GetResultForCondition<float>((float)propValFloat.PropValue, (float)referenceValues[i], tolerance, operatorType);
                            break;
                        default:
                            break;
                    }
                }
                if (result == false) return false;
            }
            return true;
        }
        /// <summary>
        /// Gets the result for condition.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="propValue">The property value.</param>
        /// <param name="refernceValue">The refernce value.</param>
        /// <param name="tolerance">The tolerance.</param>
        /// <param name="operatorType">ComparisionOperator type.</param>
        /// <returns></returns>
        private static bool GetResultForCondition<T>(T propValue, T refernceValue, double tolerance, ComparisionOperator operatorType) where T : IComparable
        {
            bool result = false;
            switch (operatorType)
            {
                case ComparisionOperator.EQUAL:
                    result = propValue.CompareTo(refernceValue) == 0;
                    break;
                case ComparisionOperator.GREATER:
                    result = propValue.CompareTo(refernceValue) > 0;
                    break;
                case ComparisionOperator.LESS:
                    result = propValue.CompareTo(refernceValue) < 0;
                    break;
                case ComparisionOperator.BETWEEN_WITH_LIMITS:
                    if (typeof(T) == typeof(double))
                        result = (double)(object)propValue >= ((double)(object)refernceValue - tolerance) && (double)(object)propValue <= ((double)(object)refernceValue + tolerance);
                    else if (typeof(T) == typeof(int))
                        result = (int)(object)propValue >= ((int)(object)refernceValue - tolerance) && (int)(object)propValue <= ((int)(object)refernceValue + tolerance);
                    break;
                case ComparisionOperator.BETWEEN_WITHOUT_LIMITS:
                    if (typeof(T) == typeof(double))
                        result = (double)(object)propValue > ((double)(object)refernceValue - tolerance) && (double)(object)propValue < ((double)(object)refernceValue + tolerance);
                    else if (typeof(T) == typeof(int))
                        result = (int)(object)propValue > ((int)(object)refernceValue - tolerance) && (int)(object)propValue < ((int)(object)refernceValue + tolerance);
                    break;
                case ComparisionOperator.NOT_EQUAL:
                    result = propValue.CompareTo(refernceValue) != 0;
                    break;
                case ComparisionOperator.GREATER_OR_EQUAL:
                    result = propValue.CompareTo(refernceValue) >= 0;
                    break;
                case ComparisionOperator.LESS_OR_EQUAL:
                    result = propValue.CompareTo(refernceValue) <= 0;
                    break;
            }
            return result;

        }
    }
}
