using System;
using System.IO;
using System.Xml;
using Ingr.SP3D.Manufacturing.Exceptions;

namespace Ingr.SP3D.Content.Manufacturing.Services
{   
    /// <summary>
    /// 
    /// </summary>
    internal class CustomReportLogService
    {
        internal StreamWriter LogFile = null;
        internal FileStream fileStream = null;
        internal XmlDocument newDocument = null;
        //internal XmlNode node = null;
        //internal XmlNode childNode = null;

        /// <summary>s
        /// Creates a text file at the input file path. 
        /// </summary>
        /// <param name="filepath">Complete path of the output file including its name.</param>
        /// <exception cref="ArgumentNullException">Raised when the list of manufacturing entities are null.</exception>
        /// <returns>Returns the LogFile.</returns>
        /// <exception cref="CannotSaveFileException">Raised when the output file cannot be saved at the input file path.</exception>
        internal StreamWriter OpenFile(string filepath)
        {
            try
            {
                fileStream = new FileStream(filepath, FileMode.Append);
                LogFile = new StreamWriter(fileStream);
                LogFile.WriteLine("Starting command: " + DateTime.Now);
                LogFile.WriteLine();
            }
            catch(Exception e)
            {
                throw new CannotSaveFileException("Cannot save file in the given filepath"+ e);
            }
            return LogFile;
        }
        
        /// <summary>
        /// Closes the text file opened for writing.
        /// </summary>
        internal void CloseFile()
        {
            LogFile.WriteLine();
            LogFile.WriteLine();
            LogFile.WriteLine("Ending command: " + DateTime.Now);
            LogFile.Close();
            LogFile = null;
        }

        /// <summary>
        /// Writes the input string into the text file.
        /// </summary>
        internal void WriteToFile(string information)
        {
            LogFile.WriteLine(information);
        }

        /// <summary>
        /// Creates an XML document at the input location.
        /// </summary>
        /// <param name="filePath"></param>
        internal void OpenXMLFile(string filePath)
        {
            try
            {
                newDocument.Load(filePath);

            }
            catch(Exception e)
            {
                throw new CannotSaveFileException("Cannot save file in the given filepath" + e);
            }
        }

        /// <summary>
        /// Creates a new node of given node name.
        /// </summary>
        /// <param name="nodeName">Input for creating th node.</param>
        /// <returns></returns>
        XmlNode CreateNode(string nodeName)
        {
            XmlNode node = null;
            node = newDocument.CreateNode(XmlNodeType.Element, nodeName, null);                                      
            return node;
        }

        /// <summary>
        /// Appends the input node element to the XML document.
        /// </summary>
        /// <param name="node"></param>
        void AppendNode(XmlNode node)
        {
            try
            {
                if (node == null)
                {
                    throw new ArgumentNullException();
                }
                newDocument.AppendChild(node);
            }
            catch(Exception e)
            {
                throw new NodeNotFoundException("Node noe found. " + e);
            }            
        }

        /// <summary>
        /// Creates an XML node with the input string and adds it as a child to the input parent node.
        /// </summary>
        /// <param name="parentNode">node under which child node has to be added</param>
        /// <param name="childNodeName">name of the child Node.</param>
        /// <returns></returns>
        XmlNode AddChildNode(XmlNode parentNode, string childNodeName)
        {
            XmlNode childNode = null;
            try
            {
                childNode = newDocument.CreateNode(XmlNodeType.Element, childNodeName, null);
                parentNode.AppendChild(childNode);
            }
            catch(Exception e)
            {
                throw new NodeNotFoundException("Node noe found. " + e);
            }
            return childNode;
        }

        /// <summary>
        /// Sets the attribute name and value for the input node.
        /// </summary>
        /// <param name="node">The XML node to which attribute information has to be provided.</param>
        /// <param name="attributeName">Name of the attribute</param>
        /// <param name="attributeValue">Value of the attribute</param>
        void AddAttribute(XmlNode node, string attributeName, string attributeValue)
        {
            XmlAttribute attribute = newDocument.CreateAttribute(attributeName,null);
            attribute.Value = attributeValue;
            node.Attributes.Append(attribute);
        }
    }
}
