/* 
 * 
 * Copyright (c) 2016, Softing Industrial Automation GmbH. All rights reserved.
 * 
 * XMItoNodeset is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace XMItoNodeset
{
    class XMItoNodeset
    {
        List<XmlDocument> _xmiDocList;
        XmlDocument _nodesetDoc;
        XmlNode _nodesetUANodeSetNode;
        XmlNode _nodesetAliasesNode;
        XmlNode _nodesetNamespaceUrisNode;
        XmlDocument _binaryTypesDoc;
        XmlNode _binaryTypesRootNode;
        XmlDocument _xmlTypesDoc;
        XmlNode _xmlTypesRootNode;
        Dictionary<String, XmlNode> _xmiClassMap;
        List<XmlNode> _xmiClassList;
        Dictionary<String, XmlNode> _xmiStructureMap;
        List<XmlNode> _xmiStructureList;
        Dictionary<String, XmlNode> _xmiEnumerationMap;
        List<XmlNode> _xmiEnumerationList;
        Dictionary<String, XmlNode> _xmiElementMapEA;
        Dictionary<String, XmlNode> _xmiAttributeMapEA;
        Dictionary<String, XmlNode> _xmiOwnedAttributeMapIgnore;
        Dictionary<String, XmlNode> _xmiOperationsMapEA;
        Dictionary<String, XmlNode> _xmiStubEAMap;
        Dictionary<String, XmlNode> _xmiXElementMap;
        Dictionary<String, String> _nodesetNodeIdMap;
        Dictionary<String, XmlNode> _xmiPackageMap;
        Int32 _nextNodeId;
        List<String> _xmiFiles;
        bool[] _xmiFilesClassIsOT;
        bool[] _xmiFilesClassRefUseStereotype;
        String _nodesetFile;
        String _nodesetImportFile;
        String _nodesetURL;
        bool _nodesetHasDT;
        String _nodesetTypeDictionaryName;
        List<String> _ignoreClassMemberList;
        string[,] _aliases;
        string _nodeIdMapFileName;
        string _binaryTypesFileName;
        string _xmlTypesFileName;
        string _xmiGenerate;
        string _wordFileName;

        Application _wordApp = null;
        Document _wordDoc = null;
        Table _wordCurrentTable = null;

        const string _nodeIdTextBinarySchema = "BinarySchema";
        const string _nodeIdTextXmlSchema = "XmlSchema";

        const string _xmiNSURN = "http://schema.omg.org/spec/XMI/2.1";
        string _xmiNSPräfix;
        const string _umlNSURN = "http://schema.omg.org/spec/UML/2.1";
        string _umlNSPräfix;

        static void Main(string[] args)
        {
            XMItoNodeset prog = new XMItoNodeset();
            prog.start(args);
        }

        bool parseCommandLineArgs(string[] args)
        {
            string command = "";
            _xmiFiles = new List<String>();
            _ignoreClassMemberList =  new List<String>();
            _xmiFilesClassIsOT = new bool[100];
            _xmiFilesClassRefUseStereotype = new bool[100];
            _nodesetTypeDictionaryName = "XMItoNodeset";
            _nodesetURL = "http://industrial.softing.com/XMItoNodeset";
            _nodesetFile = "nodeset.xml";
            _nodesetImportFile = "";
             _nextNodeId = 0;
            _nodeIdMapFileName = "NodeIdMap.txt";
            _binaryTypesFileName = "BinaryTypes.xml";
            _xmlTypesFileName = "XmlTypes.xml";
            _xmiGenerate = "";
            _wordFileName = "";

            int i = 0;

            foreach (string arg in args)
            {
                if (command == "")
                { // no command set -> has to be specified
                    if ((arg == "/xmi") || (arg == "/xmiS") || (arg == "/xmiDT") || (arg == "/nodeset") || (arg == "/nodesetUrl") || (arg == "/nodesetTypeDictionary") || (arg == "/nodesetImport") || (arg == "/nodesetStartId") || (arg == "/ignoreClassMember") || (arg == "/nodeIdMap") || (arg == "/binaryTypes") || (arg == "/xmlTypes") || (arg == "/generate") || (arg == "/word"))
                    {
                        command = arg;
                    }
                    else
                    {
                        Console.WriteLine("Invalid command: {0}", arg);
                        return false;
                    }
                }
                else
                { // command argument
                    if (command == "/xmi")
                    {
                        _xmiFiles.Add(arg);
                        _xmiFilesClassIsOT[i] = true;
                        _xmiFilesClassRefUseStereotype[i] = false;
                        i++;
                    }
                    else if (command == "/xmiS")
                    {
                        _xmiFiles.Add(arg);
                        _xmiFilesClassIsOT[i] = true;
                        _xmiFilesClassRefUseStereotype[i] = true;
                        i++;
                    }
                    else if (command == "/xmiDT")
                    {
                        _xmiFiles.Add(arg);
                        _xmiFilesClassIsOT[i] = false;
                        _xmiFilesClassRefUseStereotype[i] = false;
                        i++;
                    }
                    else if (command == "/nodeset")
                    {
                        _nodesetFile = arg;
                    }
                    else if (command == "/nodesetUrl")
                    {
                        _nodesetURL = arg;
                    }
                    else if (command == "/nodesetStartId")
                    {
                        _nextNodeId = Int32.Parse(arg);
                    }
                    else if (command == "/nodesetImport")
                    {
                        _nodesetImportFile = arg;
                    }
                    else if (command == "/nodesetTypeDictionary")
                    {
                        _nodesetTypeDictionaryName = arg;
                    }
                    else if (command == "/ignoreClassMember")
                    {
                        _ignoreClassMemberList.Add(arg);
                    }
                    else if (command == "/nodeIdMap")
                    {
                        _nodeIdMapFileName = arg;
                    }
                    else if (command == "/binaryTypes")
                    {
                        _binaryTypesFileName = arg;
                    }
                    else if (command == "/xmlTypes")
                    {
                        _xmlTypesFileName = arg;
                    }
                    else if (command == "/generate")
                    {
                        _xmiGenerate = arg;
                    }
                    else if (command == "/word")
                    {
                        _wordFileName = arg;
                    }
                    command = "";
                }
            }
            return true;
        }

        void start(string[] args)
        {
            if (!parseCommandLineArgs(args))
            {
                Console.WriteLine("XMItoNodeset /xmi <xmi file> /nodeset <nodeset file> /nodeseturl <URL for nodesetfile>");
                return;
            }

            try
            {
                if (_wordFileName != "")
                { 
                    _wordApp = new Application();
                    _wordDoc = _wordApp.Documents.Add();
                    _wordDoc.Paragraphs.SpaceAfter = 0;
                }
            }
            catch
            { }

            _xmiClassMap = new Dictionary<String, XmlNode>();
            _xmiClassList = new List<XmlNode>();
            _xmiStructureMap = new Dictionary<String, XmlNode>();
            _xmiStructureList = new List<XmlNode>();
            _xmiEnumerationMap = new Dictionary<String, XmlNode>();
            _xmiEnumerationList = new List<XmlNode>();
            _xmiElementMapEA = new Dictionary<String, XmlNode>();
            _xmiAttributeMapEA = new Dictionary<String, XmlNode>();
            _xmiOwnedAttributeMapIgnore = new Dictionary<String, XmlNode>();
            _xmiOperationsMapEA = new Dictionary<String, XmlNode>();
            _xmiStubEAMap = new Dictionary<String, XmlNode>();
            _xmiXElementMap = new Dictionary<String, XmlNode>();
            _xmiPackageMap = new Dictionary<String, XmlNode>();
            _nodesetNodeIdMap = new Dictionary<String, String>();
            _nodesetHasDT = false;

            // load _nodesetNodeIdMap
            try
            { 
                System.IO.StreamReader nodeIdMapFileR = new System.IO.StreamReader(_nodeIdMapFileName);
                string nodeIdMapFileRLine;
                if ((nodeIdMapFileRLine = nodeIdMapFileR.ReadLine()) != null)
                {
                     _nextNodeId = Int32.Parse(nodeIdMapFileRLine);
                }
                while ((nodeIdMapFileRLine = nodeIdMapFileR.ReadLine()) != null)
                {
                    string[] split = nodeIdMapFileRLine.Split('\t');
                    _nodesetNodeIdMap[split[0]] = split[1];
                }
                nodeIdMapFileR.Close();
            }
            catch
            { }

            // load XMI documents
            _xmiDocList = new List<XmlDocument>();
            foreach (string xmiFile in _xmiFiles)
            {
                XmlDocument xmiDoc = new XmlDocument();

                Console.WriteLine("Load XMI file: {0}", xmiFile);
                try
                {
                    xmiDoc.Load(xmiFile);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error loading file - {0}", e.Message);
                    return;
                }

                _xmiDocList.Add(xmiDoc);

                if (_xmiNSPräfix == null)
                {
                    _xmiNSPräfix = xmiDoc.ChildNodes[1].GetPrefixOfNamespace(_xmiNSURN);
                }
                if (_umlNSPräfix == null)
                {
                    _umlNSPräfix = xmiDoc.ChildNodes[1].GetPrefixOfNamespace(_umlNSURN);
                }
            }

            initOutputXmlDocuments();
            importNodeset();

            int i = 0;
            foreach (XmlDocument xmiDoc in _xmiDocList)
            {
                Dictionary<String, String> xmiElementsOptions = new Dictionary<string, string>();
                xmiElementsOptions.Add("IsObjectType", _xmiFilesClassIsOT[i].ToString());
                xmiElementsOptions.Add("ClassRefUseStereotype", _xmiFilesClassRefUseStereotype[i].ToString());
                i++;

                getXmiElements(xmiDoc, xmiDoc.ChildNodes, xmiElementsOptions);
            }

            createNodesetEnumerationDataTypes();
            createNodesetStructureDataTypes();
            createNodesetObjectTypes();
            if (_nodesetHasDT)
            {
                createNodesetTypeDictionary();
            }

            Console.WriteLine("Save Nodeset file: {0}", _nodesetFile);
            _nodesetDoc.Save(_nodesetFile);

            // store _nodesetNodeIdMap
            System.IO.StreamWriter nodeIdMapFile = new System.IO.StreamWriter(_nodeIdMapFileName);
            nodeIdMapFile.WriteLine("{0}", _nextNodeId);
            foreach (var pair in _nodesetNodeIdMap)
            {
                nodeIdMapFile.WriteLine("{0}\t{1}", pair.Key, pair.Value);
            }
            nodeIdMapFile.Close();      
            
            // store word document
            if ((_wordApp != null) && (_wordDoc != null))
            {
                _wordApp.ActiveDocument.SaveAs(_wordFileName, WdSaveFormat.wdFormatDocumentDefault);
                _wordDoc.Close();

                _wordApp.Quit();
                 Marshal.FinalReleaseComObject(_wordApp);
            }
        }

        void initOutputXmlDocuments()
        {
            // nodeset document 
            _nodesetDoc = new XmlDocument();
            XmlNode nodesetDocNode = _nodesetDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
            _nodesetDoc.AppendChild(nodesetDocNode);
            _nodesetUANodeSetNode = addXmlElement(_nodesetDoc, _nodesetDoc, "UANodeSet");
            addXmlAttribute(_nodesetDoc, _nodesetUANodeSetNode, "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
            addXmlAttribute(_nodesetDoc, _nodesetUANodeSetNode, "xmlns:xsd", "http://www.w3.org/2001/XMLSchema");
            addXmlAttribute(_nodesetDoc, _nodesetUANodeSetNode, "LastModified", String.Format("{0:yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fffffff'Z'}", DateTime.Now));
            addXmlAttribute(_nodesetDoc, _nodesetUANodeSetNode, "xmlns", "http://opcfoundation.org/UA/2011/03/UANodeSet.xsd");
            _nodesetNamespaceUrisNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "NamespaceUris");
            addXmlElement(_nodesetDoc, _nodesetNamespaceUrisNode, "Uri", _nodesetURL);

            addAliases();

            // binary types
            _binaryTypesDoc = new XmlDocument();
            _binaryTypesRootNode = addQualifiedXmlElement(_binaryTypesDoc, _binaryTypesDoc, "opc", "http://opcfoundation.org/BinarySchema/", "TypeDictionary");
            addXmlAttribute(_binaryTypesDoc, _binaryTypesRootNode, "xmlns:opc", "http://opcfoundation.org/BinarySchema/");
            addXmlAttribute(_binaryTypesDoc, _binaryTypesRootNode, "xmlns:ua", "http://opcfoundation.org/UA/");
            addXmlAttribute(_binaryTypesDoc, _binaryTypesRootNode, "xmlns:tns", _nodesetURL);
            addXmlAttribute(_binaryTypesDoc, _binaryTypesRootNode, "DefaultByteOrder", "LittleEndian");
            addXmlAttribute(_binaryTypesDoc, _binaryTypesRootNode, "TargetNamespace", _nodesetURL);

            addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, _binaryTypesRootNode, "opc", "http://opcfoundation.org/BinarySchema/", "Import", "Namespace", "http://opcfoundation.org/UA/", "Location", "Opc.Ua.BinarySchema.bsd");

            // xml types
            _xmlTypesDoc = new XmlDocument();
            _xmlTypesRootNode = addQualifiedXmlElement(_xmlTypesDoc, _xmlTypesDoc, "xs", "http://www.w3.org/2001/XMLSchema", "schema");
            addXmlAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xmlns:tns", _nodesetURL);
            addXmlAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xmlns:xs", "http://www.w3.org/2001/XMLSchema");
            addXmlAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xmlns:ua", "http://opcfoundation.org/UA/2008/02/Types.xsd");
            addXmlAttribute(_xmlTypesDoc, _xmlTypesRootNode, "targetNamespace", _nodesetURL);
            addXmlAttribute(_xmlTypesDoc, _xmlTypesRootNode, "elementFormDefault", "qualified");
            addQualifiedXmlElementAndOneAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "Import", "namespace", "http://opcfoundation.org/UA/2008/02/Types.xsd");
        }

        void getXmiElements(XmlDocument xmiDoc, XmlNodeList list, Dictionary<String, String> xmiElementsOptions)
        {
            foreach (XmlNode node in list)
            {
                if (node.Name == "packagedElement") // XMI
                {
                    XmlAttribute xmiType = node.Attributes[String.Format("{0}:type", _xmiNSPräfix)];
                    if (xmiType != null)
                    {
                        if (xmiType.Value == String.Format("{0}:Class", _umlNSPräfix))
                        {
                            XmlAttribute xmiId = node.Attributes[String.Format("{0}:id", _xmiNSPräfix)];
                            if (xmiId != null)
                            {
                                if (xmiElementsOptions["IsObjectType"] == "True")
                                {
                                    try
                                    {
                                        if (xmiElementsOptions["ClassRefUseStereotype"] == "True")
                                        {
                                            addXmlAttribute(xmiDoc, node, "XMitoNodeset-ClassRefUseStereotype", "true");
                                        }
                                        _xmiClassMap.Add(xmiId.Value, node);
                                        _xmiClassList.Add(node);
                                    }
                                    catch
                                    { }
                                }
                                else
                                {
                                    try
                                    {
                                        _xmiStructureMap.Add(xmiId.Value, node);
                                        _xmiStructureList.Add(node);
                                    }
                                    catch
                                    { }
                                }
                            }
                        }
                        else if (xmiType.Value == String.Format("{0}:Enumeration", _umlNSPräfix))
                        {
                            XmlAttribute xmiId = node.Attributes[String.Format("{0}:id", _xmiNSPräfix)];
                            if (xmiId != null)
                            {
                                try
                                {
                                    _xmiEnumerationMap.Add(xmiId.Value, node);
                                    _xmiEnumerationList.Add(node);
                                }
                                catch
                                { }
                            }
                        }
                        else if (xmiType.Value == String.Format("{0}:Package", _umlNSPräfix))
                        {
                            XmlAttribute xmiId = node.Attributes[String.Format("{0}:id", _xmiNSPräfix)];
                            if (xmiId != null)
                            {
                                try
                                {
                                    _xmiPackageMap.Add(xmiId.Value, node);
                                }
                                catch
                                { }
                            }
                        }        
                    }
                }

                if (node.Name == "XpackagedElement") // Manual extention
                {
                    XmlAttribute xmiId = node.Attributes[String.Format("{0}:id", _xmiNSPräfix)];
                    if (xmiId != null)
                    {
                        try
                        {
                            _xmiXElementMap.Add(xmiId.Value, node);
                        }
                        catch
                        { }
                    }
                }

                if (node.Name == "IownedAttribute") // Manual extention
                {
                    XmlAttribute xmiId = node.Attributes[String.Format("{0}:id", _xmiNSPräfix)];
                    if (xmiId != null)
                    {
                        try
                        {
                            _xmiOwnedAttributeMapIgnore.Add(xmiId.Value, node);
                        }
                        catch
                        { }
                    }
                }

                if (node.Name == "element") // EA extension
                {
                    XmlAttribute xmiType = node.Attributes[String.Format("{0}:type", _xmiNSPräfix)];
                    if (xmiType != null)
                    {
                        if (xmiType.Value == String.Format("{0}:Class", _umlNSPräfix))
                        {
                            XmlAttribute xmiId = node.Attributes[String.Format("{0}:idref", _xmiNSPräfix)];
                            if (xmiId != null)
                            {
                                try
                                {
                                    _xmiElementMapEA.Add(xmiId.Value, node);
                                }
                                catch
                                { }
                            }
                        }
                        else if (xmiType.Value == String.Format("{0}:Package", _umlNSPräfix))
                        {
                            XmlAttribute xmiId = node.Attributes[String.Format("{0}:idref", _xmiNSPräfix)];
                            if (xmiId != null)
                            {
                                string generateTag = getTag(node, "Generate");
                                if (generateTag != null)
                                {
                                    XmlNode xmiPackage = getXmiPackage(xmiId.Value);
                                    if (xmiPackage != null)
                                    {
                                        addXmlAttributeDeep(xmiDoc, xmiPackage, "XMitoNodeset-Generate", generateTag);
                                    }
                                }
                            }
                        }
                    }
                }

                if (node.Name == "attribute") // EA extension
                {
                    XmlAttribute xmiId = node.Attributes[String.Format("{0}:idref", _xmiNSPräfix)];
                    if (xmiId != null)
                    {
                        try
                        {
                            _xmiAttributeMapEA.Add(xmiId.Value, node);
                        }
                        catch
                        { }
                    }
                }

                if (node.Name == "operation") // EA extension
                {
                    XmlAttribute xmiId = node.Attributes[String.Format("{0}:idref", _xmiNSPräfix)];
                    if (xmiId != null)
                    {
                        try
                        {
                            _xmiOperationsMapEA.Add(xmiId.Value, node);
                        }
                        catch
                        { }
                    }
                }

                if (node.Name == "EAStub") // EA stub
                {
                    XmlAttribute xmiType = node.Attributes["UMLType"];
                    if (xmiType != null)
                    {
                        if (xmiType.Value == "Class")
                        {
                            XmlAttribute xmiId = node.Attributes[String.Format("{0}:id", _xmiNSPräfix)];
                            if (xmiId != null)
                            {
                                try
                                {
                                    _xmiStubEAMap.Add(xmiId.Value, node);
                                }
                                catch
                                { }
                            }
                        }
                    }
                }

                getXmiElements(xmiDoc, node.ChildNodes, xmiElementsOptions);
            }
        }

        void importNodeset()
        {
            if (_nodesetImportFile == "")
            {
                return;
            }
             
            Console.WriteLine("Load Nodeset insert file: {0}", _nodesetImportFile);
            XmlDocument nodesetImport = new XmlDocument();
            try
            {
                nodesetImport.Load(_nodesetImportFile);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error loading file - {0}", e.Message);
                return;
            }
            foreach (XmlNode nodesetImportNode in nodesetImport.DocumentElement.ChildNodes)
            {
                XmlNode nodesetImportedNode = _nodesetDoc.ImportNode(nodesetImportNode, true);
                if (nodesetImportedNode.Name == "NamespaceUris")
                {
                    foreach (XmlNode nodesetImportedSubNode in nodesetImportedNode.ChildNodes)
                    {
                        XmlNode clone = nodesetImportedSubNode.Clone();
                        _nodesetNamespaceUrisNode.AppendChild(clone);
                    }
                }
                else if (nodesetImportedNode.Name == "Aliases")
                {
                    foreach (XmlNode nodesetImportedSubNode in nodesetImportedNode.ChildNodes)
                    {
                        bool nodesetSameAttr = false;
                        foreach (XmlNode nodesetAliasNode in _nodesetAliasesNode.ChildNodes)
                        {
                            if (nodesetImportedSubNode.Attributes["Alias"].Value == nodesetAliasNode.Attributes["Alias"].Value)
                            {
                                nodesetAliasNode.InnerText = nodesetImportedSubNode.InnerText;
                                nodesetSameAttr = true;
                                break;
                            }
                        }
                        if (!nodesetSameAttr)
                        {
                            XmlNode clone = nodesetImportedSubNode.Clone();
                            _nodesetAliasesNode.AppendChild(clone);
                        }
                    }
                }
                else
                {
                    _nodesetUANodeSetNode.AppendChild(nodesetImportedNode);
                }
            }
        }

        void createNodesetObjectTypes()
        {
            foreach (XmlNode xmiNode in _xmiClassList)
            {
                XmlAttribute noOTAttribute = xmiNode.Attributes["noObjectType"];
                if (noOTAttribute != null)
                { // ignore this classes
                    continue;
                }

                if (xmiNode.Attributes["name"] == null)
                { // ignore this packages without name
                    continue;
                }

                XmlAttribute xmiGenerate = xmiNode.Attributes["XMitoNodeset-Generate"];
                if (xmiGenerate != null)
                { 
                    if (xmiGenerate.Value != _xmiGenerate)
                    {
                        continue;
                    }
                }

                string id = xmiNode.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value;
                mergeXmiNodes(xmiNode, null, id);

                string nodeId = getNodeId(String.Format("ns=1;s={0}", id));
                string name = xmiNode.Attributes["name"].Value;
                XmlNode nodesetObjectTypeNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "UAObjectType");
                addXmlAttribute(_nodesetDoc, nodesetObjectTypeNode, "NodeId", nodeId);
                addXmlAttribute(_nodesetDoc, nodesetObjectTypeNode, "BrowseName", String.Format("1:{0}", name));
                XmlAttribute xmiIsAbstract = xmiNode.Attributes["isAbstract"];
                if (xmiIsAbstract != null)
                {
                    if (xmiIsAbstract.Value == "true")
                    {
                        addXmlAttribute(_nodesetDoc, nodesetObjectTypeNode, "IsAbstract", "true");
                    }
                }

                XmlNode nodesetDisplayNameNode = addXmlElement(_nodesetDoc, nodesetObjectTypeNode, "DisplayName", name);

                XmlNode nodesetReferencesNode = addXmlElement(_nodesetDoc, nodesetObjectTypeNode, "References");

                // base class
                XmlNode baseClasNode = getBaseClass(xmiNode);
                string baseClassNodeId;
                string baseClassName;
                if (baseClasNode == null)
                {
                    baseClassNodeId = "i=58";   // OPC UA BaseObjectType
                    baseClassName = "BaseObjectType";
                }
                else
                {
                    baseClassName = baseClasNode.Attributes["name"].Value;
                    baseClassNodeId = getNodeId(String.Format("ns=1;s={0}", baseClasNode.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value), baseClassName, true);
                }

                XmlNode nodesetBackwardHasSubtypeNode = addXmlElement(_nodesetDoc, nodesetReferencesNode, "Reference", baseClassNodeId);
                addXmlAttribute(_nodesetDoc, nodesetBackwardHasSubtypeNode, "ReferenceType", "HasSubtype");
                addXmlAttribute(_nodesetDoc, nodesetBackwardHasSubtypeNode, "IsForward", "false");

                XmlNode xmiElementNode = getXmiElementEA(id);
                if (xmiElementNode != null)
                { 
                    string extension = getTag(xmiElementNode, "Extension");
                    addNodesetExtentsion(nodesetObjectTypeNode, extension);
                }

                // documentation
                Paragraph parTable = null;
                if (_wordDoc != null)
                {
                    Paragraph p2 = _wordDoc.Paragraphs.Add();
                    p2.Range.Font.Name = "Arial";
                    p2.Range.Font.Size = 10F;
                    p2.Range.Text = name;
                    p2.Range.InsertParagraphAfter();

                    parTable = _wordDoc.Paragraphs.Add();
                    _wordCurrentTable = _wordDoc.Tables.Add(parTable.Range, 5, 6);

                    _wordCurrentTable.Range.Font.Name = "Arial";
                    _wordCurrentTable.Range.Font.Size = 8F;
                    _wordCurrentTable.Range.Font.Bold = 0;

                    _wordCurrentTable.Columns[1].Width = 75;
                    _wordCurrentTable.Columns[2].Width = 60;
                    _wordCurrentTable.Columns[3].Width = 70;
                    _wordCurrentTable.Columns[4].Width = 95;
                    _wordCurrentTable.Columns[5].Width = 95;
                    _wordCurrentTable.Columns[6].Width = 70;

                    _wordCurrentTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                    _wordCurrentTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
                }

                // instance declaration objects
                addNodesetObjectTypeInstanceDeclarationObjects(nodesetReferencesNode, xmiNode, nodeId, false);

                if (parTable != null)
                {
                    _wordCurrentTable.Rows[1].Range.Font.Bold = 1;
                    _wordCurrentTable.Cell(1,1).Range.Text = "Attribute";
                    _wordCurrentTable.Rows[1].Cells[2].Merge(_wordCurrentTable.Rows[1].Cells[6]);
                    _wordCurrentTable.Cell(1,2).Range.Text = "Value";

                    _wordCurrentTable.Cell(2,1).Range.Text = "BrowseName";
                    _wordCurrentTable.Rows[2].Cells[2].Merge(_wordCurrentTable.Rows[2].Cells[6]);
                    _wordCurrentTable.Cell(2,2).Range.Text = name;

                    _wordCurrentTable.Cell(3,1).Range.Text = "IsAbstract";
                    _wordCurrentTable.Rows[3].Cells[2].Merge(_wordCurrentTable.Rows[3].Cells[6]);
                    _wordCurrentTable.Cell(3,2).Range.Text = "false";

                    _wordCurrentTable.Rows[4].Range.Font.Bold = 1;
                    _wordCurrentTable.Cell(4,1).Range.Text = "Reference";
                    _wordCurrentTable.Cell(4,2).Range.Text = "NodeClass";
                    _wordCurrentTable.Cell(4,3).Range.Text = "BrowseName";
                    _wordCurrentTable.Cell(4,4).Range.Text = "DataType";
                    _wordCurrentTable.Cell(4,5).Range.Text = "TypeDefinition";
                    _wordCurrentTable.Cell(4,6).Range.Text = "ModellingRule";		 			

                    _wordCurrentTable.Rows[5].Cells[1].Merge(_wordCurrentTable.Rows[5].Cells[6]);
                    _wordCurrentTable.Cell(5,1).Range.Text = "Subtype of " + baseClassName;

                    _wordCurrentTable = null;
//                    parTable.Range.InsertParagraphAfter();
                }
            }
        }

        XmlNode getBaseClass(XmlNode node)
        {
            XmlNode baseClass = null;

            foreach (XmlNode subNode in node.ChildNodes)
            {
                if (subNode.Name == "generalization")
                {
                    XmlAttribute baseClassId = subNode.Attributes["general"];
                    if (baseClassId != null)
                    {
                        baseClass = getXmiClass(baseClassId.Value);

                        if (baseClass == null)
                        {
                            baseClass = getXmiStubEA(baseClassId.Value);
                            if (baseClass != null)
                            {
                                XmlAttribute useAttribute = baseClass.Attributes["useClass"];
                                if (useAttribute != null)
                                {
                                    baseClass = getXmiClass(useAttribute.Value);
                                }
                            }
                        }

                        return baseClass;
                    }
                }
            }

            return baseClass;
        }

        string getTag(string id, string key)
        {
            string value = null;
            XmlNode node = getXmiElementEA(id);
            if (node != null)
            {
                value = getTag(node, key);
            }           
            return value;
        }

        string getTag(XmlNode node, string key)
        {
            string value = null;
            foreach (XmlNode operationsSubNode in node.ChildNodes)
            {
                if (operationsSubNode.Name == "tags")
                {
                    foreach (XmlNode operationsTag in operationsSubNode.ChildNodes)
                    {
                        if (operationsTag.Attributes["name"].Value == key)
                        {
                            value = operationsTag.Attributes["value"].Value;
                        }
                    }
                    break;
                }
            }
            return value;
        }

        void addNodesetObjectTypeInstanceDeclarationObjects(XmlNode nodesetNode, XmlNode xmiNode, string partentNodeId, bool isAggregation)
        {
            if (isAggregation)
            { // add mandatory members base classes
                XmlNode baseClassNode = getBaseClass(xmiNode);
                if (baseClassNode != null)
                {
                    addNodesetObjectTypeInstanceDeclarationObjects(nodesetNode, baseClassNode, partentNodeId, true);
                }
            }

            foreach (XmlNode ownedNode in xmiNode.ChildNodes)
            {
                if (ownedNode.Name == "ownedAttribute")
                {
                    XmlAttribute umlType = ownedNode.Attributes[String.Format("{0}:type", _xmiNSPräfix)];
                    if (umlType != null)
                    {
                        if (umlType.Value == String.Format("{0}:Property", _umlNSPräfix))
                        {
                            bool doIgnore = false;
                            foreach (string ignore in _ignoreClassMemberList)
                            {
                                if (ownedNode.Attributes["name"].Value == ignore)
                                {
                                    doIgnore = true;
                                }
                            }
                            if (doIgnore)
                            {
                                continue;
                            }
                            
                            string id = ownedNode.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value;
                            if (getXmiIgnoredOwnedAttributeEA(id) != null)
                            {
                                continue;
                            }
                            string nodeId = getNodeId(String.Format("ns=1;s={0}|{1}", ownedNode.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value, partentNodeId));
                            string lowerValue = "";
                            string upperValue = "";
                            bool isPlaceholder = false;
                            bool isArray = false;

                            // add component node
                            foreach (XmlNode valueNode in ownedNode.ChildNodes)
                            {
                                if (valueNode.Name == "lowerValue")
                                {
                                    lowerValue = valueNode.Attributes["value"].Value;
                                }
                                if (valueNode.Name == "upperValue")
                                {
                                    upperValue = valueNode.Attributes["value"].Value;
                                    isArray = (valueNode.Attributes["isArray"] != null);
                                }
                            }
                            foreach (XmlNode typeNode in ownedNode.ChildNodes)
                            {
                                if (typeNode.Name == "type")
                                {
                                    if ((isAggregation) && (lowerValue == "0"))
                                    { // don't add optional objects on aggegation
                                         break;
                                    }

                                    string modelingRule;
                                    string name;
                                    string refTypeToUse = "";

                                    if (lowerValue == "0")
                                    {
                                        if (upperValue == "1")
                                        {
                                            modelingRule = "Optional";
                                        }
                                        else
                                        {
                                            if (!isArray)
                                            {
                                                modelingRule = "OptionalPlaceholder";
                                                isPlaceholder = true;
                                            }
                                            else
                                            {
                                                modelingRule = "Optional";
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (upperValue == "1")
                                        {
                                            modelingRule = "Mandatory";
                                        }
                                        else
                                        {
                                            if (!isArray)
                                            {
                                                modelingRule = "MandatoryPlaceholder";
                                                isPlaceholder = true;
                                            }
                                            else
                                            {
                                                modelingRule = "Mandatory";
                                            }
                                        }
                                    }

                                    XmlAttribute xmiIdref = typeNode.Attributes[String.Format("{0}:idref", _xmiNSPräfix)];
                                    if (xmiIdref != null)
                                    {
                                        XmlNode xmiClass = getXmiClass(xmiIdref.Value);
                                        XmlNode xmiVariable = getXmiStubEA(xmiIdref.Value);

                                        if ((xmiClass == null) && (xmiVariable != null))
                                        { 
                                            XmlAttribute useAttribute = xmiVariable.Attributes["useClass"];
                                            if (useAttribute != null)
                                            {
                                                try
                                                {
                                                    xmiClass = getXmiClass(useAttribute.Value);
                                                }
                                                catch
                                                { }
                                            }
                                        }

                                        
                                        if (!isPlaceholder)
                                        {
                                            name = ownedNode.Attributes["name"].Value;
                                        }
                                        else
                                        {
                                            name = String.Format("<{0}>", ownedNode.Attributes["name"].Value);
                                        }

                                        XmlNode xmiAttributesNode = getXmiAttributeEA(id);

                                        // get reference type
                                        if (xmiNode.Attributes["XMitoNodeset-ClassRefUseStereotype"] != null)
                                        { 
                                            if (xmiAttributesNode != null)
                                            {
                                                foreach (XmlNode xmiAttributesNodeChild in xmiAttributesNode.ChildNodes)
                                                {
                                                    if (xmiAttributesNodeChild.Name == "stereotype")
                                                    {
                                                        XmlAttribute xmiAttributesNodeChildSterotype = xmiAttributesNodeChild.Attributes["stereotype"];
                                                        if (xmiAttributesNodeChildSterotype != null)
                                                        {
                                                            refTypeToUse = xmiAttributesNodeChildSterotype.Value;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        string extension = null;
                                        if (xmiAttributesNode != null)
                                        { 
                                            extension = getTag(xmiAttributesNode, "Extension");
                                        }

                                        if (xmiClass != null)
                                        { // reference to object
                                            if (refTypeToUse == "")
                                            {
                                                refTypeToUse = "DefaultObjectRefType";
                                            }

                                            XmlNode nodesetComponentReferenceNode = addXmlElement(_nodesetDoc, nodesetNode, "Reference", nodeId);
                                            addXmlAttribute(_nodesetDoc, nodesetComponentReferenceNode, "ReferenceType", refTypeToUse);
                                            
                                            XmlNode nodesetVariableNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "UAObject");
                                            addXmlAttribute(_nodesetDoc, nodesetVariableNode, "NodeId", nodeId);
                                            addXmlAttribute(_nodesetDoc, nodesetVariableNode, "BrowseName", String.Format("1:{0}", name));
                                            addXmlAttribute(_nodesetDoc, nodesetVariableNode, "ParentNodeId", partentNodeId);

                                            addXmlElement(_nodesetDoc, nodesetVariableNode, "DisplayName", name);

                                            XmlNode nodesetReferencesNode = addXmlElement(_nodesetDoc, nodesetVariableNode, "References");
                                            XmlNode nodesetRefParentNode = addXmlElement(_nodesetDoc, nodesetReferencesNode, "Reference", partentNodeId);
                                            addXmlAttribute(_nodesetDoc, nodesetRefParentNode, "ReferenceType", refTypeToUse);
                                            addXmlAttribute(_nodesetDoc, nodesetRefParentNode, "IsForward", "false");

                                            XmlNode nodesetRefTypeDefNode = addXmlElement(_nodesetDoc, nodesetReferencesNode, "Reference", getNodeId(String.Format("ns=1;s={0}", xmiClass.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value), xmiClass.Attributes["name"].Value, true));
                                            addXmlAttribute(_nodesetDoc, nodesetRefTypeDefNode, "ReferenceType", "HasTypeDefinition");
                                            XmlNode nodesetRefModelingRuleNode = addXmlElement(_nodesetDoc, nodesetReferencesNode, "Reference", modelingRule);
                                            addXmlAttribute(_nodesetDoc, nodesetRefModelingRuleNode, "ReferenceType", "HasModellingRule");

                                            if (!isAggregation)
                                            {
                                                addNodesetExtentsion(nodesetVariableNode, extension);
                                            }
                                            addNodesetObjectTypeInstanceDeclarationObjects(nodesetReferencesNode, xmiClass, nodeId, true);

                                            // documentation
                                           if ((_wordCurrentTable != null) && (!isAggregation))
                                            {
                                                Row row = _wordCurrentTable.Rows.Add();
                                                row.Cells[1].Range.Text = refTypeToUse;
                                                row.Cells[2].Range.Text = "Object";
                                                row.Cells[3].Range.Text = name;
                                                row.Cells[4].Range.Text = "";
                                                row.Cells[5].Range.Text = xmiClass.Attributes["name"].Value;
                                                row.Cells[6].Range.Text = modelingRule;
                                            }
                                        }
                                        else
                                        { // reference to variable
                                            string dtForDoc = "";
                                            string dataTypeName = getDataTypeName(xmiIdref.Value, true, ref dtForDoc);

                                            if (refTypeToUse == "")
                                            {
                                                refTypeToUse = "DefaultVariableRefType";
                                            }

                                            XmlNode nodesetComponentReferenceNode = addXmlElement(_nodesetDoc, nodesetNode, "Reference", nodeId);
                                            addXmlAttribute(_nodesetDoc, nodesetComponentReferenceNode, "ReferenceType", refTypeToUse);

                                            XmlNode nodesetVariableNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable");
                                            addXmlAttribute(_nodesetDoc, nodesetVariableNode, "NodeId", nodeId);
                                            addXmlAttribute(_nodesetDoc, nodesetVariableNode, "BrowseName", String.Format("1:{0}", name));
                                            addXmlAttribute(_nodesetDoc, nodesetVariableNode, "ParentNodeId", partentNodeId);
                                            addXmlAttribute(_nodesetDoc, nodesetVariableNode, "DataType", dataTypeName);
                                            if (isArray)
                                            {
                                                addXmlAttribute(_nodesetDoc, nodesetVariableNode, "ValueRank", "1");
                                            }

                                            addXmlElement(_nodesetDoc, nodesetVariableNode, "DisplayName", name);

                                            XmlNode nodesetReferencesNode = addXmlElement(_nodesetDoc, nodesetVariableNode, "References");
                                            XmlNode nodesetRefParentNode = addXmlElement(_nodesetDoc, nodesetReferencesNode, "Reference", partentNodeId);
                                            addXmlAttribute(_nodesetDoc, nodesetRefParentNode, "ReferenceType", refTypeToUse);
                                            addXmlAttribute(_nodesetDoc, nodesetRefParentNode, "IsForward", "false");

                                            string variableType = "i=63"; // BaseVariableType
                                            if (refTypeToUse == "HasProperty")
                                            { 
                                               variableType = "i=68";     // PropertyType
                                            }
                                            string tagHasTypeDef = null;
                                            if (xmiAttributesNode != null)
                                            { 
                                                tagHasTypeDef = getTag(xmiAttributesNode, "HasTypeDefinition");
                                                if (tagHasTypeDef != null)
                                                { // other variable type
                                                    variableType = tagHasTypeDef;
                                                }
                                            }
                                            addXmlElementAndOneAttribute(_nodesetDoc, nodesetReferencesNode, "Reference", variableType, "ReferenceType", "HasTypeDefinition");
                                            XmlNode nodesetRefModelingRuleNode = addXmlElement(_nodesetDoc, nodesetReferencesNode, "Reference", modelingRule);
                                            addXmlAttribute(_nodesetDoc, nodesetRefModelingRuleNode, "ReferenceType", "HasModellingRule");

                                            addNodesetExtentsion(nodesetVariableNode, extension);

                                            if (tagHasTypeDef != null)
                                            { // VariableType - check for madatory elements
                                                XmlNode varBaseClassNode = getXmiClassByName(tagHasTypeDef);
                                                if (varBaseClassNode != null)
                                                {
                                                    string tagNodeClass = getTag(varBaseClassNode.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value, "NodeClass");
                                                    if (tagNodeClass != null)
                                                    { 
                                                        if (tagNodeClass == "VariableType")
                                                        {
                                                           addNodesetObjectTypeInstanceDeclarationObjects(nodesetReferencesNode, varBaseClassNode, nodeId, true);
                                                        }
                                                    }
                                                }
                                            } 

                                            // documentation
                                            if ((_wordCurrentTable != null) && (!isAggregation))
                                            {
                                                Row row = _wordCurrentTable.Rows.Add();
                                                row.Cells[1].Range.Text = refTypeToUse;
                                                row.Cells[2].Range.Text = "Variable";
                                                row.Cells[3].Range.Text = name;
                                                row.Cells[4].Range.Text = dtForDoc;
                                                if (variableType == "i=63")
                                                { 
                                                    row.Cells[5].Range.Text = "BaseVariableType";
                                                }
                                                else if (variableType == "i=68")
                                                { 
                                                    row.Cells[5].Range.Text = "PropertyType";
                                                }
                                                else
                                                {
                                                    row.Cells[5].Range.Text = variableType;
                                                }
                                                row.Cells[6].Range.Text = modelingRule;
                                            }
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
                if (ownedNode.Name == "ownedOperation")
                {
                    string name = ownedNode.Attributes["name"].Value;
                    string id = ownedNode.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value;
                    string nodeId = getNodeId(String.Format("ns=1;s={0}|{1}", ownedNode.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value, partentNodeId));

                    XmlNode operationNode = null;
                    try
                    {
                        operationNode = _xmiOperationsMapEA[id];
                    }
                    catch
                    { }

                    string modelingRule = "Mandatory";
                    if (operationNode != null)
                    {
                        string mrTag = getTag(operationNode, "ModelingRule");
                        if (mrTag != null)
                        {
                            modelingRule = mrTag;
                        }
                    }

                    if ((isAggregation) && (modelingRule != "Mandatory"))
                    { // don't add optional objects on aggegation
                        continue;
                    }

                    addXmlElementAndOneAttribute(_nodesetDoc, nodesetNode, "Reference", nodeId,  "ReferenceType", "HasComponent");
                                           
                    XmlNode nodesetMethodNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "UAMethod");
                    addXmlAttribute(_nodesetDoc, nodesetMethodNode, "NodeId", nodeId);
                    addXmlAttribute(_nodesetDoc, nodesetMethodNode, "BrowseName", String.Format("1:{0}", name));
                    addXmlAttribute(_nodesetDoc, nodesetMethodNode, "ParentNodeId", partentNodeId);

                    addXmlElement(_nodesetDoc, nodesetMethodNode, "DisplayName", name);

                    XmlNode nodesetReferencesNode = addXmlElement(_nodesetDoc, nodesetMethodNode, "References");
                    addXmlElementAndTwoAttributes(_nodesetDoc, nodesetReferencesNode, "Reference", partentNodeId, "ReferenceType", "HasComponent", "IsForward", "false");
                    addXmlElementAndOneAttribute(_nodesetDoc, nodesetReferencesNode, "Reference", modelingRule, "ReferenceType", "HasModellingRule");

                    XmlNode inputArgumentsNode = _nodesetDoc.CreateElement("UAVariable");
                    XmlNode outputArgumentsNode = _nodesetDoc.CreateElement("UAVariable");

                    string nodeIdInput =  getNodeId(String.Format("{0}-InputArguments", nodeId));

                    addXmlAttribute(_nodesetDoc, inputArgumentsNode, "NodeId", nodeIdInput);
                    addXmlAttribute(_nodesetDoc, inputArgumentsNode, "BrowseName", "InputArguments");
                    addXmlAttribute(_nodesetDoc, inputArgumentsNode, "ParentNodeId", nodeId);
                    addXmlAttribute(_nodesetDoc, inputArgumentsNode, "DataType", "i=296");
                    addXmlAttribute(_nodesetDoc, inputArgumentsNode, "ValueRank", "1");

                    addXmlElement(_nodesetDoc, inputArgumentsNode, "DisplayName", "InputArguments");

                    XmlNode nodesetInputReferencesNode = addXmlElement(_nodesetDoc, inputArgumentsNode, "References");
                    addXmlElementAndTwoAttributes(_nodesetDoc, nodesetInputReferencesNode, "Reference", nodeId, "ReferenceType", "HasProperty", "IsForward", "false");
                    addXmlElementAndOneAttribute(_nodesetDoc, nodesetInputReferencesNode, "Reference", "Mandatory", "ReferenceType", "HasModellingRule");
                    addXmlElementAndOneAttribute(_nodesetDoc, nodesetInputReferencesNode, "Reference", "i=68", "ReferenceType", "HasTypeDefinition");

                    XmlNode nodesetInputValueNode = addXmlElement(_nodesetDoc, inputArgumentsNode, "Value");
                    XmlNode nodesetInputValueLEONode = addXmlElementAndOneAttribute(_nodesetDoc, nodesetInputValueNode, "ListOfExtensionObject", "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");

                    string nodeIdOutput =  getNodeId(String.Format("{0}-OutputArguments", nodeId));

                    addXmlAttribute(_nodesetDoc, outputArgumentsNode, "NodeId", nodeIdOutput);
                    addXmlAttribute(_nodesetDoc, outputArgumentsNode, "BrowseName", "OutputArguments");
                    addXmlAttribute(_nodesetDoc, outputArgumentsNode, "ParentNodeId", nodeId);
                    addXmlAttribute(_nodesetDoc, outputArgumentsNode, "DataType", "i=296");
                    addXmlAttribute(_nodesetDoc, outputArgumentsNode, "ValueRank", "1");

                    addXmlElement(_nodesetDoc, outputArgumentsNode, "DisplayName", "OutputArguments");

                    XmlNode nodesetOutputReferencesNode = addXmlElement(_nodesetDoc, outputArgumentsNode, "References");
                    addXmlElementAndTwoAttributes(_nodesetDoc, nodesetOutputReferencesNode, "Reference", nodeId, "ReferenceType", "HasProperty", "IsForward", "false");
                    addXmlElementAndOneAttribute(_nodesetDoc, nodesetOutputReferencesNode, "Reference", "Mandatory", "ReferenceType", "HasModellingRule");
                    addXmlElementAndOneAttribute(_nodesetDoc, nodesetOutputReferencesNode, "Reference", "i=68", "ReferenceType", "HasTypeDefinition");

                    XmlNode nodesetOutputValueNode = addXmlElement(_nodesetDoc, outputArgumentsNode, "Value");
                    XmlNode nodesetOutputValueLEONode = addXmlElementAndOneAttribute(_nodesetDoc, nodesetOutputValueNode, "ListOfExtensionObject", "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");


                    foreach (XmlNode methodParameterNode in ownedNode.ChildNodes)
                    {
                        if (methodParameterNode.Name == "ownedParameter")
                        {
                            XmlNode nodesetInputValueEONode = null; 

                            if (methodParameterNode.Attributes["direction"].Value == "in")
                            {
                                nodesetInputValueEONode = addXmlElement(_nodesetDoc, nodesetInputValueLEONode, "ExtensionObject");
                            }
                            else if (methodParameterNode.Attributes["direction"].Value == "out")
                            {
                                nodesetInputValueEONode = addXmlElement(_nodesetDoc, nodesetOutputValueLEONode, "ExtensionObject");
                            } 
                            else
                            {
                                break;
                            }

                            string dummy = "";
                            string dataTypeName = getDataTypeName(methodParameterNode.Attributes["type"].Value, false, ref dummy);

                            XmlNode nodesetInputValueTypeIdNode = addXmlElement(_nodesetDoc, nodesetInputValueEONode, "TypeId");
                            XmlNode nodesetInputValueId = addXmlElement(_nodesetDoc, nodesetInputValueTypeIdNode, "Identifier", "i=297");
                            XmlNode nodesetInputBodyNode = addXmlElement(_nodesetDoc, nodesetInputValueEONode, "Body");
                            XmlNode nodesetArgNode = addXmlElement(_nodesetDoc, nodesetInputBodyNode, "Argument");
                            addXmlElement(_nodesetDoc, nodesetArgNode, "Name", methodParameterNode.Attributes["name"].Value);
                            XmlNode nodesetArgDTNode = addXmlElement(_nodesetDoc, nodesetArgNode, "DataType");
                            addXmlElement(_nodesetDoc, nodesetArgDTNode, "Identifier", dataTypeName);
                            addXmlElement(_nodesetDoc, nodesetArgNode, "ValueRank", "-1");
                            addXmlElement(_nodesetDoc, nodesetArgNode, "ArrayDimensions");
                            XmlNode nodesetArgDesNode = addXmlElement(_nodesetDoc, nodesetArgNode, "Description");
                            addXmlElement(_nodesetDoc, nodesetArgDesNode, "Locale");
                            addXmlElement(_nodesetDoc, nodesetArgDesNode, "Text");
                        }
                    }

                    if (nodesetInputValueLEONode.ChildNodes.Count > 0)
                    {
                        addXmlElementAndOneAttribute(_nodesetDoc, nodesetReferencesNode, "Reference", nodeIdInput, "ReferenceType", "HasProperty");
                        _nodesetUANodeSetNode.AppendChild(inputArgumentsNode);
                    }
                    if (nodesetOutputValueLEONode.ChildNodes.Count > 0)
                    {
                        addXmlElementAndOneAttribute(_nodesetDoc, nodesetReferencesNode, "Reference", nodeIdOutput, "ReferenceType", "HasProperty");
                        _nodesetUANodeSetNode.AppendChild(outputArgumentsNode);
                    }

                    // documentation
                    if ((_wordCurrentTable != null) && (!isAggregation))
                    {
                        Row row = _wordCurrentTable.Rows.Add();
                        row.Cells[1].Range.Text = "HasComponent";
                        row.Cells[2].Range.Text = "Method";
                        row.Cells[3].Range.Text = name;
                        row.Cells[4].Range.Text = "";
                        row.Cells[5].Range.Text = name + "Method";
                        row.Cells[6].Range.Text = modelingRule;
                    }
                }
            }
        }

        string getDataTypeName(string id, bool aliasAllowed, ref string readableName)
        {
            string dataTypeName = "?#?";
            bool doAddAlias = true;

            XmlNode xmiVariable = getXmiStubEA(id);
            XmlNode xmiStructure =getXmiStructure(id);
            XmlNode xmiEnumeration = getXmiEnumeration(id); 

            if (xmiVariable != null)
            {
                dataTypeName = xmiVariable.Attributes["name"].Value;
                XmlAttribute isDTAttribute = xmiVariable.Attributes["isDatatype"];
                if (isDTAttribute != null)
                {
                    dataTypeName = isDTAttribute.Value;
                    doAddAlias = false;
                    readableName = dataTypeName;
                }
                XmlAttribute useSAttribute = xmiVariable.Attributes["useStructure"];
                if (useSAttribute != null)
                {
                    XmlNode xmiStruct = getXmiStructure(useSAttribute.Value);
                    if (xmiStruct != null)
                    { 
                        dataTypeName = getNodeId(String.Format("ns=1;s={0}", xmiStruct.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value));
                        doAddAlias = false;
                        readableName =  xmiStruct.Attributes["name"].Value;
                    }
                }
                XmlAttribute useEAttribute = xmiVariable.Attributes["useEnumeration"];
                if (useEAttribute != null)
                {
                    XmlNode xmiEnum = getXmiEnumeration(useEAttribute.Value);
                    if (xmiEnum != null)
                    { 
                        dataTypeName = getNodeId(String.Format("ns=1;s={0}", xmiEnum.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value));
                        doAddAlias = false;
                        readableName =  xmiEnum.Attributes["name"].Value;
                    }
                }
            }
            else if (xmiStructure != null)
            {
                dataTypeName = getNodeId(String.Format("ns=1;s={0}", xmiStructure.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value), xmiStructure.Attributes["name"].Value, aliasAllowed);
                doAddAlias = false;
                readableName =  xmiStructure.Attributes["name"].Value;
            }
            else if (xmiEnumeration != null)
            {
                dataTypeName = getNodeId(String.Format("ns=1;s={0}", xmiEnumeration.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value), xmiEnumeration.Attributes["name"].Value, aliasAllowed);
                doAddAlias = false;
                readableName =  xmiEnumeration.Attributes["name"].Value;
            }

            if (doAddAlias)
            { 
                addAlias(dataTypeName, "i=1");
                if (!aliasAllowed)
                {
                    dataTypeName = getAlias(dataTypeName);
                }
            }

            return dataTypeName;
        }

        void createNodesetEnumerationDataTypes()
        {
            foreach (XmlNode xmiNode in _xmiEnumerationList)
            {
                XmlAttribute noOTAttribute = xmiNode.Attributes["noDataType"];
                if (noOTAttribute != null)
                { // ignore this classes
                    continue;
                }

                XmlAttribute xmiGenerate = xmiNode.Attributes["XMitoNodeset-Generate"];
                if (xmiGenerate != null)
                { 
                    if (xmiGenerate.Value != _xmiGenerate)
                    {
                        continue;
                    }
                }

                mergeXmiNodes(xmiNode, null, xmiNode.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value);

                _nodesetHasDT = true;

                string nodeId = getNodeId(String.Format("ns=1;s={0}", xmiNode.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value));
                string enumValuesNodeId = getNodeId(String.Format("ns=1;s={0}#EnumValues", xmiNode.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value));
                string exoValue;
                string exoText;
                string exoDescription;
                string enumName = xmiNode.Attributes["name"].Value;

                // nodeset
                XmlNode nodesetDataTypeNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UADataType", "NodeId", nodeId, "BrowseName", String.Format("1:{0}", enumName));
                addXmlElement(_nodesetDoc, nodesetDataTypeNode, "DisplayName", enumName);

                XmlNode nodesetReferencesNode = addXmlElement(_nodesetDoc, nodesetDataTypeNode, "References");
                addXmlElementAndTwoAttributes(_nodesetDoc, nodesetReferencesNode, "Reference", "Enumeration", "ReferenceType", "HasSubtype", "IsForward", "false");
                addXmlElementAndOneAttribute(_nodesetDoc, nodesetReferencesNode, "Reference", enumValuesNodeId, "ReferenceType", "HasProperty");
                XmlNode nodesetDefinitionNode = addXmlElementAndOneAttribute(_nodesetDoc, nodesetDataTypeNode, "Definition", "Name", enumName);


                XmlNode nodesetEnumValuesNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", enumValuesNodeId, "BrowseName", "1:EnumValues");
                addXmlAttribute(_nodesetDoc, nodesetEnumValuesNode, "DataType", "i=7594");
                addXmlAttribute(_nodesetDoc, nodesetEnumValuesNode, "ValueRank", "1");
                addXmlElement(_nodesetDoc, nodesetEnumValuesNode, "DisplayName", "EnumValues");

                XmlNode nodesetReferencesEVNode = addXmlElement(_nodesetDoc, nodesetEnumValuesNode, "References");
                addXmlElementAndOneAttribute(_nodesetDoc, nodesetReferencesEVNode, "Reference", "PropertyType", "ReferenceType", "HasTypeDefinition");
                addXmlElementAndOneAttribute(_nodesetDoc, nodesetReferencesEVNode, "Reference", "Mandatory", "ReferenceType", "HasModellingRule");
                addXmlElementAndTwoAttributes(_nodesetDoc, nodesetReferencesEVNode, "Reference", nodeId, "ReferenceType", "HasProperty", "IsForward", "false");
                XmlNode nodesetValueNode = addXmlElement(_nodesetDoc, nodesetEnumValuesNode, "Value");
                XmlNode nodesetListExONode = addXmlElementAndOneAttribute(_nodesetDoc, nodesetValueNode, "ListOfExtensionObject", "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");

                // binary types
                XmlNode binaryEnumTypeNode = addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, _binaryTypesRootNode, "opc", "http://opcfoundation.org/BinarySchema/", "EnumeratedType", "Name", enumName, "LengthInBits", "32");

                // xml types
                XmlNode xmlEnumTypeNode = addQualifiedXmlElementAndOneAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "simpleType", "name", enumName);
                XmlNode xmlEnumTypeNodeRes = addQualifiedXmlElementAndOneAttribute(_xmlTypesDoc, xmlEnumTypeNode, "xs", "http://www.w3.org/2001/XMLSchema", "restriction", "base", "xs:string");
                addQualifiedXmlElementAndTwoAttributes(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "element", "name", enumName, "type", String.Format("{0}", enumName));
                XmlNode xmlComplexTypeNode = addQualifiedXmlElementAndOneAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "complexType", "name", String.Format("ListOf{0}", enumName));
                XmlNode xmlComplexTypeNodeSeq = addQualifiedXmlElement(_xmlTypesDoc, xmlComplexTypeNode, "xs", "http://www.w3.org/2001/XMLSchema", "sequence");
                XmlNode xmlComplexTypeNodeEl = addQualifiedXmlElement(_xmlTypesDoc, xmlComplexTypeNodeSeq, "xs", "http://www.w3.org/2001/XMLSchema", "element");
                addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl, "name", enumName);
                addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl, "type", String.Format("{0}", enumName));
                addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl, "minOccurs", "0");
                addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl, "maxOccurs", "unbounded");
                XmlNode xmlComplexTypeNodeEl2 = addQualifiedXmlElement(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "element");
                addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl2, "name", String.Format("ListOf{0}", enumName));
                addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl2, "type", String.Format("ListOf{0}", enumName));
                addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl2, "nillable", "true");

                foreach (XmlNode ownedLiteralNode in xmiNode.ChildNodes)
                {
                    if (ownedLiteralNode.Name == "ownedLiteral")
                    {
                        XmlNode attributeNode = null;
                        exoText = ownedLiteralNode.Attributes["name"].Value;
                        exoValue = "";
                        exoDescription = "";

                        try
                        {
                            string literalId = ownedLiteralNode.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value;
                            attributeNode = _xmiAttributeMapEA[literalId];
                        }
                        catch { }

                        if (attributeNode != null)
                        {
                            foreach (XmlNode attrChildNode in attributeNode.ChildNodes)
                            {
                                if (attrChildNode.Name == "initial")
                                {
                                    XmlAttribute iniBodyAttribute = attrChildNode.Attributes["body"];
                                    if (iniBodyAttribute != null)
                                    {
                                        exoValue = iniBodyAttribute.Value;
                                    }
                                }
                                else if (attrChildNode.Name == "documentation")
                                {
                                    XmlAttribute docValueAttribute = attrChildNode.Attributes["value"];
                                    if (docValueAttribute != null)
                                    {
                                        exoDescription = docValueAttribute.Value;
                                    }
                                }
                            }
                        }

                        // nodeset
                        XmlNode nodesetFieldNode = addXmlElementAndTwoAttributes(_nodesetDoc, nodesetDefinitionNode, "Field", "Name", exoText, "Value", exoValue);
                        if (exoDescription.Length > 0)
                        {
                            addXmlElement(_nodesetDoc, nodesetFieldNode, "Description", exoDescription);
                        }
                        addEnumExtensionObject(_nodesetDoc, nodesetListExONode, exoValue, exoText, exoDescription);

                        // binary types
                       addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, binaryEnumTypeNode, "opc", "http://opcfoundation.org/BinarySchema/", "EnumeratedValue", "Name", exoText, "Value", exoValue);

                        // xml typess
                       addQualifiedXmlElementAndOneAttribute(_xmlTypesDoc, xmlEnumTypeNodeRes, "xs", "http://www.w3.org/2001/XMLSchema", "enumeration", "value", String.Format("{0}_{1}", exoText, exoValue));
                    }
                }
            }
        }

        void addEnumExtensionObject(XmlDocument doc, XmlNode node, string value, string text, string description)
        {
            XmlNode extensionObject = addXmlElement(_nodesetDoc, node, "ExtensionObject");
            XmlNode typeId = addXmlElement(_nodesetDoc, extensionObject, "TypeId");
            addXmlElement(_nodesetDoc, typeId, "Identifier", "i=7616");
            XmlNode body = addXmlElement(_nodesetDoc, extensionObject, "Body");
            XmlNode enumValueType = addXmlElement(_nodesetDoc, body, "EnumValueType");
            addXmlElement(_nodesetDoc, enumValueType, "Value", value);
            XmlNode displayName = addXmlElement(_nodesetDoc, enumValueType, "DisplayName");
            addXmlElement(_nodesetDoc, displayName, "Locale");
            addXmlElement(_nodesetDoc, displayName, "Text", text);
            XmlNode descriptionN = addXmlElement(_nodesetDoc, enumValueType, "Description");
            addXmlElement(_nodesetDoc, descriptionN, "Locale");
            addXmlElement(_nodesetDoc, descriptionN, "Text", description);
        }


        void createNodesetStructureDataTypes()
        {
            foreach (XmlNode xmiNode in _xmiStructureList)
            {
                XmlAttribute noOTAttribute = xmiNode.Attributes["noDataType"];
                if (noOTAttribute != null)
                { // ignore this classes
                    continue;
                }

                if (xmiNode.Attributes["name"] == null)
                { // ignore this packages without name
                    continue;
                }

                XmlAttribute xmiGenerate = xmiNode.Attributes["XMitoNodeset-Generate"];
                if (xmiGenerate != null)
                { 
                    if (xmiGenerate.Value != _xmiGenerate)
                    {
                        continue;
                    }
                }

                mergeXmiNodes(xmiNode, null, xmiNode.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value);

                 _nodesetHasDT = true;


                string structName = xmiNode.Attributes["name"].Value;

                string nodeId = getNodeId(String.Format("ns=1;s={0}", xmiNode.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value));

                // nodeset
                XmlNode nodesetDataTypeNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UADataType", "NodeId", nodeId, "BrowseName", String.Format("1:{0}", structName));
                addXmlElement(_nodesetDoc, nodesetDataTypeNode, "DisplayName", structName);

                XmlNode nodesetReferencesNode = addXmlElement(_nodesetDoc, nodesetDataTypeNode, "References");
                addXmlElementAndTwoAttributes(_nodesetDoc, nodesetReferencesNode, "Reference", "Structure", "ReferenceType", "HasSubtype", "IsForward", "false");
                XmlNode nodesetDefinitionNode = addXmlElementAndOneAttribute(_nodesetDoc, nodesetDataTypeNode, "Definition", "Name", structName);

                XmlNode nodesetSchemaEncoding = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAObject", "NodeId", getNodeId(String.Format("{0}-BinaryEnconding", nodeId)), "BrowseName", "Default Binary");
                addXmlAttribute(_nodesetDoc, nodesetSchemaEncoding, "SymbolicName", "Default Binary");
                addXmlElement(_nodesetDoc, nodesetSchemaEncoding, "DisplayName", "Default Binary");
                XmlNode nodesetSchemaEncodingReferencesNode = addXmlElement(_nodesetDoc, nodesetSchemaEncoding, "References");
                addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaEncodingReferencesNode, "Reference", "i=76", "ReferenceType", "HasTypeDefinition");
                addXmlElementAndTwoAttributes(_nodesetDoc, nodesetSchemaEncodingReferencesNode, "Reference", nodeId, "ReferenceType", "HasEncoding", "IsForward", "false");
                addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaEncodingReferencesNode, "Reference", getNodeId(String.Format("{0}-BinaryDescription", nodeId)), "ReferenceType", "HasDescription");
 
                XmlNode nodesetSchemaXmlEncoding = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAObject", "NodeId", getNodeId(String.Format("{0}-XMLEnconding", nodeId)), "BrowseName", "Default XML");
                addXmlAttribute(_nodesetDoc, nodesetSchemaXmlEncoding, "SymbolicName", "Default XML");
                addXmlElement(_nodesetDoc, nodesetSchemaXmlEncoding, "DisplayName", "Default XML");
                XmlNode nodesetSchemaXmlEncodingReferencesNode = addXmlElement(_nodesetDoc, nodesetSchemaXmlEncoding, "References");
                addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaXmlEncodingReferencesNode, "Reference", "i=76", "ReferenceType", "HasTypeDefinition");
                addXmlElementAndTwoAttributes(_nodesetDoc, nodesetSchemaXmlEncodingReferencesNode, "Reference", nodeId, "ReferenceType", "HasEncoding", "IsForward", "false");
                addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaXmlEncodingReferencesNode, "Reference", getNodeId(String.Format("{0}-XMLDescription", nodeId)), "ReferenceType", "HasDescription");
                                             
                // nodeset binary schema types
                XmlNode nodesetSchemaDescription = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", getNodeId(String.Format("{0}-BinaryDescription", nodeId)), "BrowseName", String.Format("1:{0}", structName));
                addXmlAttribute(_nodesetDoc, nodesetSchemaDescription, "ParentNodeId", getNodeId(_nodeIdTextBinarySchema));
                addXmlAttribute(_nodesetDoc, nodesetSchemaDescription, "DataType", "String");
                addXmlElement(_nodesetDoc, nodesetSchemaDescription, "DisplayName", structName);
                XmlNode nodesetSchemaDescriptionReferencesNode = addXmlElement(_nodesetDoc, nodesetSchemaDescription, "References");
                addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaDescriptionReferencesNode, "Reference", "i=69", "ReferenceType", "HasTypeDefinition");
                addXmlElementAndTwoAttributes(_nodesetDoc, nodesetSchemaDescriptionReferencesNode, "Reference", getNodeId(_nodeIdTextBinarySchema), "ReferenceType", "HasComponent", "IsForward", "false");
                XmlNode nodesetSchemaDescriptionValueNode = addXmlElement(_nodesetDoc, nodesetSchemaDescription, "Value");
                addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaDescriptionValueNode, "String", structName, "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");

                // nodeset XML schema types
                XmlNode nodesetSchemaXmlDescription = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", getNodeId(String.Format("{0}-XMLDescription", nodeId)), "BrowseName", String.Format("1:{0}", structName));
                addXmlAttribute(_nodesetDoc, nodesetSchemaXmlDescription, "ParentNodeId", getNodeId(_nodeIdTextXmlSchema));
                addXmlAttribute(_nodesetDoc, nodesetSchemaXmlDescription, "DataType", "String");
                addXmlElement(_nodesetDoc, nodesetSchemaXmlDescription, "DisplayName", structName);
                XmlNode nodesetSchemaXmlDescriptionReferencesNode = addXmlElement(_nodesetDoc, nodesetSchemaXmlDescription, "References");
                addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaXmlDescriptionReferencesNode, "Reference", "i=69", "ReferenceType", "HasTypeDefinition");
                addXmlElementAndTwoAttributes(_nodesetDoc, nodesetSchemaXmlDescriptionReferencesNode, "Reference", getNodeId(_nodeIdTextXmlSchema), "ReferenceType", "HasComponent", "IsForward", "false");
                XmlNode nodesetSchemaXmlDescriptionValueNode = addXmlElement(_nodesetDoc, nodesetSchemaXmlDescription, "Value");
                addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaXmlDescriptionValueNode, "String", String.Format("//xs:element[@name='{0}']", structName), "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");

                // binary types
                XmlNode binaryStructTypeNode = addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, _binaryTypesRootNode, "opc", "http://opcfoundation.org/BinarySchema/", "StructuredType", "Name", structName, "BaseType", "ua:ExtensionObject");

                // XML types
                XmlNode xmlStructTypeNode = addQualifiedXmlElementAndOneAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "complexType", "name", structName);
                XmlNode xmlStructTypeSqNode = addQualifiedXmlElement(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "sequence");
                addQualifiedXmlElementAndTwoAttributes(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "element", "name", structName, "type", String.Format("tns:{0}", structName));

                // count the number of optional fields
                int numOptionalFields=0;
                foreach (XmlNode ownedAttributeNode in xmiNode.ChildNodes)
                {
                    if (ownedAttributeNode.Name == "ownedAttribute")
                    {
                        XmlAttribute umlType = ownedAttributeNode.Attributes[String.Format("{0}:type", _xmiNSPräfix)];
                        if (umlType != null)
                        {
                            if (umlType.Value == String.Format("{0}:Property", _umlNSPräfix))
                            {
                                // add component node
                                foreach (XmlNode lowerValueNode in ownedAttributeNode.ChildNodes)
                                {
                                    if (lowerValueNode.Name == "lowerValue")
                                    {
                                        if (lowerValueNode.Attributes["value"].Value == "0")
                                        { 
                                            numOptionalFields++;
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                // binary types - add switch bits
                for (int i = 0; i < numOptionalFields; i++)
                {
                    addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, binaryStructTypeNode, "opc", "http://opcfoundation.org/BinarySchema/", "Field", "Name", String.Format("Bit{0}", i), "TypeName", "opc:Bit");
                }
                int remainingBitsInByte = 8 - (numOptionalFields % 8);
                if (remainingBitsInByte < 8) 
                {
                    XmlNode binaryRemainingBitsInByteNode = addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, binaryStructTypeNode, "opc", "http://opcfoundation.org/BinarySchema/", "Field", "Name", "Reserved", "TypeName", "opc:Bit");
                    addXmlAttribute(_binaryTypesDoc, binaryRemainingBitsInByteNode, "Length", String.Format("{0}", remainingBitsInByte));
                }

                int curOptionalField = -1;
                foreach (XmlNode ownedAttributeNode in xmiNode.ChildNodes)
                {
                    if (ownedAttributeNode.Name == "ownedAttribute")
                    {
                        XmlAttribute umlType = ownedAttributeNode.Attributes[String.Format("{0}:type", _xmiNSPräfix)];
                        string lowerValue = "1";

                        if (umlType != null)
                        {
                            if (umlType.Value == String.Format("{0}:Property", _umlNSPräfix))
                            {
                                // add component node
                                foreach (XmlNode lowerValueNode in ownedAttributeNode.ChildNodes)
                                {
                                    if (lowerValueNode.Name == "lowerValue")
                                    {
                                        lowerValue = lowerValueNode.Attributes["value"].Value;
                                        if (lowerValueNode.Attributes["value"].Value == "0")
                                        { 
                                            curOptionalField++;
                                        }
                                        break;
                                    }
                                }
                                foreach (XmlNode typeNode in ownedAttributeNode.ChildNodes)
                                {
                                    if (typeNode.Name == "type")
                                    {
                                        XmlAttribute xmiIdref = typeNode.Attributes[String.Format("{0}:idref", _xmiNSPräfix)];
                                        if (xmiIdref != null)
                                        {
                                            XmlNode xmiStructure = getXmiStructure(xmiIdref.Value);
                                            XmlNode xmiEnumeration = getXmiEnumeration(xmiIdref.Value);
                                            XmlNode xmiVariable = getXmiStubEA(xmiIdref.Value);

                                            if ((xmiStructure == null) && (xmiVariable != null))
                                            { 
                                                XmlAttribute useAttribute = xmiVariable.Attributes["useStructure"];
                                                if (useAttribute != null)
                                                {
                                                    xmiStructure = getXmiStructure(useAttribute.Value);
                                                }
                                            }

                                            if ((xmiEnumeration == null) && (xmiVariable != null))
                                            { 
                                                XmlAttribute useAttribute = xmiVariable.Attributes["useEnumeration"];
                                                if (useAttribute != null)
                                                {
                                                    xmiEnumeration = getXmiEnumeration(useAttribute.Value);
                                                }
                                            }

                                            string dataTypeNodeId = "?#?";
                                            string dataTypeBinary = "opc:?#?";
                                            string dataTypeXml = "xs:?#?";

                                            if (xmiStructure != null)
                                            {
                                                dataTypeNodeId = getNodeId(String.Format("ns=1;s={0}", xmiStructure.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value));
                                                dataTypeBinary = String.Format("tns:{0}", xmiStructure.Attributes["name"].Value);
                                                dataTypeXml = String.Format("tns:{0}", xmiStructure.Attributes["name"].Value);
                                            }
                                            else if (xmiEnumeration != null)
                                            {
                                                dataTypeNodeId = getNodeId(String.Format("ns=1;s={0}", xmiEnumeration.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value));
                                                dataTypeBinary = String.Format("tns:{0}", xmiEnumeration.Attributes["name"].Value);
                                                dataTypeXml = String.Format("tns:{0}", xmiEnumeration.Attributes["name"].Value);
                                            }
                                            else if (xmiVariable != null)
                                            {
                                                dataTypeNodeId = xmiVariable.Attributes["name"].Value;
                                                XmlAttribute isDTAttribute = xmiVariable.Attributes["isDatatype"];
                                                if (isDTAttribute != null)
                                                {
                                                    dataTypeNodeId = isDTAttribute.Value;
                                                    dataTypeBinary = getBinaryDatatypeName(dataTypeNodeId);
                                                    dataTypeXml = getXmlDatatypeName(dataTypeNodeId);
                                                }
                                            }

                                            // nodeset
                                            XmlNode nodesetFieldNode = addXmlElementAndTwoAttributes(_nodesetDoc, nodesetDefinitionNode, "Field", "Name", ownedAttributeNode.Attributes["name"].Value , "DataType", dataTypeNodeId);
                                            if (lowerValue == "0")
                                            {
                                                addXmlAttribute(_nodesetDoc, nodesetFieldNode, "IsOptional", "true");
                                            }
                                            
                                            // binary types
                                            XmlNode binaryFieldNode = addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, binaryStructTypeNode, "opc", "http://opcfoundation.org/BinarySchema/", "Field", "Name", ownedAttributeNode.Attributes["name"].Value, "TypeName", dataTypeBinary);
                                            if (lowerValue == "0")
                                            {
                                                addXmlAttribute(_binaryTypesDoc, binaryFieldNode, "SwitchField", String.Format("Bit{0}", curOptionalField));
                                                addXmlAttribute(_binaryTypesDoc, binaryFieldNode, "SwitchValue", "1");
                                            }

                                            // XML types
                                            XmlNode xmlFieldNode = addQualifiedXmlElementAndTwoAttributes(_xmlTypesDoc, xmlStructTypeSqNode, "xs", "http://www.w3.org/2001/XMLSchema", "element", "name", ownedAttributeNode.Attributes["name"].Value, "TypeName", dataTypeXml);
                                            if (lowerValue == "0")
                                            {
                                                addXmlAttribute(_xmlTypesDoc, xmlFieldNode, "minOccurs", "0");
                                                addXmlAttribute(_xmlTypesDoc, xmlFieldNode, "nillable", "true");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        void createNodesetTypeDictionary()
        {
            _binaryTypesDoc.Save(_binaryTypesFileName);
            _xmlTypesDoc.Save(_xmlTypesFileName);

            string binarySchemaUriNodeId = getNodeId("BinarySchema_NamespaceUri");
            XmlNode nodesetBinarySchemaNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", getNodeId(_nodeIdTextBinarySchema), "BrowseName", String.Format("1:Opc.Ua.{0}", _nodesetTypeDictionaryName));
            addXmlAttribute(_nodesetDoc, nodesetBinarySchemaNode, "SymbolicName", String.Format("{0}_BinarySchema", _nodesetTypeDictionaryName));
            addXmlAttribute(_nodesetDoc, nodesetBinarySchemaNode, "DataType", "ByteString");
            addXmlElement(_nodesetDoc, nodesetBinarySchemaNode, "DisplayName", String.Format("Opc.Ua.{0}", _nodesetTypeDictionaryName));

            XmlNode nodesetBinarySchemaReferencesNode = addXmlElement(_nodesetDoc, nodesetBinarySchemaNode, "References");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetBinarySchemaReferencesNode, "Reference", "i=93","ReferenceType", "HasComponent", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetBinarySchemaReferencesNode, "Reference", binarySchemaUriNodeId, "ReferenceType", "HasProperty");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetBinarySchemaReferencesNode, "Reference", "i=72", "ReferenceType", "HasTypeDefinition");
            XmlNode nodesetBinarySchemaValueNode = addXmlElement(_nodesetDoc, nodesetBinarySchemaNode, "Value");

            XmlNode nodesetBinarySchemaValueBSNode = addXmlElementAndOneAttribute(_nodesetDoc, nodesetBinarySchemaValueNode, "ByteString", "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");

            byte[] binarySchemaData = Encoding.UTF8.GetBytes(_binaryTypesDoc.OuterXml);
            nodesetBinarySchemaValueBSNode.InnerText = Convert.ToBase64String(binarySchemaData);

            XmlNode nodesetBinarySchemaUriNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", binarySchemaUriNodeId, "BrowseName", "1:NamespaceUri");
            addXmlAttribute(_nodesetDoc, nodesetBinarySchemaUriNode, "ParentNodeId", getNodeId("BinarySchema"));
            addXmlAttribute(_nodesetDoc, nodesetBinarySchemaUriNode, "DataType", "String");
            addXmlElement(_nodesetDoc, nodesetBinarySchemaUriNode, "DisplayName", "NamespaceUri");
            addXmlElement(_nodesetDoc, nodesetBinarySchemaUriNode, "Description", "A URI that uniquely identifies the dictionary.");
            XmlNode nodesetBinarySchemaUriReferencesNode = addXmlElement(_nodesetDoc, nodesetBinarySchemaUriNode, "References");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetBinarySchemaUriReferencesNode, "Reference", getNodeId("BinarySchema"),"ReferenceType", "HasProperty", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetBinarySchemaUriReferencesNode, "Reference", "i=68", "ReferenceType", "HasTypeDefinition");
            XmlNode nodesetBinarySchemaUriValueNode = addXmlElement(_nodesetDoc, nodesetBinarySchemaUriNode, "Value");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetBinarySchemaUriValueNode, "String", _nodesetURL, "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");

            string xmlSchemaUriNodeId = getNodeId("XmlSchema_NamespaceUri");
            XmlNode nodesetxmlSchemaNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", getNodeId(_nodeIdTextXmlSchema), "BrowseName", String.Format("1:Opc.Ua.{0}", _nodesetTypeDictionaryName));
            addXmlAttribute(_nodesetDoc, nodesetxmlSchemaNode, "SymbolicName", String.Format("{0}_XmlSchema", _nodesetTypeDictionaryName));
            addXmlAttribute(_nodesetDoc, nodesetxmlSchemaNode, "DataType", "ByteString");
            addXmlElement(_nodesetDoc, nodesetxmlSchemaNode, "DisplayName", String.Format("Opc.Ua.{0}", _nodesetTypeDictionaryName));

            XmlNode nodesetxmlSchemaReferencesNode = addXmlElement(_nodesetDoc, nodesetxmlSchemaNode, "References");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetxmlSchemaReferencesNode, "Reference", "i=92","ReferenceType", "HasComponent", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetxmlSchemaReferencesNode, "Reference", xmlSchemaUriNodeId, "ReferenceType", "HasProperty");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetxmlSchemaReferencesNode, "Reference", "i=72", "ReferenceType", "HasTypeDefinition");
            XmlNode nodesetxmlSchemaValueNode = addXmlElement(_nodesetDoc, nodesetxmlSchemaNode, "Value");

            XmlNode nodesetxmlSchemaValueBSNode = addXmlElementAndOneAttribute(_nodesetDoc, nodesetxmlSchemaValueNode, "ByteString", "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");

            byte[] xmlSchemaData = Encoding.UTF8.GetBytes(_xmlTypesDoc.OuterXml);
            nodesetxmlSchemaValueBSNode.InnerText = Convert.ToBase64String(xmlSchemaData);

            XmlNode nodesetXmlSchemaUriNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", xmlSchemaUriNodeId, "BrowseName", "1:NamespaceUri");
            addXmlAttribute(_nodesetDoc, nodesetXmlSchemaUriNode, "ParentNodeId", getNodeId("XmlSchema"));
            addXmlAttribute(_nodesetDoc, nodesetXmlSchemaUriNode, "DataType", "String");
            addXmlElement(_nodesetDoc, nodesetXmlSchemaUriNode, "DisplayName", "NamespaceUri");
            addXmlElement(_nodesetDoc, nodesetXmlSchemaUriNode, "Description", "A URI that uniquely identifies the dictionary.");
            XmlNode nodesetXmlSchemaUriReferencesNode = addXmlElement(_nodesetDoc, nodesetXmlSchemaUriNode, "References");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetXmlSchemaUriReferencesNode, "Reference", getNodeId("XmlSchema"),"ReferenceType", "HasProperty", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetXmlSchemaUriReferencesNode, "Reference", "i=68", "ReferenceType", "HasTypeDefinition");
            XmlNode nodesetXmlSchemaUriValueNode = addXmlElement(_nodesetDoc, nodesetXmlSchemaUriNode, "Value");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetXmlSchemaUriValueNode, "String", _nodesetURL, "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");
        }

        void addAliases()
        {
            _aliases = new string[,] {
                { "Boolean", "i=1", "opc:Boolean", "xs:boolean" },
                { "SByte", "i=2", "opc:SByte", "xs:byte" },
                { "Byte", "i=3", "opc:SByte", "xs:unsignedByte" },
                { "Int16", "i=4", "opc:Int16", "xs:short" },
                { "UInt16", "i=5", "opc:UInt16", "xs:unsignedShort" },
                { "Int32", "i=6", "opc:Int32", "xs:int" },
                { "UInt32", "i=7", "opc:UInt32", "xs:unsignedInt" },
                { "Int64", "i=8", "opc:Int64", "xs:long" },
                { "UInt64", "i=9", "opc:UInt64", "xs:unsignedLong" },
                { "Float", "i=10", "opc:Float", "xs:float" },
                { "Double", "i=11", "opc:Double", "xs:double" },
                { "String", "i=12", "opc:String", "xs:string" },
                { "ByteString", "i=15", "opc:ByteString", "xs:base64Binary" },          
                { "Structure", "i=22", "", "" },     
                { "BaseDataType", "i=24", "", "" },   
                { "Enumeration", "i=29", "", "" },          
                { "Organizes", "i=35", "", "" },                        
                { "HasModellingRule", "i=37", "", "" },
                { "HasEncoding", "i=38", "", "" },
                { "HasDescription", "i=39", "", "" },
                { "HasTypeDefinition", "i=40", "", "" },
                { "HasSubtype", "i=45", "", "" },
                { "HasProperty", "i=46", "", "" },
                { "HasComponent", "i=47", "", "" },
                { "PropertyType", "i=68", "", "" },               
                { "Mandatory", "i=78", "", "" },
                { "Optional", "i=80", "", "" },
                { "OptionalPlaceholder", "i=11508", "", "" },
                { "MandatoryPlaceholder", "i=11510", "", "" },
                { "DefaultVariableRefType", "i=47", "", "" },
                { "DefaultObjectRefType", "i=47", "", "" },
            };

            _nodesetAliasesNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "Aliases");
            for (int i = 0; i < (_aliases.Length) / 4; i++)
            {
                XmlNode nodesetAliasNode = addXmlElement(_nodesetDoc, _nodesetAliasesNode, "Alias", _aliases[i,1]);
                addXmlAttribute(_nodesetDoc, nodesetAliasNode, "Alias", _aliases[i, 0]);
            }
        }



        string getBinaryDatatypeName(string dt)
        {
            for (int i = 0; i < (_aliases.Length) / 4; i++)
            {
                if (_aliases[i, 0] == dt)
                {
                    return _aliases[i, 2]; 
                }
            }
            return "";
        }

        string getXmlDatatypeName(string dt)
        {
            for (int i = 0; i < (_aliases.Length) / 4; i++)
            {
                if (_aliases[i, 0] == dt)
                {
                    return _aliases[i, 3]; 
                }
            }
            return "";
        }

        string getNodeId(string strNodeId)
        {
            string nodeId = null;
            try
            {
                nodeId = _nodesetNodeIdMap[strNodeId];
            }
            catch
            { }

            if (nodeId == null)
            {
                nodeId = String.Format("ns=1;i={0}", _nextNodeId);
                _nextNodeId++;
                _nodesetNodeIdMap.Add(strNodeId, nodeId);
            }

            return nodeId;
        }

        string getNodeId(string strNodeId, string name, bool aliasAllowed)
        {
            string nodeId = null;
            string alias = getAlias(name);
            if (alias == null)
            {
                nodeId = getNodeId(strNodeId);
            }
            else
            {
                if (aliasAllowed)
                {
                    nodeId = name;
                }
                else
                {
                    nodeId = alias;
                }
            }
            return nodeId;
        }

        string getAlias(string name)
        {
            string alias = null;
            foreach (XmlNode node in _nodesetAliasesNode.ChildNodes)
            {
                if (node.Attributes["Alias"].Value == name)
                {
                    alias = node.InnerText;
                }
            }
            return alias;
        }

        void addAlias(string name, string value)
        {
            foreach (XmlNode node in _nodesetAliasesNode.ChildNodes)
            {
                if (node.Attributes["Alias"].Value == name)
                {
                    return;
                }
            }

            // alias not found -> add it
            XmlNode nodesetAliasNode = addXmlElement(_nodesetDoc, _nodesetAliasesNode, "Alias", value);
            addXmlAttribute(_nodesetDoc, nodesetAliasNode, "Alias", name);
        }

        XmlAttribute addXmlAttribute(XmlDocument doc, XmlNode node, string name, string value)
        {
            XmlAttribute attr = doc.CreateAttribute(name);
            attr.Value = value;
            node.Attributes.Append(attr);
            return attr;
        }

        XmlAttribute addXmlAttributeDeep(XmlDocument doc, XmlNode node, string name, string value)
        {
            XmlAttribute attr = doc.CreateAttribute(name);
            attr.Value = value;
            node.Attributes.Append(attr);

            foreach (XmlNode child in node.ChildNodes)
            {
                addXmlAttributeDeep(doc, child, name, value);
            }
            return attr;
        }

        XmlNode addXmlElement(XmlDocument doc, XmlNode parent, string name, string innerText)
        {
            XmlNode node = addXmlElement(doc, parent, name);
            node.InnerText = innerText;
            return node;
        }

        XmlNode addXmlElement(XmlDocument doc, XmlNode parent, string name)
        {
            XmlNode node = doc.CreateElement(name);
            parent.AppendChild(node);
            return node;
        }

        XmlNode addXmlElementAndOneAttribute(XmlDocument doc, XmlNode parent, string elName, string attrName, string attrValue)
        {
            XmlNode node = addXmlElement(doc, parent, elName);
            addXmlAttribute(doc, node, attrName, attrValue);
            return node;
        }

        XmlNode addXmlElementAndOneAttribute(XmlDocument doc, XmlNode parent, string elName, string innerText, string attrName, string attrValue)
        {
            XmlNode node = addXmlElement(doc, parent, elName, innerText);
            addXmlAttribute(doc, node, attrName, attrValue);
            return node;
        }

        XmlNode addXmlElementAndTwoAttributes(XmlDocument doc, XmlNode parent, string elName, string attrName1, string attrValue1, string attrName2, string attrValue2)
        {
            XmlNode node = addXmlElement(doc, parent, elName);
            addXmlAttribute(doc, node, attrName1, attrValue1);
            addXmlAttribute(doc, node, attrName2, attrValue2);
            return node;
        }

        XmlNode addXmlElementAndTwoAttributes(XmlDocument doc, XmlNode parent, string elName, string innerText, string attrName1, string attrValue1, string attrName2, string attrValue2)
        {
            XmlNode node = addXmlElement(doc, parent, elName, innerText);
            addXmlAttribute(doc, node, attrName1, attrValue1);
            addXmlAttribute(doc, node, attrName2, attrValue2);
            return node;
        }

        XmlNode addQualifiedXmlElement(XmlDocument doc, XmlNode parent, string prefix, string uri, string name, string innerText)
        {
            XmlNode node = addQualifiedXmlElement(doc, parent, prefix, uri, name);
            node.InnerText = innerText;
            return node;
        }

        XmlNode addQualifiedXmlElement(XmlDocument doc, XmlNode parent, string prefix, string uri, string name)
        {
            XmlNode node = doc.CreateElement(prefix, name, uri);
            parent.AppendChild(node);
            return node;
        }

        XmlNode addQualifiedXmlElementAndOneAttribute(XmlDocument doc, XmlNode parent, string prefix, string uri, string elName, string attrName, string attrValue)
        {
            XmlNode node = addQualifiedXmlElement(doc, parent, prefix, uri, elName);
            addXmlAttribute(doc, node, attrName, attrValue);
            return node;
        }

        XmlNode addQualifiedXmlElementAndOneAttribute(XmlDocument doc, XmlNode parent, string prefix, string uri, string elName, string innerText, string attrName, string attrValue)
        {
            XmlNode node = addQualifiedXmlElement(doc, parent, prefix, uri, elName, innerText);
            addXmlAttribute(doc, node, attrName, attrValue);
            return node;
        }

        XmlNode addQualifiedXmlElementAndTwoAttributes(XmlDocument doc, XmlNode parent, string prefix, string uri, string elName, string attrName1, string attrValue1, string attrName2, string attrValue2)
        {
            XmlNode node = addQualifiedXmlElement(doc, parent, prefix, uri, elName);
            addXmlAttribute(doc, node, attrName1, attrValue1);
            addXmlAttribute(doc, node, attrName2, attrValue2);
            return node;
        }

        XmlNode addQualifiedXmlElementAndTwoAttributes(XmlDocument doc, XmlNode parent, string prefix, string uri, string elName, string innerText, string attrName1, string attrValue1, string attrName2, string attrValue2)
        {
            XmlNode node = addQualifiedXmlElement(doc, parent, prefix, uri, elName, innerText);
            addXmlAttribute(doc, node, attrName1, attrValue1);
            addXmlAttribute(doc, node, attrName2, attrValue2);
            return node;
        }

        void addNodesetExtentsion(XmlNode node, string extension)
        {
            if (extension != null)
            {
                XmlNode nodesetExtensionsNode = addXmlElement(_nodesetDoc, node, "Extensions");
                XmlNode nodesetExtensionNode = addXmlElement(_nodesetDoc, nodesetExtensionsNode, "Extension");
                XmlDocument extDoc = new XmlDocument();
                extDoc.LoadXml(extension);
                XmlNode nodesetExtensionContents = _nodesetDoc.ImportNode(extDoc.DocumentElement, true);
                nodesetExtensionNode.AppendChild(nodesetExtensionContents);
            }
        }

        XmlNode getXmlChlidByXmiId(XmlNode parent, string id)
        {
            XmlNode node = null;
            foreach (XmlNode child in parent.ChildNodes)
            {
                XmlAttribute idAttr = child.Attributes[String.Format("{0}:id", _xmiNSPräfix)];
                if (idAttr != null)
                {
                    if (idAttr.Value == id)
                    {
                        node = child;
                        break;
                    }
                }
            }
            return node;
        }

        void mergeXmiNodes(XmlNode xmi, XmlNode xmiX, string id)
        {
            if (xmiX == null)
            {
                try
                {
                    xmiX = _xmiXElementMap[id];
                }
                catch
                { }

                if (xmiX == null)
                {
                    return;
                } 
            }

            foreach (XmlAttribute attrX in xmiX.Attributes)
            {
                if (attrX.Name == String.Format("{0}:id", _xmiNSPräfix))
                {
                    continue;
                }

                XmlAttribute attr = xmi.Attributes[attrX.Name];
                if (attr == null)
                { // add attrX to child
                    addXmlAttribute(xmi.OwnerDocument, xmi, attrX.Name, attrX.Value);
                }
                else
                { // check if to change attribute
                    if (attrX.Value != attr.Value)
                    {
                        attr.Value = attrX.Value;
                    }
                }
            }

            foreach (XmlNode childX in xmiX.ChildNodes)
            {
                XmlAttribute idAttrX = childX.Attributes[String.Format("{0}:id", _xmiNSPräfix)];
                if (idAttrX != null)
                {
                    XmlNode child = getXmlChlidByXmiId(xmi, idAttrX.Value);
                    if (child != null)
                    { // merge nodes
                        mergeXmiNodes(child, childX, idAttrX.Value);
                    }
                    else
                    { // add node
                        XmlNode nodesetImportedNode = xmi.OwnerDocument.ImportNode(childX, true);
                        XmlNode clone = nodesetImportedNode.Clone();
                        xmi.AppendChild(clone);
                    }
                }
            }
        }

        XmlNode getXmiClass(string id)
        { 
            XmlNode xmi = null;

            try
            {
                xmi = _xmiClassMap[id];
            }
            catch
            { }

            if (xmi != null)
            {
                XmlAttribute noOTAttribute = xmi.Attributes["noObjectType"];
                if (noOTAttribute != null)
                {
                    xmi = null;
                }         
            }

            if (xmi != null)
            { 
                XmlNode xmiX = null;
                try
                {
                    xmiX = _xmiXElementMap[id];
                }
                catch
                { }

                if (xmiX != null)
                {
                    mergeXmiNodes(xmi, xmiX, id);
                }
            }
            return xmi;
        }

        XmlNode getXmiStubEA(string id)
        { 
            XmlNode xmi = null;
            try
            {
                xmi = _xmiStubEAMap[id];
            }
            catch
            { }
            return xmi;
        }

        XmlNode getXmiElementEA(string id)
        { 
            XmlNode xmi = null;
            try
            {
                xmi = _xmiElementMapEA[id];
            }
            catch
            { }
            return xmi;
        }

        XmlNode getXmiAttributeEA(string id)
        { 
            XmlNode xmi = null;
            try
            {
                xmi = _xmiAttributeMapEA[id];
            }
            catch
            { }
            return xmi;
        }

        XmlNode getXmiIgnoredOwnedAttributeEA(string id)
        { 
            XmlNode xmi = null;
            try
            {
                xmi = _xmiOwnedAttributeMapIgnore[id];
            }
            catch
            { }
            return xmi;
        }

        XmlNode getXmiPackage(string id)
        { 
            XmlNode xmi = null;
            try
            {
                xmi = _xmiPackageMap[id];
            }
            catch
            { }
            return xmi;
        }

        XmlNode getXmiStructure(string id)
        { 
            XmlNode xmi = null;
            try
            {
                xmi = _xmiStructureMap[id];
            }
            catch
            { }

            if (xmi != null)
            { 
                XmlNode xmiX = null;
                try
                {
                    xmiX = _xmiXElementMap[id];
                }
                catch
                { }

                if (xmiX != null)
                {
                    mergeXmiNodes(xmi, xmiX, id);
                }
            }
            return xmi;
        }

        XmlNode getXmiEnumeration(string id)
        { 
            XmlNode xmi = null;
            try
            {
                xmi = _xmiEnumerationMap[id];
            }
            catch
            { }

            if (xmi != null)
            { 
                XmlNode xmiX = null;
                try
                {
                    xmiX = _xmiXElementMap[id];
                }
                catch
                { }

                if (xmiX != null)
                {
                    mergeXmiNodes(xmi, xmiX, id);
                }
            }
            return xmi;
        }

        XmlNode getXmiClassByName(string name)
        {
            foreach (KeyValuePair<String, XmlNode> kv in _xmiClassMap)
            {
                XmlAttribute attrName = kv.Value.Attributes["name"];
                if (attrName != null)
                {
                    if (attrName.Value == name)
                    {
                        return getXmiClass(kv.Value.Attributes[String.Format("{0}:id", _xmiNSPräfix)].Value);
                    }
                }
            }
            return null;
        }

    }
}
