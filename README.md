# XMItoNodeset
XMItoNodeset is a command line tool to convert Enterprise Architect UML class diagrams to OPC UA nodeset files.

## Build
XMItoNodeset is written in C# using the .NET framework V4.5.2. <br>
Microsoft Visual Studio 2015 is used to build it. 

## Command line arguments
* /xmi < file name ><br>
  Name of the XMI file to convert.<br>
  UML classes are converted into OPC UA object types.
  <br>This argument could be used multiple times.
* /xmiDT < file name ><br>
  Name of the XMI file to convert.<br>
  UML classes are converted into OPC UA data types.
  <br>This argument could be used multiple times.
* /xmiS < file name ><br>
  Name of the XMI file to convert.<br>
  UML classes are converted according to the stereotype into OPC UA object types or OPC UA data types.
  <br>This argument could be used multiple times.
* /nodeset < file name ><br>
  Name of the generated nodeset file
* /nodesetUrl < URL string ><br>
  URL used for the generation of the nodeset
* /nodesetTypeDictionary < name ><br>
  Name of the type dictionary in the nodeset
  <br> [optional; default: "XMItoNodeset"]
* /nodesetImport < file name ><br>
  Name of the nodeset file to import
  <br> [optional]
* /nodesetStartId < id ><br>
  Start integer node id for the conversion 
  <br> [optional; default: 0]
* /nodeIdMap < file name ><br>
  Name of the node id mapping file
  <br>[optional; default: "NodeIdMap.txt"]
* /binaryTypes < file name ><br>
  Name of the binary types schema file 
  <br> [optional, default: "BinaryTypes.xml"]
* /xmlTypes < file name ><br>
  Name of the XML types schema file 
  <br> [optional, default: "XMLTypes.xml"]
* /generate < name ><br>
  Restrict the generation of the nodeset to a subset of the UML classes
  <br>[optional]
* /ignoreClassMember
  Name of the class member to ignore in the conversion
  <br>[optional]
* /word < file name >  
  Name of the MS Word file to generate
  <br>[optional]  

