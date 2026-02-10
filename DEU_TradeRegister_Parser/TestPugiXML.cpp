
#include "pugixml.hpp"

#include "TestPugiXML.h"


// DG 2025.12.30
void Test_CreateXMLFile()
{
  w_string wsXMLFileName = L"C:\\Work\\2025.12.30\\TestFile.xml";

  pugi::xml_document oXmlDocument;

  // Add node
  pugi::xml_node oTestNodeNode = oXmlDocument.append_child(L"TestNode");

  // Set attribute
  pugi::xml_attribute oCurrentDayAttribute = oTestNodeNode.append_attribute(L"CurrentDay");
  oCurrentDayAttribute.set_value(28);

  // Save XML to file
  oXmlDocument.save_file(W_StrToWStr(wsXMLFileName).c_str(),
    PUGIXML_TEXT("  "),
    pugi::format_default | pugi::format_write_bom,
    pugi::encoding_utf16_le);
}
// DG