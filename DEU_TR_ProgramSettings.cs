using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace DEU_TradeRegister_WebExtractor
{
  internal class DEU_TR_ProgramSettings
  {
    static public String GetSettingsXmlFileName()
    {
      String sExe_FileName = "";
      String sSettingsXml_FileName = "";

      Process oCurrentProcess = Process.GetCurrentProcess();

      sExe_FileName = oCurrentProcess.MainModule.FileName;

      String sBaseFileName = sExe_FileName.Remove(sExe_FileName.Length - 4);

      sSettingsXml_FileName = sBaseFileName + "_Settings.xml";

      return sSettingsXml_FileName;
    }

    static String ReadAttributeValue_by_XPath(
      String v_sXmlNode_XPath,
      String v_sXmlAttribute_Name)
    {
      String sResult = "";


      String sSettingsXml_FileName = GetSettingsXmlFileName();

      XmlDocument oSettingsXmlFile = new XmlDocument();

      oSettingsXmlFile.Load(sSettingsXml_FileName);


      XmlNode oXmlNode = oSettingsXmlFile.SelectSingleNode(v_sXmlNode_XPath);

      if (oXmlNode != null)
      {
        XmlAttribute oXmlAttribute = oXmlNode.Attributes[v_sXmlAttribute_Name];

        if (oXmlAttribute != null)
        {
          sResult = oXmlAttribute.Value;
        }
        else
        {
          throw new Exception(
            "The following attribute does not exist in the following node in Settings file: " +
            "\"" + v_sXmlAttribute_Name + "\" in " +
            "\"" + v_sXmlNode_XPath + "\". " +
            "(\"" + sSettingsXml_FileName + "\")");
        }
      }
      else
      {
        throw new Exception(
          "The following node does not exist in Settings file: " +
          "\"" + v_sXmlNode_XPath + "\". " +
          "(\"" + sSettingsXml_FileName + "\")");
      }

      oSettingsXmlFile = null;


      return sResult;
    }

    static void WriteAttributeValue_by_XPath(
      String v_sXmlNode_XPath,
      String v_sXmlAttribute_Name,
      String v_sXmlAttribute_Value)
    {
      String sSettingsXml_FileName = GetSettingsXmlFileName();

      XmlDocument oSettingsXmlFile = new XmlDocument();

      oSettingsXmlFile.Load(sSettingsXml_FileName);


      XmlNode oXmlNode = oSettingsXmlFile.SelectSingleNode(v_sXmlNode_XPath);

      if (oXmlNode != null)
      {
        XmlAttribute oXmlAttribute = oXmlNode.Attributes[v_sXmlAttribute_Name];

        if (oXmlAttribute != null)
        {
          oXmlAttribute.Value = v_sXmlAttribute_Value;
        }
        else
        {
          throw new Exception(
            "The following attribute does not exist in the following node in Settings file: " +
            "\"" + v_sXmlAttribute_Name + "\" in " +
            "\"" + v_sXmlNode_XPath + "\". " +
            "(\"" + sSettingsXml_FileName + "\")");
        }
      }
      else
      {
        throw new Exception(
          "The following node does not exist in Settings file: " +
          "\"" + v_sXmlNode_XPath + "\". " +
          "(\"" + sSettingsXml_FileName + "\")");
      }

      oSettingsXmlFile.Save(sSettingsXml_FileName);

      oSettingsXmlFile = null;
    }

    public static List<String> ReadProxyList()
    {
      List<String> proxyList = new List<String>();
      
      String sSettingsXml_FileName = GetSettingsXmlFileName();
      XmlDocument oSettingsXmlFile = new XmlDocument();
      oSettingsXmlFile.Load(sSettingsXml_FileName);
      
      XmlNodeList proxyNodes = oSettingsXmlFile.SelectNodes("/Settings/ProxyList/Proxy");
      
      if (proxyNodes != null)
      {
        foreach (XmlNode proxyNode in proxyNodes)
        {
          XmlAttribute valueAttr = proxyNode.Attributes["value"];
          if (valueAttr != null && !String.IsNullOrWhiteSpace(valueAttr.Value))
          {
            proxyList.Add(valueAttr.Value.Trim());
          }
        }
      }
      
      return proxyList;
    }

    public static void Read_CommonSettings_Settings(
      out String v_sTradeRegisterSeachPageUrl,
      out String v_sLogFolder,
      out String v_sOutputFolder,
      out String v_sBrowserDownloadFolder,
      out String v_sOutputCSVFile,
      out String v_sBrowserRestartInterval,
      out String v_sProcessingMode)
    {
      v_sTradeRegisterSeachPageUrl = ReadAttributeValue_by_XPath(
        "/Settings/CommonSettings/Option[@name='TradeRegisterSeachPageUrl']",
        "value");

      v_sLogFolder = ReadAttributeValue_by_XPath(
        "/Settings/CommonSettings/Option[@name='LogFolder']",
        "value");

      v_sOutputFolder = ReadAttributeValue_by_XPath(
        "/Settings/CommonSettings/Option[@name='OutputFolder']",
        "value");

      v_sBrowserDownloadFolder = ReadAttributeValue_by_XPath(
        "/Settings/CommonSettings/Option[@name='BrowserDownloadFolder']",
        "value");

      v_sOutputCSVFile = ReadAttributeValue_by_XPath(
        "/Settings/CommonSettings/Option[@name='OutputCSVFile']",
        "value");

      v_sBrowserRestartInterval = ReadAttributeValue_by_XPath(
        "/Settings/CommonSettings/Option[@name='BrowserRestartInterval']",
        "value");

      v_sProcessingMode = ReadAttributeValue_by_XPath(
        "/Settings/CommonSettings/Option[@name='ProcessingMode']",
        "value");
    }

    public static void Read_RangeMode_Settings(
      out String v_sStartRegisterNumber,
      out String v_sEndRegisterNumber)
    {
      v_sStartRegisterNumber = ReadAttributeValue_by_XPath(
        "/Settings/RangeMode_Settings/Option[@name='StartRegisterNumber']",
        "value");

      v_sEndRegisterNumber = ReadAttributeValue_by_XPath(
        "/Settings/RangeMode_Settings/Option[@name='EndRegisterNumber']",
        "value");
    }

    public static void Update_StartRegisterNumber_InSettingsFile(
      String v_sStartRegisterNumber)
    {
      WriteAttributeValue_by_XPath(
        "/Settings/RangeMode_Settings/Option[@name='StartRegisterNumber']",
        "value",
        v_sStartRegisterNumber);
    }

    public static void InsertOrUpdate_UnsuccessfullyDownloadedRegisterNumber__in__ListMode_Settings(
      String v_sTradeRegisterNumber,
      Boolean v_bSuccessfullyDownloaded,
      String v_sErrorOfAType_MessageText,
      int v_iUnsuccessfulLinkCount,
      int v_iNumberOfLinksForDownload) 
    {
      String sSettingsXml_FileName = GetSettingsXmlFileName();

      XmlDocument oSettingsXmlFile = new XmlDocument();

      oSettingsXmlFile.Load(sSettingsXml_FileName);


      String sXmlNode_XPath = "/Settings/ListMode_Settings/RegisterNumberList/Option" +
        "[@name='RegisterNumber' and @value='" + v_sTradeRegisterNumber + "']";

      XmlNode oXmlNode = oSettingsXmlFile.SelectSingleNode(sXmlNode_XPath);

      if (oXmlNode == null)
      {
        String sRegisterNumberList_XmlNode_XPath = "/Settings/ListMode_Settings/RegisterNumberList";

        XmlNode oRegisterNumberList_XmlNode = oSettingsXmlFile.SelectSingleNode(
          sRegisterNumberList_XmlNode_XPath);

        if(oRegisterNumberList_XmlNode != null )
        {
          XmlNode oOption_XmlNode = oSettingsXmlFile.CreateNode(XmlNodeType.Element, "Option", "");

          oRegisterNumberList_XmlNode.AppendChild(oOption_XmlNode);


          XmlAttribute oName_XmlAttribute = oSettingsXmlFile.CreateAttribute("name");
          oName_XmlAttribute.Value = "RegisterNumber";
          oOption_XmlNode.Attributes.Append(oName_XmlAttribute);

          XmlAttribute oValue_XmlAttribute = oSettingsXmlFile.CreateAttribute("value");
          oValue_XmlAttribute.Value = "";
          oOption_XmlNode.Attributes.Append(oValue_XmlAttribute);

          XmlAttribute oDownloaded_XmlAttribute = oSettingsXmlFile.CreateAttribute("downloaded");
          oDownloaded_XmlAttribute.Value = "False";
          oOption_XmlNode.Attributes.Append(oDownloaded_XmlAttribute);

          XmlAttribute oErrorMessage_XmlAttribute = oSettingsXmlFile.CreateAttribute("error_message");
          oErrorMessage_XmlAttribute.Value = "";
          oOption_XmlNode.Attributes.Append(oErrorMessage_XmlAttribute);

          XmlAttribute oUnsuccessfulLinkCount_XmlAttribute = oSettingsXmlFile.CreateAttribute("unsuccessful_link_count");
          oUnsuccessfulLinkCount_XmlAttribute.Value = "0";
          oOption_XmlNode.Attributes.Append(oUnsuccessfulLinkCount_XmlAttribute);

          XmlAttribute oNumberOfLinksForDownload_XmlAttribute = oSettingsXmlFile.CreateAttribute("number_of_links_for_download");
          oNumberOfLinksForDownload_XmlAttribute.Value = "0";
          oOption_XmlNode.Attributes.Append(oNumberOfLinksForDownload_XmlAttribute);

          oXmlNode = oOption_XmlNode;
        }
        else
        {
          throw new Exception(
            "The following node does not exist in Settings file: " +
            "\"" + sRegisterNumberList_XmlNode_XPath + "\". " +
            "(\"" + sSettingsXml_FileName + "\")");
        }
      }
      


      if (oXmlNode != null)
      {
        // "value"
        XmlAttribute oValue_XmlAttribute = oXmlNode.Attributes["value"];

        if (oValue_XmlAttribute != null)
        {
          oValue_XmlAttribute.Value = v_sTradeRegisterNumber;
        }
        else
        {
          throw new Exception(
            "The following attribute does not exist in the following node in Settings file: " +
            "\"value\" in " +
            "\"" + sXmlNode_XPath + "\". " +
            "(\"" + sSettingsXml_FileName + "\")");
        }

        // "downloaded"
        XmlAttribute oDownloaded_XmlAttribute = oXmlNode.Attributes["downloaded"];

        if (oDownloaded_XmlAttribute != null)
        {
          oDownloaded_XmlAttribute.Value = v_bSuccessfullyDownloaded.ToString();
        }
        else
        {
          throw new Exception(
            "The following attribute does not exist in the following node in Settings file: " +
            "\"downloaded\" in " +
            "\"" + sXmlNode_XPath + "\". " +
            "(\"" + sSettingsXml_FileName + "\")");
        }

        // "error_message"
        XmlAttribute oErrorMessage_XmlAttribute = oXmlNode.Attributes["error_message"];

        if (oErrorMessage_XmlAttribute != null)
        {
          oErrorMessage_XmlAttribute.Value = v_sErrorOfAType_MessageText;
        }
        else
        {
          throw new Exception(
            "The following attribute does not exist in the following node in Settings file: " +
            "\"error_message\" in " +
            "\"" + sXmlNode_XPath + "\". " +
            "(\"" + sSettingsXml_FileName + "\")");
        }

        // "unsuccessful_link_count"
        XmlAttribute oUnsuccessfulLinkCount_XmlAttribute = oXmlNode.Attributes["unsuccessful_link_count"];

        if (oUnsuccessfulLinkCount_XmlAttribute != null)
        {
          oUnsuccessfulLinkCount_XmlAttribute.Value = v_iUnsuccessfulLinkCount.ToString();
        }
        else
        {
          throw new Exception(
            "The following attribute does not exist in the following node in Settings file: " +
            "\"unsuccessful_link_count\" in " +
            "\"" + sXmlNode_XPath + "\". " +
            "(\"" + sSettingsXml_FileName + "\")");
        }

        // "number_of_links_for_download"
        XmlAttribute oNumberOfLinksForDownload_XmlAttribute = oXmlNode.Attributes["number_of_links_for_download"];

        if (oNumberOfLinksForDownload_XmlAttribute != null)
        {
          oNumberOfLinksForDownload_XmlAttribute.Value = v_iNumberOfLinksForDownload.ToString();
        }
        else
        {
          throw new Exception(
            "The following attribute does not exist in the following node in Settings file: " +
            "\"number_of_links_for_download\" in " +
            "\"" + sXmlNode_XPath + "\". " +
            "(\"" + sSettingsXml_FileName + "\")");
        }
      }
      else
      {
        throw new Exception(
          "The following node does not exist in Settings file: " +
          "\"" + sXmlNode_XPath + "\". " +
          "(\"" + sSettingsXml_FileName + "\")");
      }

      oSettingsXmlFile.Save(sSettingsXml_FileName);

      oSettingsXmlFile = null;


      Update_ListMode_Statistics();
    }

    public static void Read_ListMode_Settings(
      out List<String> v_oRegisterNumberList)
    {
      v_oRegisterNumberList = new List<String>();


      String sSettingsXml_FileName = GetSettingsXmlFileName();

      XmlDocument oSettingsXmlFile = new XmlDocument();

      oSettingsXmlFile.Load(sSettingsXml_FileName);


      String sXmlNodes_XPath = "/Settings/ListMode_Settings/RegisterNumberList/Option" +
        "[@name='RegisterNumber' and @downloaded='False']";

      XmlNodeList oXmlNodeList = oSettingsXmlFile.SelectNodes(sXmlNodes_XPath);

      if (oXmlNodeList != null)
      {
        int iNodeId = 1;
        int iNodeCount = oXmlNodeList.Count;


        foreach (XmlNode oXmlNode in oXmlNodeList)
        {
          XmlAttribute oValue_XmlAttribute = oXmlNode.Attributes["value"];

          if (oValue_XmlAttribute != null)
          {
            String sRegisterNumber = oValue_XmlAttribute.Value;

            if (!String.IsNullOrEmpty(sRegisterNumber))
            {
              v_oRegisterNumberList.Add(sRegisterNumber);
            }
            else 
            {
              throw new Exception(
                "The following attribute is empty in the following node in Settings file: " +
                "\"value\" in node with id = " +
                iNodeId.ToString() + " (of " + iNodeCount.ToString() + ") " +
                "\"" + sXmlNodes_XPath + "\". " +
                "(\"" + sSettingsXml_FileName + "\")");
            }
          }
          else
          {
            throw new Exception(
              "The following attribute does not exist in the following node in Settings file: " +
              "\"value\" in node with id = " + 
              iNodeId.ToString() + " (of " + iNodeCount.ToString() + ") " +
              "\"" + sXmlNodes_XPath + "\". " +
              "(\"" + sSettingsXml_FileName + "\")");
          }

          iNodeId++;
        }
      }

      oSettingsXmlFile = null;
    }


    public static void Update_ListMode_Statistics()
    {
      int iUnsuccessfulLinkTotalCount = 0;

      String sSettingsXml_FileName = GetSettingsXmlFileName();

      XmlDocument oSettingsXmlFile = new XmlDocument();

      oSettingsXmlFile.Load(sSettingsXml_FileName);


      // Gather Statistics
      String sXmlNodes_XPath = "/Settings/ListMode_Settings/RegisterNumberList/Option" +
        "[@name='RegisterNumber' and @downloaded='False']";

      XmlNodeList oXmlNodeList = oSettingsXmlFile.SelectNodes(sXmlNodes_XPath);

      if (oXmlNodeList != null)
      {
        int iNodeId = 1;
        int iNodeCount = oXmlNodeList.Count;


        foreach (XmlNode oXmlNode in oXmlNodeList)
        {
          XmlAttribute oUnsuccessfulLinkCount_XmlAttribute = oXmlNode.Attributes["unsuccessful_link_count"];

          if (oUnsuccessfulLinkCount_XmlAttribute != null)
          {
            String sUnsuccessfulLinkCount = oUnsuccessfulLinkCount_XmlAttribute.Value;

            if (!String.IsNullOrEmpty(sUnsuccessfulLinkCount))
            {
              int iUnsuccessfulLinkCount = 0;

              if (Int32.TryParse(sUnsuccessfulLinkCount, out iUnsuccessfulLinkCount))
              {
                iUnsuccessfulLinkTotalCount += iUnsuccessfulLinkCount;
              }
              else
              {
                throw new Exception(
                  "The following attribute cannot be converted to Int32 number " +
                  "in the following node in Settings file: " +
                  "\"unsuccessful_link_count\" (value==\"" + sUnsuccessfulLinkCount + "\") " +
                  "in node with id = " +
                  iNodeId.ToString() + " (of " + iNodeCount.ToString() + ") " +
                  "\"" + sXmlNodes_XPath + "\". " +
                  "(\"" + sSettingsXml_FileName + "\")");
              }
            }
            else
            {
              throw new Exception(
                "The following attribute is empty in the following node in Settings file: " +
                "\"unsuccessful_link_count\" in node with id = " +
                iNodeId.ToString() + " (of " + iNodeCount.ToString() + ") " +
                "\"" + sXmlNodes_XPath + "\". " +
                "(\"" + sSettingsXml_FileName + "\")");
            }
          }

          iNodeId++;
        }
      }


      // Update Statistics
      String sSettings_XmlNode_XPath = "/Settings";
      String sListMode_Statistics_XmlNode_XPath = "/Settings/ListMode_Statistics";
      String sElement_XmlNode_XPath = "/Settings/ListMode_Statistics/Element";

      XmlNode oSettings_XmlNode = oSettingsXmlFile.SelectSingleNode(sSettings_XmlNode_XPath);
      if (oSettings_XmlNode == null)
      {
        throw new Exception(
          "The following node does not exist in Settings file: " +
          "\"" + sSettings_XmlNode_XPath + "\". " +
          "(\"" + sSettingsXml_FileName + "\")");
      }

      XmlNode oListMode_Statistics = oSettingsXmlFile.SelectSingleNode(sListMode_Statistics_XmlNode_XPath);
      if(oListMode_Statistics == null)
      {
        oListMode_Statistics = oSettingsXmlFile.CreateNode(XmlNodeType.Element, "ListMode_Statistics", "");

        oSettings_XmlNode.AppendChild(oListMode_Statistics);
      }

      XmlNode oElement_XmlNode = oSettingsXmlFile.SelectSingleNode(sElement_XmlNode_XPath);
      if (oElement_XmlNode == null)
      {
        oElement_XmlNode = oSettingsXmlFile.CreateNode(XmlNodeType.Element, "Element", "");

        oListMode_Statistics.AppendChild(oElement_XmlNode);


        XmlAttribute oUnsuccessfulLinkTotalCount_XmlAttribute = oSettingsXmlFile.CreateAttribute("unsuccessful_link_total_count");
        oUnsuccessfulLinkTotalCount_XmlAttribute.Value = "0";
        oElement_XmlNode.Attributes.Append(oUnsuccessfulLinkTotalCount_XmlAttribute);
      }


      {
        String sUnsuccessfulLinkTotalCount_XmlAttribute_Name = "unsuccessful_link_total_count";

        XmlAttribute oUnsuccessfulLinkTotalCount_XmlAttribute = 
          oElement_XmlNode.Attributes[sUnsuccessfulLinkTotalCount_XmlAttribute_Name];

        if (oUnsuccessfulLinkTotalCount_XmlAttribute != null)
        {
          oUnsuccessfulLinkTotalCount_XmlAttribute.Value = iUnsuccessfulLinkTotalCount.ToString();
        }
        else
        {
          throw new Exception(
            "The following attribute does not exist in the following node in Settings file: " +
            "\"" + sUnsuccessfulLinkTotalCount_XmlAttribute_Name + "\" in " +
            "\"" + sElement_XmlNode_XPath + "\". " +
            "(\"" + sSettingsXml_FileName + "\")");
        }
      }


      oSettingsXmlFile.Save(sSettingsXml_FileName);

      oSettingsXmlFile = null;
    }

  }
}
