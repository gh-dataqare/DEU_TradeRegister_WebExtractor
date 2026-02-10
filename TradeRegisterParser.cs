using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;

namespace DEU_TradeRegister_WebExtractor
{
  public class CompanyTradeRegisterData
  {
    // Registration data of company
    public string LocalCourtId { get; set; }
    public string LocalCourtName { get; set; }
    public string TypeOfRegister { get; set; }
    public string RegistrationNumber { get; set; }

    // Registration date of company
    public string RegistrationDate_FormattedDate { get; set; }
    public string RegistrationDate_UnformattedDate { get; set; }

    // Other data of company
    public string CompanyName { get; set; }
    public string LegalForm { get; set; }
    public string CityName_WhereMainOfficeIsLocated { get; set; }

    // Postal address of the main office
    public string PostalAddressType { get; set; }
    public string StreetName { get; set; }
    public string HouseNumber { get; set; }
    public string PostalCode { get; set; }
    public string CityName { get; set; }
    public string CountryName { get; set; }

    // Type of representation
    public string TypeOfRepresentation { get; set; }

    // Share capital (GmbH)
    public string ShareCapital_GmbH_TotalAmountOfMoney { get; set; }
    public string ShareCapital_GmbH_OutdatedCurrency { get; set; }
    public string ShareCapital_GmbH_ActualCurrency { get; set; }

    // Share capital (AG)
    public string ShareCapital_AG_TotalAmountOfMoney { get; set; }
    public string ShareCapital_AG_ActualCurrency { get; set; }

    // Branch of company
    public string BranchText { get; set; }

    public CompanyTradeRegisterData()
    {
      Clear();
    }

    public void Clear()
    {
      LocalCourtId = string.Empty;
      LocalCourtName = string.Empty;
      TypeOfRegister = string.Empty;
      RegistrationNumber = string.Empty;

      RegistrationDate_FormattedDate = string.Empty;
      RegistrationDate_UnformattedDate = string.Empty;

      CompanyName = string.Empty;
      LegalForm = string.Empty;
      CityName_WhereMainOfficeIsLocated = string.Empty;

      PostalAddressType = string.Empty;
      StreetName = string.Empty;
      HouseNumber = string.Empty;
      PostalCode = string.Empty;
      CityName = string.Empty;
      CountryName = string.Empty;

      TypeOfRepresentation = string.Empty;

      ShareCapital_GmbH_TotalAmountOfMoney = string.Empty;
      ShareCapital_GmbH_OutdatedCurrency = string.Empty;
      ShareCapital_GmbH_ActualCurrency = string.Empty;

      ShareCapital_AG_TotalAmountOfMoney = string.Empty;
      ShareCapital_AG_ActualCurrency = string.Empty;

      BranchText = string.Empty;
    }
  }

  public class TradeRegisterParser
  {
    public static CompanyTradeRegisterData ParseXMLFile(
      string xmlFilePath,
      APX.ClJobProtocol jobProtocol)
    {
      CompanyTradeRegisterData data = new CompanyTradeRegisterData();

      try
      {
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.Load(xmlFilePath);

        XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
        nsmgr.AddNamespace("tns", xmlDoc.DocumentElement.NamespaceURI);

        // Parse registration data
        data.LocalCourtId = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:grunddaten/tns:verfahrensdaten/tns:instanzdaten/tns:auswahl_instanzbehoerde/tns:gericht/code");

        data.TypeOfRegister = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:grunddaten/tns:verfahrensdaten/tns:instanzdaten/tns:aktenzeichen/tns:auswahl_aktenzeichen/tns:aktenzeichen.strukturiert/tns:register/code");

        data.RegistrationNumber = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:grunddaten/tns:verfahrensdaten/tns:instanzdaten/tns:aktenzeichen/tns:auswahl_aktenzeichen/tns:aktenzeichen.strukturiert/tns:laufendeNummer");

        // Parse registration dates
        data.RegistrationDate_FormattedDate = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:satzungsdatum/tns:aktuellesSatzungsdatum");

        data.RegistrationDate_UnformattedDate = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:satzungsdatum/tns:satzungsdatumFreitext");

        // Parse company data
        data.CompanyName = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:bezeichnung/tns:bezeichnung.aktuell");

        data.LegalForm = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:angabenZurRechtsform/tns:rechtsform/code");

        data.CityName_WhereMainOfficeIsLocated = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:sitz/tns:ort");

        // Parse postal address
        data.PostalAddressType = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:anschrift/tns:anschriftstyp/code");

        data.StreetName = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:anschrift/tns:strasse");

        data.HouseNumber = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:anschrift/tns:hausnummer");

        data.PostalCode = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:anschrift/tns:postleitzahl");

        data.CityName = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:anschrift/tns:ort");

        data.CountryName = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:anschrift/tns:staat/code");

        // Parse type of representation
        data.TypeOfRepresentation = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:vertretung/tns:allgemeineVertretungsregelung");

        // Parse share capital (GmbH)
        data.ShareCapital_GmbH_TotalAmountOfMoney = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:kapital/tns:nennkapital/tns:betrag/tns:zahl");

        data.ShareCapital_GmbH_OutdatedCurrency = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:kapital/tns:nennkapital/tns:betrag/tns:waehrung/code");

        data.ShareCapital_GmbH_ActualCurrency = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:kapital/tns:nennkapital/tns:betragAktualisiert/tns:waehrung/code");

        // Parse share capital (AG)
        data.ShareCapital_AG_TotalAmountOfMoney = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:kapital/tns:grundkapital/tns:betrag/tns:zahl");

        data.ShareCapital_AG_ActualCurrency = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:kapital/tns:grundkapital/tns:betrag/tns:waehrung/code");

        // Parse branch text
        data.BranchText = GetXPathValue(xmlDoc, nsmgr,
          "/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:gegenstand");
      }
      catch (Exception ex)
      {
        jobProtocol.WriteLine("Error parsing XML file '" + xmlFilePath + "': " + ex.Message);
        throw;
      }

      return data;
    }

    private static string GetXPathValue(XmlDocument doc, XmlNamespaceManager nsmgr, string xpath)
    {
      try
      {
        XmlNode node = doc.SelectSingleNode(xpath, nsmgr);
        if (node?.InnerText != null)
        {
          // Remove all newline and carriage return characters
          string value = node.InnerText.Trim();
          value = value.Replace("\r\n", " ");
          value = value.Replace("\n", " ");
          value = value.Replace("\r", " ");
          // Also collapse multiple spaces into single space
          while (value.Contains("  "))
          {
            value = value.Replace("  ", " ");
          }
          return value;
        }
        return string.Empty;
      }
      catch
      {
        return string.Empty;
      }
    }
  }
}
