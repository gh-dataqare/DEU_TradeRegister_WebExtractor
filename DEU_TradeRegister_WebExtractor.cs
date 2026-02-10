using APX;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Xml;

namespace DEU_TradeRegister_WebExtractor
{
  public class DEU_TradeRegister_WebExtractor
  {
    protected String m_sTradeRegisterSeachPageUrl = "";
    protected String m_sLogFolder = "";
    protected String m_sOutputFolder = "";
    protected String m_sBrowserDownloadFolder = "";
    protected String m_sOutputCSVFile = "";
    protected int m_iBrowserRestartInterval = 5;
    protected StreamingCSVWriter m_csvWriter = null;
    protected int m_iNetworkErrorCounter = 0;
    protected Random m_random = new Random();
    protected List<String> m_proxyList = new List<String>();
    protected int m_currentProxyIndex = 0;

    // Array of user agents to rotate through
    protected static readonly string[] UserAgents = new string[]
    {
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36",
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36",
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36",
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0",
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:122.0) Gecko/20100101 Firefox/122.0",
      "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36",
      "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36"
    };

    /*
    protected String m_sFolderForDownloadViaInterceptionOfHttpResponce = "";
    protected String m_sFileDownloadedViaResponse_FileName = "";
    protected Boolean m_bTradeRegisterDataFile_was_SuccessfullyDownloaded__via_HttpResponse = false;
    protected Object m_oFileDownloadedViaResponse_CS = new Object();
    */

    protected Boolean m_bLogHttpData = false;
    protected String m_sHttpDataFolder = "";


    protected int m_iWarningOfType1Count = 0;


    public DEU_TradeRegister_WebExtractor(
      String v_sTradeRegisterSeachPageUrl,
      String v_sLogFolder,
      String v_sOutputFolder,
      String v_sBrowserDownloadFolder,
      String v_sOutputCSVFile,
      int v_iBrowserRestartInterval,
      List<String> v_proxyList)
    {
      m_sTradeRegisterSeachPageUrl = v_sTradeRegisterSeachPageUrl;
      m_sLogFolder = v_sLogFolder;
      m_sOutputFolder = v_sOutputFolder;
      m_sBrowserDownloadFolder = v_sBrowserDownloadFolder;
      m_sOutputCSVFile = v_sOutputCSVFile;
      m_iBrowserRestartInterval = v_iBrowserRestartInterval;
      m_iNetworkErrorCounter = 0;
      m_proxyList = v_proxyList ?? new List<String>();
      m_currentProxyIndex = 0;

      
      /*
      m_sFolderForDownloadViaInterceptionOfHttpResponce = Path.Combine(m_sOutputFolder, "ResponseFiles");

      if (Directory.Exists(m_sFolderForDownloadViaInterceptionOfHttpResponce))
        Directory.Delete(m_sFolderForDownloadViaInterceptionOfHttpResponce);

      Directory.CreateDirectory(m_sFolderForDownloadViaInterceptionOfHttpResponce);
      */


      // Please, set this member to 'true'
      //   in order to enabale logging of HTTP Response data
      //   into Log folder.
      m_bLogHttpData = true;
      m_sHttpDataFolder = "";

      m_iWarningOfType1Count = 0;
    }

    void ParseAndWriteXMLToCSV(
      String v_sXMLFilePath,
      ClJobProtocol v_oJobProtocol,
      String v_sPrefixStringForProtocol)
    {
      try
      {
        // Initialize CSV writer if not already created
        if (m_csvWriter == null && !string.IsNullOrEmpty(m_sOutputCSVFile))
        {
          m_csvWriter = new StreamingCSVWriter(m_sOutputCSVFile, ';', '"');
          m_csvWriter.WriteHeader(
            "LocalCourtId",
            "LocalCourtName",
            "TypeOfRegister",
            "RegistrationNumber",
            "RegistrationDate_FormattedDate",
            "RegistrationDate_UnformattedDate",
            "CompanyName",
            "LegalForm",
            "CityName_WhereMainOfficeIsLocated",
            "PostalAddressType",
            "StreetName",
            "HouseNumber",
            "PostalCode",
            "CityName",
            "CountryName",
            "TypeOfRepresentation",
            "ShareCapital_GmbH_TotalAmountOfMoney",
            "ShareCapital_GmbH_OutdatedCurrency",
            "ShareCapital_GmbH_ActualCurrency",
            "ShareCapital_AG_TotalAmountOfMoney",
            "ShareCapital_AG_ActualCurrency",
            "BranchText"
          );
          v_oJobProtocol.WriteLine("CSV output file initialized: " + m_sOutputCSVFile);
        }

        if (m_csvWriter != null)
        {
          // Parse the XML file
          CompanyTradeRegisterData data = TradeRegisterParser.ParseXMLFile(v_sXMLFilePath, v_oJobProtocol);
          
          // Write to CSV immediately (streaming mode)
          m_csvWriter.WriteCompanyData(data);
          
          v_oJobProtocol.WriteLine(v_sPrefixStringForProtocol + "Data parsed and written to CSV.");
        }
      }
      catch (Exception ex)
      {
        v_oJobProtocol.WriteLine(v_sPrefixStringForProtocol + "Error parsing XML: " + ex.Message);
      }
    }

    void ClearStatistics_before_SearchAtDEUTradeRegisterWebSite()
    {
      m_iWarningOfType1Count = 0;
    }

    void PrintStatistics(
      ClJobProtocol v_oJobProtocol,
      String v_sPrefixStringForProtocol)
    {
      if (m_iWarningOfType1Count > 0)
      {
        v_oJobProtocol.WriteLine(
          v_sPrefixStringForProtocol +
          "Warning of type '1' count: " + m_iWarningOfType1Count.ToString());
      }
    }

    private void StartMonitoringOfHttpResponces(
      WebDriver v_oBrowserWrapper)
    {
      if (m_bLogHttpData)
      {
        DateTime oCurrentTime = DateTime.Now;

        String sCurrentTimeStr =
          oCurrentTime.ToString("yyyy-MM-dd HH-mm-ss") +
          "_" + oCurrentTime.Ticks.ToString();

        String sFolderName = "HttpData_" + sCurrentTimeStr;

        m_sHttpDataFolder = Path.Combine(m_sLogFolder, sFolderName);

        if (Directory.Exists(m_sHttpDataFolder))
          Directory.Delete(m_sHttpDataFolder);

        Directory.CreateDirectory(m_sHttpDataFolder);
      }

      v_oBrowserWrapper.Manage().Network.NetworkRequestSent += Network_NetworkRequestSent;
      v_oBrowserWrapper.Manage().Network.NetworkResponseReceived += Network_NetworkResponseReceived;
      v_oBrowserWrapper.Manage().Network.StartMonitoring();
    }

    private void StopMonitoringOfHttpResponces(
      WebDriver v_oBrowserWrapper)
    {
      v_oBrowserWrapper.Manage().Network.StopMonitoring();
    }

    private void DumpRequestDataToLogFolder(NetworkRequestSentEventArgs e)
    {
      Console.Out.WriteLine();
      Console.Out.WriteLine(".........");
      Console.Out.WriteLine("I am inside Network_NetworkRequestSent().");


      try
      {
        // Create a new folder for files with data of current Request.
        DateTime oCurrentTime = DateTime.Now;
        String sCurrentTimeStr = oCurrentTime.Ticks.ToString();

        String sCurrentRequestFolderName = Path.Combine(m_sHttpDataFolder,
          sCurrentTimeStr);

        sCurrentRequestFolderName = sCurrentRequestFolderName + "_Request";


        if (Directory.Exists(sCurrentRequestFolderName))
          Directory.Delete(sCurrentRequestFolderName);

        Directory.CreateDirectory(sCurrentRequestFolderName);


        // Creatre "RequestDescription.txt"
        String sRequestHeaders = "";
        List<KeyValuePair<string, string>> oKeyValuePairList = e.RequestHeaders.ToList();

        for (int i = 0; i < oKeyValuePairList.Count; i++)
        {
          String sKey = oKeyValuePairList[i].Key;
          String sValue = oKeyValuePairList[i].Value;

          sRequestHeaders += "Key = \"" + sKey + "\"" + Environment.NewLine;
          sRequestHeaders += "Value = \"" + sValue + "\"" + Environment.NewLine;
          sRequestHeaders += Environment.NewLine;
        }

        String sRequestDescriptionStr =
          "RequestId = \"" + e.RequestId + "\"" + Environment.NewLine +
          "RequestMethod = \"" + e.RequestMethod + "\"" + Environment.NewLine +
          "RequestPostData = " + e.RequestPostData + Environment.NewLine +
          "ResponseUrl = \"" + e.RequestUrl + "\"" + Environment.NewLine +
           Environment.NewLine;

        sRequestDescriptionStr +=
          "RequestHeaders:" + Environment.NewLine +
          sRequestHeaders;


        String sRequestDescriptionFileName = "RequestDescription.txt";
        String sRequestDescriptionFullFileName = Path.Combine(sCurrentRequestFolderName, sRequestDescriptionFileName);

        Console.Out.WriteLine("Saving file: \"" + sRequestDescriptionFullFileName + "\".");

        File.WriteAllText(sRequestDescriptionFullFileName, sRequestDescriptionStr);
      }
      catch (Exception ex)
      {
        String sLogString =
          "Exception. Class name: " +
          ex.GetType().FullName + ". Message: \"" +
          ex.Message + "\"" + Environment.NewLine +
          "Stack trace:" + Environment.NewLine +
          ex.StackTrace;

        Console.Out.WriteLine(sLogString);
        Console.Out.WriteLine();
      }


      Console.Out.WriteLine(".........");
      Console.Out.WriteLine();
    }

    private void DumpResponseDataToLogFolder(NetworkResponseReceivedEventArgs e)
    {
      Console.Out.WriteLine();
      Console.Out.WriteLine(",,,,,,,,,");
      Console.Out.WriteLine("I am inside Network_NetworkResponseReceived().");


      try
      {
        // Create a new folder for files with data of current Response.
        DateTime oCurrentTime = DateTime.Now;
        String sCurrentTimeStr = oCurrentTime.Ticks.ToString();

        String sCurrentResponseFolderName = Path.Combine(m_sHttpDataFolder,
          sCurrentTimeStr);

        sCurrentResponseFolderName = sCurrentResponseFolderName + "_Response";

        if (e.ResponseResourceType.IndexOfAny(Path.GetInvalidFileNameChars()) == -1 &&
            e.ResponseResourceType.IndexOfAny(Path.GetInvalidPathChars()) == -1)
        {
          sCurrentResponseFolderName = sCurrentResponseFolderName +
            "_" + e.ResponseResourceType;
        }


        if (Directory.Exists(sCurrentResponseFolderName))
          Directory.Delete(sCurrentResponseFolderName);

        Directory.CreateDirectory(sCurrentResponseFolderName);


        // Creatre "ResponseDescription.txt"
        String sResponseHeaders = "";
        String sFileForDownload_FileName = "";
        List<KeyValuePair<string, string>> oKeyValuePairList = e.ResponseHeaders.ToList();

        for (int i = 0; i < oKeyValuePairList.Count; i++)
        {
          String sKey = oKeyValuePairList[i].Key;
          String sValue = oKeyValuePairList[i].Value;

          sResponseHeaders += "Key = \"" + sKey + "\"" + Environment.NewLine;
          sResponseHeaders += "Value = \"" + sValue + "\"" + Environment.NewLine;
          sResponseHeaders += Environment.NewLine;

          // Extract the name of a file,
          //   that ws selected by user of browser
          //   for downloading.
          if (sKey == "Content-Disposition")
          {
            String sPatternStr = "attachment;filename=";

            int iFileNamePos = sValue.IndexOf(sPatternStr);
            if (iFileNamePos == 0)
            {
              String sFileNameX = sValue.Substring(sPatternStr.Length);

              if (sFileNameX.Length > 2)
              {
                if (sFileNameX[0] == '\"' &&
                  sFileNameX[sFileNameX.Length - 1] == '\"')
                {
                  sFileNameX = sFileNameX.Substring(1, sFileNameX.Length - 2);

                  sFileForDownload_FileName = sFileNameX;
                }
              }
            }
          }
        }

        String sResponseDescriptionStr =
          "RequestId = \"" + e.RequestId + "\"" + Environment.NewLine +
          "ResponseResourceType = \"" + e.ResponseResourceType + "\"" + Environment.NewLine +
          "ResponseStatusCode = " + e.ResponseStatusCode + Environment.NewLine +
          "ResponseUrl = \"" + e.ResponseUrl + "\"" + Environment.NewLine +
           Environment.NewLine;


        if (e.ResponseResourceType == "Document")
        {
          sResponseDescriptionStr +=
            "FileForDownload_FileName = \"" + sFileForDownload_FileName + "\"" + Environment.NewLine +
            Environment.NewLine;
        }

        sResponseDescriptionStr +=
          "ResponseHeaders:" + Environment.NewLine +
          sResponseHeaders;


        String sResponseDescriptionFileName = "ResponseDescription.txt";
        String sResponseDescriptionFullFileName = Path.Combine(sCurrentResponseFolderName, sResponseDescriptionFileName);

        Console.Out.WriteLine("Saving file: \"" + sResponseDescriptionFullFileName + "\".");

        File.WriteAllText(sResponseDescriptionFullFileName, sResponseDescriptionStr);


        // Creatre "ResponseContent.bin"
        String sResponseContentFileName = "ResponseContent.bin";
        String sResponseContentFullFileName = Path.Combine(sCurrentResponseFolderName, sResponseContentFileName);

        Console.Out.WriteLine("Saving file: \"" + sResponseContentFullFileName + "\".");

        File.WriteAllBytes(sResponseContentFullFileName, e.ResponseContent.ReadAsByteArray());

      }
      catch (Exception ex)
      {
        String sLogString =
          "Exception. Class name: " +
          ex.GetType().FullName + ". Message: \"" +
          ex.Message + "\"" + Environment.NewLine +
          "Stack trace:" + Environment.NewLine +
          ex.StackTrace;

        Console.Out.WriteLine(sLogString);
        Console.Out.WriteLine();
      }


      Console.Out.WriteLine(",,,,,,,,,");
      Console.Out.WriteLine();
    }

    private void Network_NetworkRequestSent(object sender, NetworkRequestSentEventArgs e)
    {
      if (m_bLogHttpData)
        DumpRequestDataToLogFolder(e);
    }

    private void Network_NetworkResponseReceived(object sender, NetworkResponseReceivedEventArgs e)
    {
      if (m_bLogHttpData)
        DumpResponseDataToLogFolder(e);

      /*
      try
      { 
        // Retrieve the name of file
        //   that was selected by user of Web Browser
        //   for downloading.
        String sFileForDownload_FileName = "";
        List<KeyValuePair<string, string>> oKeyValuePairList = e.ResponseHeaders.ToList();

        for (int i = 0; i < oKeyValuePairList.Count; i++)
        {
          String sKey = oKeyValuePairList[i].Key;
          String sValue = oKeyValuePairList[i].Value;

          if (sKey == "Content-Disposition")
          {
            String sPatternStr = "attachment;filename=";

            int iFileNamePos = sValue.IndexOf(sPatternStr);
            if (iFileNamePos == 0)
            {
              String sFileNameX = sValue.Substring(sPatternStr.Length);

              if (sFileNameX.Length > 2)
              {
                if (sFileNameX[0] == '\"' &&
                  sFileNameX[sFileNameX.Length - 1] == '\"')
                {
                  sFileNameX = sFileNameX.Substring( 1, sFileNameX.Length - 2);

                  sFileForDownload_FileName = sFileNameX;

                  break;
                }
              }
            }
          }
        }


        // Creatre file
        //   and save binary data of HTTP REsponse
        //   to that file.
        if (sFileForDownload_FileName.Length > 0)
        {
          String sResponseContentFileName = sFileForDownload_FileName;
          String sResponseContentFullFileName = Path.Combine(
            m_sFolderForDownloadViaInterceptionOfHttpResponce, 
            sResponseContentFileName);

          Console.Out.WriteLine("Saving file: \"" + sResponseContentFullFileName + "\".");

          File.WriteAllBytes(sResponseContentFullFileName, e.ResponseContent.ReadAsByteArray());

          lock (m_oFileDownloadedViaResponse_CS)
          {
            m_sFileDownloadedViaResponse_FileName = sResponseContentFullFileName;
            m_bTradeRegisterDataFile_was_SuccessfullyDownloaded__via_HttpResponse = true;
          }
        }
        else 
        {
          throw new Exception("Cannot retrieve a file name for downloading from HTTP Response!");
        }
      }
      catch (Exception ex)
      {
        String sLogString =
          "Exception. Class name: " +
          ex.GetType().FullName + ". Message: \"" +
          ex.Message + "\"" + "\n" +
          "Stack trace:" + "\n" +
          ex.StackTrace;

        m_oJobProtocol.WriteLine(sLogString);
        m_oJobProtocol.WriteLine();
      }
      */

    }

    void PrintContentOfComboBox()
    {
      /*
       
      ReadOnlyCollection<IWebElement> oRegisterTypeComboBox_ItemList =
        oRegisterTypeComboBox.FindElements(By.TagName("option"));
      if (oRegisterTypeComboBox_ItemList != null)
      {
        Console.Out.WriteLine("Count of elements in Register Type combo box: \"" +
          oRegisterTypeComboBox_ItemList.Count + "\"");

        for (int i = 0; i < oRegisterTypeComboBox_ItemList.Count; i++)
        {
          IWebElement oComboBoxOptionElement = oRegisterTypeComboBox_ItemList[i];
          String sOptionValue = oComboBoxOptionElement.GetAttribute("value");
          if (sOptionValue != null)
          {
            Console.Out.WriteLine("Value of item " + i.ToString() + " is \"" + sOptionValue + "\".");
          }
        }

        Console.Out.WriteLine();
      }

      */
    }

    IWebElement TryFindElementAtWebPageById(
      WebDriver v_oBrowserWrapper,
      String v_sId)
    {
      IWebElement oResult = null;

      ReadOnlyCollection<IWebElement> oFondElemnts = null;

      oFondElemnts = v_oBrowserWrapper.FindElements(By.Id(v_sId));
      if (oFondElemnts.Count > 0)
        oResult = oFondElemnts[0];

      return oResult;
    }

    IWebElement FindElementAtWebPageById_or_ThrowException(
      WebDriver v_oBrowserWrapper,
      String v_sId)
    {
      IWebElement oResult = null;

      oResult = TryFindElementAtWebPageById(v_oBrowserWrapper, v_sId);
      if (oResult == null)
        throw new Exception("Cannot find at current Web Page an element with id='" + v_sId + "'.");

      return oResult;
    }

    void WaitUntilWebElementIsActive(
      WebDriver v_oBrowserWrapper,
      IWebElement v_oWebElement,
      UInt32 v_uiWaitTimeInSeconds)
    {
      if (v_oBrowserWrapper == null)
        throw new Exception("Input argumnet v_oBrowserWrapper is 'null'.");

      if (v_oWebElement == null)
        throw new Exception("Input argumnet v_oWebElement is 'null'.");

      if (v_uiWaitTimeInSeconds > 1000)
        throw new Exception("Input argumnet v_uiWaitTimeInSeconds > 1000.");

      // Scroll element into view first to ensure it's visible
      try
      {
        IJavaScriptExecutor js = (IJavaScriptExecutor)v_oBrowserWrapper;
        js.ExecuteScript("arguments[0].scrollIntoView({block: 'center'});", v_oWebElement);
        Thread.Sleep(300);
      }
      catch
      {
        // If scrolling fails, continue anyway
      }

      DateTime oStartDateTime = DateTime.Now;

      bool bExitLoop = false;

      while (!bExitLoop)
      {
        IWebElement oActiveElement = v_oBrowserWrapper.SwitchTo().ActiveElement();

        if (oActiveElement != null &&
            oActiveElement.Equals(v_oWebElement))
        {
          bExitLoop = true;
        }
        else
        {
          DateTime oEndDateTime = DateTime.Now;
          TimeSpan oElapsedTime = oEndDateTime - oStartDateTime;

          if (oElapsedTime.TotalSeconds < v_uiWaitTimeInSeconds)
          {
            Thread.Sleep(100);
          }
          else
          {
            throw new Exception("Waiting interval has been expired in function WaitUntilWebElementIsActive(). " +
              "v_uiWaitTimeInSeconds = " + v_uiWaitTimeInSeconds.ToString() + " sec.");
          }
        }
      }
    }

    // Dismisses the cookie consent banner if present
    // Uses two alternative methods for robustness
    void DismissCookieBanner(
      WebDriver v_oBrowserWrapper,
      ClJobProtocol v_oJobProtocol,
      String v_sPrefixStringForProtocol)
    {
      try
      {
        // Wait for the page to fully load and cookie banner to appear
        Thread.Sleep(1000);

        bool bCookieBannerDismissed = false;

        // Method 1: Try to click the "Okay" button directly by element ID
        try
        {
          IWebElement oCookieButton = TryFindElementAtWebPageById(v_oBrowserWrapper, "cookieForm:j_idt18");
          if (oCookieButton != null && oCookieButton.Displayed && oCookieButton.Enabled)
          {
            oCookieButton.Click();
            Thread.Sleep(500);
            bCookieBannerDismissed = true;
            v_oJobProtocol.WriteLine(v_sPrefixStringForProtocol + "Cookie banner dismissed (Method 1: Direct click).");
          }
        }
        catch (Exception ex1)
        {
          v_oJobProtocol.WriteLine(v_sPrefixStringForProtocol + "Method 1 failed: " + ex1.Message);
        }

        // Method 2: Try to find and click using XPath and CSS selectors
        if (!bCookieBannerDismissed)
        {
          try
          {
            // Try finding by class name
            var oCookieButtons = v_oBrowserWrapper.FindElements(By.ClassName("cookie-btn"));
            if (oCookieButtons != null && oCookieButtons.Count > 0)
            {
              IWebElement oButton = oCookieButtons[0];
              if (oButton.Displayed && oButton.Enabled)
              {
                oButton.Click();
                Thread.Sleep(500);
                bCookieBannerDismissed = true;
                v_oJobProtocol.WriteLine(v_sPrefixStringForProtocol + "Cookie banner dismissed (Method 2: Class name).");
              }
            }
          }
          catch (Exception ex2)
          {
            v_oJobProtocol.WriteLine(v_sPrefixStringForProtocol + "Method 2 failed: " + ex2.Message);
          }
        }

        // Method 3: Use JavaScript to hide the banner if clicking fails
        if (!bCookieBannerDismissed)
        {
          try
          {
            IJavaScriptExecutor js = (IJavaScriptExecutor)v_oBrowserWrapper;
            js.ExecuteScript("var cookiePanel = document.getElementById('cookieForm:cookiePanel'); if(cookiePanel) { cookiePanel.style.display = 'none'; }");
            Thread.Sleep(500);
            v_oJobProtocol.WriteLine(v_sPrefixStringForProtocol + "Cookie banner dismissed (Method 3: JavaScript hide).");
          }
          catch (Exception ex3)
          {
            v_oJobProtocol.WriteLine(v_sPrefixStringForProtocol + "Method 3 failed: " + ex3.Message);
          }
        }
      }
      catch (Exception ex)
      {
        // Non-critical error - continue even if cookie dismissal fails
        v_oJobProtocol.WriteLine(v_sPrefixStringForProtocol + "Warning: Could not dismiss cookie banner: " + ex.Message);
      }
    }

    // Returns new Active Web Element
    void WaitUntilChangeOfActiveWebElement(
      WebDriver v_oBrowserWrapper,
      IWebElement v_oPrevActiveWebElement,
      UInt32 v_uiWaitTimeInSeconds)
    {
      if (v_oBrowserWrapper == null)
        throw new Exception("Input argumnet v_oBrowserWrapper is 'null'.");

      if (v_uiWaitTimeInSeconds > 1000)
        throw new Exception("Input argumnet v_uiWaitTimeInSeconds > 1000.");


      DateTime oStartDateTime = DateTime.Now;

      bool bExitLoop = false;

      while (!bExitLoop)
      {
        IWebElement oActiveElement = v_oBrowserWrapper.SwitchTo().ActiveElement();

        bool bActiveElementHadChanged = false;

        if (v_oPrevActiveWebElement == null)
        {
          if (oActiveElement != null)
            bActiveElementHadChanged = true;
        }
        else
        {
          if (oActiveElement != null)
          {
            if (!oActiveElement.Equals(v_oPrevActiveWebElement))
              bActiveElementHadChanged = true;
          }
        }

        if (bActiveElementHadChanged)
        {
          bExitLoop = true;
        }
        else
        {
          DateTime oEndDateTime = DateTime.Now;
          TimeSpan oElapsedTime = oEndDateTime - oStartDateTime;

          if (oElapsedTime.TotalSeconds < v_uiWaitTimeInSeconds)
          {
            Thread.Sleep(100);
          }
          else
          {
            throw new Exception("Waiting interval has been expired in function WaitUntilChangeOfActiveWebElement(). " +
              "v_uiWaitTimeInSeconds = " + v_uiWaitTimeInSeconds.ToString() + " sec.");
          }
        }
      }
    }

    void WaitUntilBrowserUnlocksFile(
      String v_sTheNameOfFileThatIsBeenDownloaded,
      Int32 v_iWaitTimeInSeconds)
    {
      bool bFileIsLocked = false;
      DateTime oStartDateTime = DateTime.Now;

      do
      {
        bFileIsLocked = false;

        if (File.Exists(v_sTheNameOfFileThatIsBeenDownloaded))
        {
          FileStream oTempFielStream = null;

          try
          {
            oTempFielStream = File.Open(v_sTheNameOfFileThatIsBeenDownloaded, FileMode.Open, FileAccess.Read, FileShare.None);
          }
          catch (System.IO.IOException)
          {
            bFileIsLocked = true;
          }

          if (oTempFielStream != null)
            oTempFielStream.Close();
        }
        else
        {
          break;
        }

        if (bFileIsLocked)
        {
          DateTime oEndDateTime = DateTime.Now;
          TimeSpan oElapsedTime = oEndDateTime - oStartDateTime;

          if (oElapsedTime.TotalSeconds > v_iWaitTimeInSeconds)
          {
            throw new Exception("Cannot lock dowloaded file during specified wait period (" +
              v_iWaitTimeInSeconds.ToString() + " sec.): \"" +
              v_sTheNameOfFileThatIsBeenDownloaded + "\"");
          }
          else
          {
            Thread.Sleep(100);
          }
        }
      }
      while (bFileIsLocked);
    }

    // Returns the name of a new file,
    // I. e. the name of file,
    //   that web browser started to download.
    String WaitUntilBrowserStartToDownloadFile(
      String v_sBrowserDownloadFolder,
      String[] v_aoFileNameListBeforeClickForDownload,
      Int32 v_iWaitTimeInSeconds,
      out Boolean v_bNewFileWasCreated_In_BrowserDownloadFolder)
    {
      String sNewFileName = "";
      v_bNewFileWasCreated_In_BrowserDownloadFolder = false;
      Boolean bExitLoop = false;

      DateTime oStartDateTime = DateTime.Now;

      do
      {
        String[] aoFileNameListAfterClickForDownload = null;

        aoFileNameListAfterClickForDownload = Directory.GetFiles(v_sBrowserDownloadFolder);
        if (aoFileNameListAfterClickForDownload.Length > v_aoFileNameListBeforeClickForDownload.Length)
        {
          for (int i = 0; i < aoFileNameListAfterClickForDownload.Length; i++)
          {
            String sFileNameFromActualList = aoFileNameListAfterClickForDownload[i];

            if (!sFileNameFromActualList.EndsWith(".tmp") &&
                !sFileNameFromActualList.EndsWith(".crdownload"))
            {
              bool bFilePresentInPreviosList = false;

              for (int j = 0; j < v_aoFileNameListBeforeClickForDownload.Length; j++)
              {
                String sFileNameFromPreviousList = v_aoFileNameListBeforeClickForDownload[j];

                if (sFileNameFromActualList == sFileNameFromPreviousList)
                {
                  bFilePresentInPreviosList = true;
                  break;
                }
              }

              if (!bFilePresentInPreviosList)
              {
                sNewFileName = sFileNameFromActualList;
                break;
              }
            }
          }
        }

        if (sNewFileName != "")
        {
          v_bNewFileWasCreated_In_BrowserDownloadFolder = true;

          bExitLoop = true;
        }
        else
        {
          DateTime oEndDateTime = DateTime.Now;
          TimeSpan oElapsedTime = oEndDateTime - oStartDateTime;

          if (oElapsedTime.TotalSeconds > v_iWaitTimeInSeconds)
          {
            bExitLoop = true;
          }
          else
          {
            Thread.Sleep(100);
          }
        }
      }
      while (!bExitLoop);

      return sNewFileName;
    }

    void CheckIfWebPageContainsErrorMessage(
      WebDriver v_oBrowserWrapper,
      ClJobProtocol v_oJobProtocol,
      String v_sInfoPartOfErrorFileName,
      out String v_sErrorOfAType_MessageText,
      out Boolean v_bErrorOfAType_At_TradeRegisterPortal,
      out Boolean v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
    {
      v_sErrorOfAType_MessageText = "";
      v_bErrorOfAType_At_TradeRegisterPortal = false;
      v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = false;

      // Check if one of special error pages was shown
      //   by Web Server of DEU Trade Register portal.

      String sHTMLPageSource = v_oBrowserWrapper.PageSource;

      //
      // <p>Errorcode: A 110
      //
      // </p>
      //
      // Due to technical problems
      // you cannot access this document now.
      // Please try again later on.
      // If you have any further questions,
      // please contact the service centre by phone
      // on +49 (0)2331 985 112.
      //
      // Also detect:
      // HTTP Status 500 - org.jboss.weld.contexts.BusyConversationException: WELD-000322: Conversation lock timed out
      //
      if (sHTMLPageSource.IndexOf("Errorcode: A") != -1 ||
          sHTMLPageSource.IndexOf("HTTP Status 500") != -1 ||
          sHTMLPageSource.IndexOf("BusyConversationException") != -1 ||
          sHTMLPageSource.IndexOf("WELD-000322") != -1)
      {
        v_bErrorOfAType_At_TradeRegisterPortal = true;

        String sErrroCodeForFileName = "";

        int iStartOfMessagePos = sHTMLPageSource.IndexOf("Errorcode: A");
        if (iStartOfMessagePos != -1)
        {
          String sMessageText = sHTMLPageSource.Substring(iStartOfMessagePos);

          int iEndPos = sMessageText.IndexOf("</p>");

          if (iEndPos != -1)
            sMessageText = sMessageText.Remove(iEndPos);

          char[] achCharsForTriming = new char[] { ' ', '\t', '\n', '\r' };

          sMessageText = sMessageText.TrimEnd(achCharsForTriming);

          v_sErrorOfAType_MessageText = sMessageText;


          if (sMessageText.Length > 11)
          {
            sErrroCodeForFileName = sMessageText.Substring(11);

            if (sErrroCodeForFileName.IndexOfAny(Path.GetInvalidFileNameChars()) != -1 ||
                sErrroCodeForFileName.IndexOfAny(Path.GetInvalidPathChars()) != -1)
            {
              sErrroCodeForFileName = "";
            }
          }
        }
        else if (sHTMLPageSource.IndexOf("HTTP Status 500") != -1 ||
                 sHTMLPageSource.IndexOf("BusyConversationException") != -1)
        {
          v_sErrorOfAType_MessageText = "HTTP Status 500 - BusyConversationException";
          sErrroCodeForFileName = "HTTP500_BusyConversation";
        }



        String sHTMLPageSourceFileName = Path.Combine(m_sLogFolder,
          "HTMLPageSourceFor_Errorcode" + "_" +
          sErrroCodeForFileName + "_" +
          v_sInfoPartOfErrorFileName + "_" +
          DateTime.Now.Ticks.ToString() + ".txt");

        if (File.Exists(sHTMLPageSourceFileName))
          File.Delete(sHTMLPageSourceFileName);

        File.WriteAllText(sHTMLPageSourceFileName, sHTMLPageSource, Encoding.Unicode);

        v_oJobProtocol.WriteLine("DEU Trage Register Web Site error! " +
          "HTML source of Web Page was saved to the following file: \"" +
          sHTMLPageSourceFileName + "\".");




        v_oBrowserWrapper.Navigate().Back();
        Thread.Sleep(1000);

        sHTMLPageSource = v_oBrowserWrapper.PageSource;
        if (sHTMLPageSource.IndexOf("<p>Error ID: 440</p>") != -1 || 
            sHTMLPageSource.IndexOf("Errorcode: 0") != -1)
        {
          v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = true;
          v_bErrorOfAType_At_TradeRegisterPortal = false;
        }
      }
      else
      {
        // 
        // <h1>Expired session:</h1>
        // <p>Error ID: 440</p>
        // Also check for Errorcode: 0 which indicates session issues
        if (sHTMLPageSource.IndexOf("<p>Error ID: 440</p>") != -1 || 
            sHTMLPageSource.IndexOf("Errorcode: 0") != -1)
        {
          v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = true;
        }
        else
        {
          // throw new Exception(
          //   "Cannot find any new files inside 'Downloads' folder " +
          //   "of Web Browser during specified wait period.\n" +
          //   "iWaitTimeInSeconds == " + iWaitTimeInSeconds.ToString() + " sec.,\n" +
          //   "v_sBrowserDownloadFolder == \"" + m_sBrowserDownloadFolder + "\".");
        }
      }
    }

    String DownloadFileByClickingOnAncor(
      WebDriver v_oBrowserWrapper,
      ClJobProtocol v_oJobProtocol,
      String v_sInfoPartOfErrorFileName,
      IWebElement v_oAncor,
      out String v_sErrorOfAType_MessageText,
      out Boolean v_bErrorOfAType_At_TradeRegisterPortal,
      out Boolean v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
    {
      v_sErrorOfAType_MessageText = "";
      v_bErrorOfAType_At_TradeRegisterPortal = false;
      v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = false;


      String[] aoFileNameListBeforeClickForDownload = null;
      String sDownloadedFileName = null;
      Boolean bNewFileWasCreated_In_BrowserDownloadFolder = false;
      int iWaitTimeInSeconds_for_ErrorWebPage = 20;
      int iWaitTimeInSeconds_for_DataFileToBeDownloaded = 100;

      aoFileNameListBeforeClickForDownload = Directory.GetFiles(m_sBrowserDownloadFolder);

      /*
      lock (m_oFileDownloadedViaResponse_CS)
      {
        m_sFileDownloadedViaResponse_FileName = "";
        m_bTradeRegisterDataFile_was_SuccessfullyDownloaded__via_HttpResponse = false;
      }
      */

      // v_oAncor.Click();

      v_oAncor.SendKeys(Keys.Enter);


      /*
      for (int t = 0; t < 100; t++)
      {
        Thread.Sleep(100);

        lock (m_oFileDownloadedViaResponse_CS)
        {
          if (m_bTradeRegisterDataFile_was_SuccessfullyDownloaded__via_HttpResponse == true)
          {
            sDownloadedFileName = m_sFileDownloadedViaResponse_FileName;
            bNewFileWasCreated_In_BrowserDownloadFolder = true;

            m_sFileDownloadedViaResponse_FileName = "";
            m_bTradeRegisterDataFile_was_SuccessfullyDownloaded__via_HttpResponse = false;

            break;
          }
        }
      }
      */

      // Step 1: Wait a short period of time,
      //   because Web Page containing Error code may be displayed.
      sDownloadedFileName = WaitUntilBrowserStartToDownloadFile(
        m_sBrowserDownloadFolder,
        aoFileNameListBeforeClickForDownload,
        iWaitTimeInSeconds_for_ErrorWebPage,
        out bNewFileWasCreated_In_BrowserDownloadFolder);


      if (bNewFileWasCreated_In_BrowserDownloadFolder)
      {
        // -- Main approach: Downloading of file by browser.
        WaitUntilBrowserUnlocksFile(sDownloadedFileName, 10);
      }
      else
      {
        // Check if one of special error pages was shown
        //   by Web Server of DEU Trade Register portal.

        CheckIfWebPageContainsErrorMessage(
          v_oBrowserWrapper,
          v_oJobProtocol,
          v_sInfoPartOfErrorFileName,
          out v_sErrorOfAType_MessageText,
          out v_bErrorOfAType_At_TradeRegisterPortal,
          out v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440);

        // Close any extra windows that might have been opened with error page
        if (v_bErrorOfAType_At_TradeRegisterPortal || v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
        {
          try
          {
            var allHandles = v_oBrowserWrapper.WindowHandles;
            if (allHandles.Count > 1)
            {
              // Keep the first window, close all others
              string mainWindow = allHandles[0];
              for (int i = allHandles.Count - 1; i > 0; i--)
              {
                v_oBrowserWrapper.SwitchTo().Window(allHandles[i]);
                v_oBrowserWrapper.Close();
                v_oJobProtocol.WriteLine("Closed extra browser window/tab that showed error page.");
              }
              v_oBrowserWrapper.SwitchTo().Window(mainWindow);
            }
          }
          catch (Exception exWin)
          {
            v_oJobProtocol.WriteLine("Warning: Error closing extra browser windows: " + exWin.Message);
          }
        }

        if (!v_bErrorOfAType_At_TradeRegisterPortal &&
            !v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
        {

          // Step 2: Wait a long period of time,
          //   because Web Server may be busy with client requests.
          sDownloadedFileName = WaitUntilBrowserStartToDownloadFile(
            m_sBrowserDownloadFolder,
            aoFileNameListBeforeClickForDownload,
            iWaitTimeInSeconds_for_DataFileToBeDownloaded,
            out bNewFileWasCreated_In_BrowserDownloadFolder);


          if (bNewFileWasCreated_In_BrowserDownloadFolder)
          {
            // -- Main approach: Downloading of file by browser.
            WaitUntilBrowserUnlocksFile(sDownloadedFileName, 10);
          }
          else
          {
            // Check if one of special error pages was shown
            //   by Web Server of DEU Trade Register portal.

            CheckIfWebPageContainsErrorMessage(
              v_oBrowserWrapper,
              v_oJobProtocol,
              v_sInfoPartOfErrorFileName,
              out v_sErrorOfAType_MessageText,
              out v_bErrorOfAType_At_TradeRegisterPortal,
              out v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440);

            // Close any extra windows that might have been opened with error page
            if (v_bErrorOfAType_At_TradeRegisterPortal || v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
            {
              try
              {
                var allHandles = v_oBrowserWrapper.WindowHandles;
                if (allHandles.Count > 1)
                {
                  // Keep the first window, close all others
                  string mainWindow = allHandles[0];
                  for (int i = allHandles.Count - 1; i > 0; i--)
                  {
                    v_oBrowserWrapper.SwitchTo().Window(allHandles[i]);
                    v_oBrowserWrapper.Close();
                    v_oJobProtocol.WriteLine("Closed extra browser window/tab that showed error page.");
                  }
                  v_oBrowserWrapper.SwitchTo().Window(mainWindow);
                }
              }
              catch (Exception exWin)
              {
                v_oJobProtocol.WriteLine("Warning: Error closing extra browser windows: " + exWin.Message);
              }
            }

            if (!v_bErrorOfAType_At_TradeRegisterPortal &&
                !v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
            {
              int iWaitTimeInSeconds = 0;

              iWaitTimeInSeconds =
                iWaitTimeInSeconds_for_ErrorWebPage +
                iWaitTimeInSeconds_for_DataFileToBeDownloaded;

              throw new Exception(
                "Cannot find any new files inside 'Downloads' folder " +
                "of Web Browser during specified wait period." +
                Environment.NewLine +
                "iWaitTimeInSeconds == " + iWaitTimeInSeconds.ToString() + " sec.," + 
                Environment.NewLine +
                "v_sBrowserDownloadFolder == \"" + m_sBrowserDownloadFolder + "\".");
            }
          }
        }
      }

      return sDownloadedFileName;
    }

    void DownloadFilesFromTableOfSearchResults(
      WebDriver v_oBrowserWrapper,
      ClJobProtocol v_oJobProtocol,
      String v_sPrefixStringForProtocol,
      String v_sTradeRegisterNumber,
      out Boolean v_bErrorOfAType_At_TradeRegisterPortal,
      out Boolean v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
    {
      v_bErrorOfAType_At_TradeRegisterPortal = false;
      v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = false;

      try
      {
        DownloadFilesFromTableOfSearchResults_Internal(
          v_oBrowserWrapper,
          v_oJobProtocol,
          v_sPrefixStringForProtocol,
          v_sTradeRegisterNumber,
          out v_bErrorOfAType_At_TradeRegisterPortal,
          out v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440);
      }
      catch (OpenQA.Selenium.StaleElementReferenceException exStale)
      {
        v_oJobProtocol.WriteLine(
          v_sPrefixStringForProtocol +
          "Warning: StaleElementReferenceException occurred (DOM changed after portal error). " +
          "Treating as session issue and will restart browser. Error: " + exStale.Message);
        
        v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = true;
        v_bErrorOfAType_At_TradeRegisterPortal = false;
      }
      catch (OpenQA.Selenium.WebDriverException exWebDriver)
      {
        v_oJobProtocol.WriteLine(
          v_sPrefixStringForProtocol +
          "Warning: WebDriverException occurred (timeout or communication failure with browser). " +
          "Treating as session issue and will restart browser. Error: " + exWebDriver.Message);
        
        v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = true;
        v_bErrorOfAType_At_TradeRegisterPortal = false;
      }
    }

    void DownloadFilesFromTableOfSearchResults_Internal(
      WebDriver v_oBrowserWrapper,
      ClJobProtocol v_oJobProtocol,
      String v_sPrefixStringForProtocol,
      String v_sTradeRegisterNumber,
      out Boolean v_bErrorOfAType_At_TradeRegisterPortal,
      out Boolean v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
    {
      v_bErrorOfAType_At_TradeRegisterPortal = false;
      v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = false;

      DateTime oStartTime = DateTime.Now;

      v_oJobProtocol.WriteLine(
        v_sPrefixStringForProtocol +
        "Start processing of page with Search Results. " +
        oStartTime.ToString("yyyy.MM.dd HH:mm:ss"));

      IWebElement oBodyElement = v_oBrowserWrapper.FindElement(By.TagName("body"));

      v_oJobProtocol.WriteLine(
        v_sPrefixStringForProtocol +
        "Move keyboard focus to the first link for download of 'company profile' document.");

      // Move keyboard focus to the first link for download of 'company profile' document.
      for (int i = 0; i < 34; i++)
      {
        oBodyElement.SendKeys(Keys.Tab);
      }

      Thread.Sleep(100);


      List<IWebElement> oAncorList = new List<IWebElement>();

      ReadOnlyCollection<IWebElement> oFoundSpanElemnts = null;

      oFoundSpanElemnts = v_oBrowserWrapper.FindElements(By.XPath("//span[@class='underlinedText']"));

      for (int i = 0; i < oFoundSpanElemnts.Count; i++)
      {
        IWebElement oCurSpanElement = oFoundSpanElemnts[i];

        String sInnerText = oCurSpanElement.Text;

        // Statistics for HRB 30704:
        //  - for "SI": 48 seconds,
        //  - for "AD": 41 seconds.

        if (sInnerText == "SI")
        {
          IWebElement oParentWebElement = oCurSpanElement.FindElement(By.XPath("parent::node()"));

          if (oParentWebElement != null)
          {
            String sParentWebElement_TagName = oParentWebElement.TagName;

            if (sParentWebElement_TagName == "a")
            {
              oAncorList.Add(oParentWebElement);
            }
          }
        }
      }



      int iNumberOfLinksForDownload = oAncorList.Count;

      v_oJobProtocol.WriteLine(
        v_sPrefixStringForProtocol +
        "Number of links for download: " + iNumberOfLinksForDownload + ". ");

      // IWebElement oActiveElement1 = oBrowserWrapper.SwitchTo().ActiveElement();
      // PrintFieldsOfWebElement(oJobProtocol, oActiveElement1, "oActiveElement1");

      // "After last" element at Search Page  <a class="ui-paginator-page ui-state-default ui-state-active ui-corner-all" 

      v_oJobProtocol.WriteLine(
        v_sPrefixStringForProtocol + "Start loop of processing of links for download.");

      int iCurentFileId = 1;
      int iUnsuccessfulLinkCount = 0;
      List<String> oDownloadedFileList = new List<string>();
      bool bExitLoopOnResultsPage = false;

      do
      {
        IWebElement oCurActiveWebElement = v_oBrowserWrapper.SwitchTo().ActiveElement();

        if (oCurActiveWebElement == null)
          throw new Exception("oCurActiveWebElement == null");

        String sActiveWebElement_TagName = oCurActiveWebElement.TagName;
        String sActiveWebElement_Class = oCurActiveWebElement.GetAttribute("class");
        String sActiveWebElement_Text = oCurActiveWebElement.Text;

        Boolean bPaginatorReached = false;
        Boolean bFooterReached = false;

        if (sActiveWebElement_TagName == "a" &&
            sActiveWebElement_Class == "ui-paginator-page ui-state-default ui-state-active ui-corner-all ui-state-focus")
        {
          bPaginatorReached = true;
        }
        else
        {
          if (sActiveWebElement_TagName == "a" &&
              sActiveWebElement_Class == "ui-commandlink ui-widget footerText" &&
              sActiveWebElement_Text == "Website Details")
          {
            // The code below was written for case,
            //   when 0 records were found.
            //
            // For example,
            //   search by Register Number "200000" finds 0 records.
            bFooterReached = true;
          }
        }

        if (bPaginatorReached ||
            bFooterReached)
        {
          bExitLoopOnResultsPage = true;

          v_oJobProtocol.WriteLine(
            v_sPrefixStringForProtocol + "Exit loop of processing of links for download.");
        }
        else
        {
          if (sActiveWebElement_TagName == "a" &&
              sActiveWebElement_Class == "ui-commandlink ui-widget dokumentList"
              )
          {
            if (sActiveWebElement_Text == "SI")
            {
              String sInfoPartOfErrorFileName = "";

              sInfoPartOfErrorFileName =
                "HRB_" + v_sTradeRegisterNumber + "_" +
                "(" + iCurentFileId.ToString() + "_of_" + iNumberOfLinksForDownload.ToString() + ")";

              String sErrorOfAType_MessageText = "";
              String sDownloadedFileName = "";

              // Check if file already exists and is valid
              String sExistingFile = FindExistingDownloadedFile(v_sTradeRegisterNumber, iCurentFileId, v_oJobProtocol);
              
              if (!String.IsNullOrEmpty(sExistingFile))
              {
                // File already exists and XML is valid, skip download
                v_oJobProtocol.WriteLine(
                  v_sPrefixStringForProtocol +
                  $"File already exists and is valid, skipping download [{iCurentFileId} of {iNumberOfLinksForDownload}]: {Path.GetFileName(sExistingFile)}");
                
                sDownloadedFileName = sExistingFile;
                
                // Navigate to next link without downloading
                oBodyElement.SendKeys(Keys.Tab);
                WaitUntilChangeOfActiveWebElement(v_oBrowserWrapper, oCurActiveWebElement, 5);
              }
              else
              {
                // File doesn't exist or is invalid, proceed with download
                oBodyElement.SendKeys(Keys.Tab);
                WaitUntilChangeOfActiveWebElement(v_oBrowserWrapper, oCurActiveWebElement, 5);

                IWebElement oCurActiveWebElement2 = v_oBrowserWrapper.SwitchTo().ActiveElement();
                oBodyElement.SendKeys(Keys.Shift + Keys.Tab);
                WaitUntilChangeOfActiveWebElement(v_oBrowserWrapper, oCurActiveWebElement2, 5);


                IWebElement oSI_Ancor = v_oBrowserWrapper.SwitchTo().ActiveElement();

                sDownloadedFileName = DownloadFileByClickingOnAncor(
                  v_oBrowserWrapper,
                  v_oJobProtocol,
                  sInfoPartOfErrorFileName,
                  oSI_Ancor,
                  out sErrorOfAType_MessageText,
                  out v_bErrorOfAType_At_TradeRegisterPortal,
                  out v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440);
              }


              if (!v_bErrorOfAType_At_TradeRegisterPortal &&
                  !v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
              {
                if (!String.IsNullOrEmpty(sDownloadedFileName))
                {
                  v_oJobProtocol.WriteLine(
                    v_sPrefixStringForProtocol +
                    "Data file was successfully downloaded " +
                    "(" + iCurentFileId.ToString() + " of " + iNumberOfLinksForDownload.ToString() + "): " +
                    "\"" + sDownloadedFileName + "\".");

                  oDownloadedFileList.Add(sDownloadedFileName);
                }

                iCurentFileId++;
              }
              else
              {
                if (v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
                {
                  v_oJobProtocol.WriteLine(
                    v_sPrefixStringForProtocol +
                    "Step 'Downloading of .xml files for found companies'. " +
                    "An error at DEU Trade Register web server has occured. " +
                    "\"Error ID: 440 (Expired session.)\" " +
                    "Register number \"" + v_sTradeRegisterNumber + "\" was NOT downloaded. " +
                    "Exit loop of processing of links for download. " +
                    "I will try to restart web browser " +
                    "and repeat search for this register number again.");

                  bExitLoopOnResultsPage = true;
                }
                else
                {
                  Boolean bSkipDownloadingOfCurrentAncor = false;


                  // if (sErrorOfAType_MessageText == "Errorcode: A 200")
                  // {
                  //   bSkipDownloadingOfCurrentAncor = true;
                  // }

                  // Skip all links with error of type "Errorcode: A".
                  bSkipDownloadingOfCurrentAncor = true;
                  v_bErrorOfAType_At_TradeRegisterPortal = false;


                  if (bSkipDownloadingOfCurrentAncor)
                  {
                    iUnsuccessfulLinkCount++;

                    DEU_TR_ProgramSettings.InsertOrUpdate_UnsuccessfullyDownloadedRegisterNumber__in__ListMode_Settings(
                      v_sTradeRegisterNumber,
                      false,
                      sErrorOfAType_MessageText,
                      iUnsuccessfulLinkCount,
                      iNumberOfLinksForDownload);

                    v_oJobProtocol.WriteLine(
                      v_sPrefixStringForProtocol +
                      "The following link was skipped for downloading " +
                      "because of internal error at DEU Trade Register Web Site: " +
                      "Position (" + iCurentFileId.ToString() + " of " + iNumberOfLinksForDownload.ToString() + "). " +
                      "Erorr string: \"" + sErrorOfAType_MessageText + "\".");

                    iCurentFileId++;
                  }
                  else
                  {
                    v_oJobProtocol.WriteLine(
                      v_sPrefixStringForProtocol +
                      "Step 'Downloading of .xml files for found companies'. " +
                      "An error at DEU Trade Register web server has occured. " +
                      "\"Errorcode: A nnn\" " +
                      "(\"" + sErrorOfAType_MessageText + "\") " +
                      "Register number \"" + v_sTradeRegisterNumber + "\" was NOT downloaded. " +
                      "Exit loop of processing of links for download. " +
                      "I will try to restart web browser " +
                      "and repeat search for this register number again.");

                    bExitLoopOnResultsPage = true;
                  }
                }
              }
            }
          }
          else
          {
            if (sActiveWebElement_Text == "[ Show all ]" ||
                sActiveWebElement_Text == "[ Mask ]" ||
                sActiveWebElement_Text == "Status information" ||

                sActiveWebElement_Text == "[ Alle anzeigen ]" ||
                sActiveWebElement_Text == "[ Ausblenden ]" ||
                sActiveWebElement_Text == "Statusinformationen")
            {
              // Just skip this links
            }
            else
            {
              PrintFieldsOfWebElement(v_oJobProtocol, oCurActiveWebElement, "oCurActiveWebElement");

              throw new Exception("Unexpected element in keyboard 'Tab' sequence!");
            }
          }
        }


        // Move focus of keyboard to the next element on current Web Page
        if (!bExitLoopOnResultsPage)
        {
          oBodyElement.SendKeys(Keys.Tab);

          WaitUntilChangeOfActiveWebElement(v_oBrowserWrapper, oCurActiveWebElement, 5);
        }

      } while (!bExitLoopOnResultsPage);


      // Sleep for Virus Scan.
      Thread.Sleep(500);


      // Move all downloaded files from "Downloads" folder of Web Browser
      //   to "Output" folder of my program.
      if (!v_bErrorOfAType_At_TradeRegisterPortal &&
          !v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 )
      {
        // Create sub-folder in "Output" folder for current Register Number.
        String sOutputFolder = Path.Combine(m_sOutputFolder, v_sTradeRegisterNumber);

        if (Directory.Exists(sOutputFolder))
          Directory.Delete(sOutputFolder, true);

        Directory.CreateDirectory(sOutputFolder);

        // Move files.
        for (int i = 0; i < oDownloadedFileList.Count(); i++)
        {
          String sDownloadedFileName = oDownloadedFileList[i];
          String sFileNameWOPath = Path.GetFileName(sDownloadedFileName);
          String sFileNameForOutputFolder = Path.Combine(sOutputFolder, sFileNameWOPath);

          if (File.Exists(sFileNameForOutputFolder))
            File.Delete(sFileNameForOutputFolder);

          File.Move(sDownloadedFileName, sFileNameForOutputFolder);
          
          // Parse XML and write to CSV immediately after download
          ParseAndWriteXMLToCSV(sFileNameForOutputFolder, v_oJobProtocol, v_sPrefixStringForProtocol);
        }


        v_oJobProtocol.WriteLine(
          v_sPrefixStringForProtocol +
          "All downloaded files were moved from Browser's \"Downloads\" folder " +
          "to the following folder: \"" + sOutputFolder + "\".");


        //
        // // Save HTML source code of current page to a special file
        // //   for potential future controlling of correctness
        // //   of site extraction process.
        // String sHTMLPageSource = v_oBrowserWrapper.PageSource;
        // String sPageSourceFileName = Path.Combine(v_sOutputFolder, "PageSource.html");
        //
        // if (File.Exists(sPageSourceFileName))
        //   File.Delete(sPageSourceFileName);
        // 
        // File.WriteAllText(sPageSourceFileName, sHTMLPageSource);
        //
      }


      DateTime oEndTime = DateTime.Now;
      TimeSpan oWorkTimeSpan = oEndTime - oStartTime;

      v_oJobProtocol.WriteLine(
        v_sPrefixStringForProtocol +
        "End processing of page with Search Results: " +
        oEndTime.ToString("yyyy.MM.dd HH:mm:ss"));

      v_oJobProtocol.WriteLine(
        v_sPrefixStringForProtocol +
        "Processing time span of page with Search Results: " + oWorkTimeSpan.ToString());
    }

    String DownloadFile_by_Executing_of_OnClick_JavaScript(
      WebDriver v_oBrowserWrapper,
      ClJobProtocol v_oJobProtocol,
      String v_sOnClick_JavaScript_Str,
      out Boolean v_bErrorOfAType_At_TradeRegisterPortal)
    {
      v_bErrorOfAType_At_TradeRegisterPortal = false;

      String[] aoFileNameListBeforeClickForDownload = null;
      String sDownloadedFileName = null;
      Boolean bNewFileWasCreated_In_BrowserDownloadFolder = false;
      int iWaitTimeInSeconds = 10;

      aoFileNameListBeforeClickForDownload = Directory.GetFiles(m_sBrowserDownloadFolder);


      v_oJobProtocol.WriteLine("\"onclick\" == \"" + v_sOnClick_JavaScript_Str + "\".");


      // Force Web Browser do download an .xml data file
      //   from DEU Trade Register portal.
      v_oBrowserWrapper.ExecuteScript(v_sOnClick_JavaScript_Str);


      sDownloadedFileName = WaitUntilBrowserStartToDownloadFile(
        m_sBrowserDownloadFolder,
        aoFileNameListBeforeClickForDownload,
         iWaitTimeInSeconds,
        out bNewFileWasCreated_In_BrowserDownloadFolder);


      if (bNewFileWasCreated_In_BrowserDownloadFolder)
      {
        WaitUntilBrowserUnlocksFile(sDownloadedFileName, 10);
      }
      else
      {
        String sHTMLPageSource = v_oBrowserWrapper.PageSource;

        if (sHTMLPageSource.IndexOf("Errorcode: ") != -1)
        {
          v_bErrorOfAType_At_TradeRegisterPortal = true;
        }
        else
        {
          throw new Exception(
            "Cannot find any new files inside 'Downloads' folder " +
            "of Web Browser during specified wait period." +
            Environment.NewLine +
            "iWaitTimeInSeconds == " + iWaitTimeInSeconds.ToString() + " sec.," +
            Environment.NewLine +
            "v_sBrowserDownloadFolder == \"" + m_sBrowserDownloadFolder + "\".");
        }
      }

      return sDownloadedFileName;
    }

    void DownloadFilesFromTableOfSearchResults_by_Executing_of_OnClick_JavaScript(
      WebDriver v_oBrowserWrapper,
      ClJobProtocol v_oJobProtocol,
      String v_sPrefixStringForProtocol,
      String v_sTradeRegisterNumber)
    {
      DateTime oStartTime = DateTime.Now;

      v_oJobProtocol.WriteLine(
        v_sPrefixStringForProtocol +
        "Start processing of page with Search Results (\"OnClick_JavaScript\" approach). " +
        oStartTime.ToString("yyyy.MM.dd HH:mm:ss"));


      IWebElement oBodyElement = v_oBrowserWrapper.FindElement(By.TagName("body"));

      oBodyElement.SendKeys(Keys.LeftControl + Keys.End);


      List<String> oOnClick_JavaScript_String_List = new List<string>();

      ReadOnlyCollection<IWebElement> oFoundSpanElemnts = null;

      oFoundSpanElemnts = v_oBrowserWrapper.FindElements(By.XPath("//span[@class='underlinedText']"));

      for (int i = 0; i < oFoundSpanElemnts.Count; i++)
      {
        IWebElement oCurSpanElement = oFoundSpanElemnts[i];

        String sInnerText = oCurSpanElement.Text;

        // Statistics for HRB 30704:
        //  - for "SI": 48 seconds,
        //  - for "AD": 41 seconds.

        if (sInnerText == "SI")
        {
          IWebElement oParentWebElement = oCurSpanElement.FindElement(By.XPath("parent::node()"));

          if (oParentWebElement != null)
          {
            String sParentWebElement_TagName = oParentWebElement.TagName;
            String sParentWebElement_Class = oParentWebElement.GetAttribute("class");

            v_oJobProtocol.WriteLine(
              v_sPrefixStringForProtocol +
              "sParentWebElement_TagName == \"" + sParentWebElement_TagName + "\", " +
              "sParentWebElement_Class == \"" + sParentWebElement_Class + "\", " +
              "OnClick java script == \"" + oParentWebElement.GetAttribute("onclick") + "\".");

            if (sParentWebElement_TagName == "a"
               // &&
               // sParentWebElement_Class == "ui-paginator-page ui-state-default ui-state-active ui-corner-all ui-state-focus"
               )
            {
              String sOnClick_JavaScript_Str = oParentWebElement.GetAttribute("onclick");

              oOnClick_JavaScript_String_List.Add(sOnClick_JavaScript_Str);
            }
          }
        }
      }



      int iNumberOfLinksForDownload = oOnClick_JavaScript_String_List.Count;

      v_oJobProtocol.WriteLine(
        v_sPrefixStringForProtocol +
        "Number of links for download: " + iNumberOfLinksForDownload + ". ");



      v_oJobProtocol.WriteLine(
        v_sPrefixStringForProtocol + "Start loop of processing of links for download.");

      List<String> oDownloadedFileList = new List<string>();

      for (int iLinkId = 0; iLinkId < oOnClick_JavaScript_String_List.Count; iLinkId++)
      {
        String sOnClick_JavaScript_Str = oOnClick_JavaScript_String_List[iLinkId];

        String sDownloadedFileName = "";
        Boolean bErrorOfAType_At_TradeRegisterPortal = false;


        sDownloadedFileName = DownloadFile_by_Executing_of_OnClick_JavaScript(
          v_oBrowserWrapper,
          v_oJobProtocol,
          sOnClick_JavaScript_Str,
          out bErrorOfAType_At_TradeRegisterPortal);


        if (bErrorOfAType_At_TradeRegisterPortal)
        {
          String sMessageText = "";

          String sHTMLPageSource = v_oBrowserWrapper.PageSource;

          int iStartOfMessagePos = sHTMLPageSource.IndexOf("Errorcode: ");
          if (iStartOfMessagePos != -1)
          {
            sMessageText = sHTMLPageSource.Substring(iStartOfMessagePos);

            int iEndPos = sMessageText.IndexOf("</p>");

            if (iEndPos != -1)
              sMessageText = sMessageText.Remove(iEndPos);
          }

          v_oJobProtocol.WriteLine(
            "Warning type 1. " +
            "Error of type 'A' at Web Server of DEU Trade Register Portal." +
            Environment.NewLine +
            "Message text: \"" + sMessageText + "\".");

          m_iWarningOfType1Count++;

          v_oBrowserWrapper.Navigate().Back();

          Thread.Sleep(1000);
        }
        else
        {
          if (!String.IsNullOrEmpty(sDownloadedFileName))
          {
            v_oJobProtocol.WriteLine(
              v_sPrefixStringForProtocol +
              "The following file was successfully downloaded " +
              "(" + (iLinkId + 1).ToString() + " of " + iNumberOfLinksForDownload.ToString() + "): " +
              "\"" + sDownloadedFileName + "\".");

            oDownloadedFileList.Add(sDownloadedFileName);
          }
        }
      }


      // Sleep for Virus Scan.
      Thread.Sleep(500);


      // Create sub-folder in "Output" folder for current Register Number.
      String sOutputFolder = Path.Combine(m_sOutputFolder, v_sTradeRegisterNumber);

      if (Directory.Exists(sOutputFolder))
        Directory.Delete(sOutputFolder, true);

      Directory.CreateDirectory(sOutputFolder);


      // Copy all downloaded files from "Downloads" folder of Web Browser
      //   to "Output" folder of my program.
      for (int i = 0; i < oDownloadedFileList.Count(); i++)
      {
        String sDownloadedFileName = oDownloadedFileList[i];
        String sFileNameWOPath = Path.GetFileName(sDownloadedFileName);
        String sFileNameForOutputFolder = Path.Combine(sOutputFolder, sFileNameWOPath);

        if (File.Exists(sFileNameForOutputFolder))
          File.Delete(sFileNameForOutputFolder);

        File.Move(sDownloadedFileName, sFileNameForOutputFolder);
        
        // Parse XML and write to CSV immediately after download
        ParseAndWriteXMLToCSV(sFileNameForOutputFolder, v_oJobProtocol, v_sPrefixStringForProtocol);
      }


      //
      // // Save HTML source code of current page to a special file
      // //   for potential future controlling of correctness
      // //   of site extraction process.
      // String sHTMLPageSource = v_oBrowserWrapper.PageSource;
      // String sPageSourceFileName = Path.Combine(v_sOutputFolder, "PageSource.html");
      //
      // if (File.Exists(sPageSourceFileName))
      //   File.Delete(sPageSourceFileName);
      // 
      // File.WriteAllText(sPageSourceFileName, sHTMLPageSource);
      //


      DateTime oEndTime = DateTime.Now;
      TimeSpan oWorkTimeSpan = oEndTime - oStartTime;

      v_oJobProtocol.WriteLine(
        v_sPrefixStringForProtocol +
        "End processing of page with Search Results (\"OnClick_JavaScript\" approach): " +
        oEndTime.ToString("yyyy.MM.dd HH:mm:ss"));

      v_oJobProtocol.WriteLine(
        v_sPrefixStringForProtocol +
        "Processing time span of page with Search Results: " + oWorkTimeSpan.ToString());
    }

    // Check for presence on current Web Page
    //   a special '<span>' tag with the following text:
    //  "The maximum number of 100 hits has been exceeded. Please limit your request further."
    bool CheckWhetherMore100ResultsWereFound(
      WebDriver v_oBrowserWrapper)
    {
      bool bResult = false;

      IWebElement oMore100HitsSpan = null;

      oMore100HitsSpan = TryFindElementAtWebPageById(
        v_oBrowserWrapper,
        "ergebnissForm:ergebnisseAnzahl_label");

      if (oMore100HitsSpan != null)
        bResult = true;

      return bResult;
    }

    bool IsValidXmlFile(String v_sFilePath, ClJobProtocol v_oJobProtocol)
    {
      try
      {
        if (!File.Exists(v_sFilePath))
          return false;

        FileInfo fileInfo = new FileInfo(v_sFilePath);
        if (fileInfo.Length == 0)
        {
          v_oJobProtocol.WriteLine($"Warning: File is empty: {v_sFilePath}");
          return false;
        }

        // Try to load and parse XML
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.Load(v_sFilePath);

        // Check if root element exists
        if (xmlDoc.DocumentElement == null)
        {
          v_oJobProtocol.WriteLine($"Warning: XML has no root element: {v_sFilePath}");
          return false;
        }

        return true;
      }
      catch (XmlException xmlEx)
      {
        v_oJobProtocol.WriteLine($"Warning: Invalid XML structure in {v_sFilePath}: {xmlEx.Message}");
        return false;
      }
      catch (Exception ex)
      {
        v_oJobProtocol.WriteLine($"Warning: Error validating XML file {v_sFilePath}: {ex.Message}");
        return false;
      }
    }

    String FindExistingDownloadedFile(String v_sTradeRegisterNumber, int v_iFileNumber, ClJobProtocol v_oJobProtocol)
    {
      try
      {
        // Pattern: *_HRB_30217+SI*.xml or *_HRB_30217_*_SI*.xml
        // Examples: 
        // - SN-Chemnitz_HRB_30217+SI-20260207103112.xml
        // - BR-Potsdam_HRB_40192+SI-20260118093803.xml
        
        string searchPattern = $"*HRB_{v_sTradeRegisterNumber}*SI*.xml";
        string[] files = Directory.GetFiles(m_sBrowserDownloadFolder, searchPattern);

        if (files.Length == 0)
          return null;

        // Sort files by creation time and return the one at the position (v_iFileNumber - 1)
        // Assuming files are ordered chronologically as they were downloaded
        var sortedFiles = files.OrderBy(f => File.GetCreationTime(f)).ToArray();

        if (sortedFiles.Length >= v_iFileNumber)
        {
          string candidateFile = sortedFiles[v_iFileNumber - 1];
          
          // Validate XML structure
          if (IsValidXmlFile(candidateFile, v_oJobProtocol))
          {
            return candidateFile;
          }
          else
          {
            v_oJobProtocol.WriteLine($"Found existing file but XML validation failed, will re-download: {candidateFile}");
            // Delete corrupted file
            try
            {
              File.Delete(candidateFile);
              v_oJobProtocol.WriteLine($"Deleted corrupted file: {candidateFile}");
            }
            catch (Exception exDel)
            {
              v_oJobProtocol.WriteLine($"Warning: Could not delete corrupted file: {exDel.Message}");
            }
            return null;
          }
        }

        return null;
      }
      catch (Exception ex)
      {
        v_oJobProtocol.WriteLine($"Warning: Error checking for existing file: {ex.Message}");
        return null;
      }
    }

    void SearchByRegisterNumberAndDownloadResults_SI(
      WebDriver v_oBrowserWrapper,
      ClJobProtocol v_oJobProtocol,
      String v_sTradeRegisterNumber,
      out Boolean v_bErrorOfAType_At_TradeRegisterPortal,
      out Boolean v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
    {
      v_bErrorOfAType_At_TradeRegisterPortal = false;
      v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = false;

      String sTrageRegisterType = "HRB";
      String sPrefixStringForProtocol = sTrageRegisterType + " " + v_sTradeRegisterNumber + " ";

      DateTime oStartTime = DateTime.Now;

      v_oJobProtocol.WriteLine(
        sPrefixStringForProtocol +
        "Start time: " + oStartTime.ToString("yyyy.MM.dd HH:mm:ss"));


      // Clear Statistics.
      ClearStatistics_before_SearchAtDEUTradeRegisterWebSite();


      // Navigate browser to Search Page of German Trade Register.
      try
      {
        v_oBrowserWrapper.Navigate().GoToUrl(m_sTradeRegisterSeachPageUrl);
        Thread.Sleep(500);
        
        // Check if page loaded successfully or shows "can't be reached"
        string pageTitle = v_oBrowserWrapper.Title.ToLower();
        string pageSource = v_oBrowserWrapper.PageSource.ToLower();
        
        if (pageTitle.Contains("site can't be reached") || 
            pageTitle.Contains("this site can't be reached") ||
            pageSource.Contains("err_connection_refused") ||
            pageSource.Contains("err_connection_timed_out") ||
            pageSource.Contains("err_name_not_resolved"))
        {
          m_iNetworkErrorCounter++;
          int waitSeconds = Math.Min(30 * m_iNetworkErrorCounter, 300); // Max 5 minutes
          
          v_oJobProtocol.WriteLine(sPrefixStringForProtocol + 
            "ERROR: Site can't be reached. Network error counter: " + m_iNetworkErrorCounter + 
            ". Waiting " + waitSeconds + " seconds before retry...");
          
          Thread.Sleep(waitSeconds * 1000);
          
          v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = true;
          return;
        }
        
        // Reset counter on successful connection
        if (m_iNetworkErrorCounter > 0)
        {
          v_oJobProtocol.WriteLine(sPrefixStringForProtocol + 
            "Connection successful. Resetting network error counter (was: " + m_iNetworkErrorCounter + ").");
          m_iNetworkErrorCounter = 0;
        }
      }
      catch (Exception ex)
      {
        m_iNetworkErrorCounter++;
        int waitSeconds = Math.Min(30 * m_iNetworkErrorCounter, 300); // Max 5 minutes
        
        v_oJobProtocol.WriteLine(sPrefixStringForProtocol + 
          "Network error during navigation: " + ex.Message + 
          ". Error counter: " + m_iNetworkErrorCounter + 
          ". Waiting " + waitSeconds + " seconds before retry...");
        
        Thread.Sleep(waitSeconds * 1000);
        
        v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = true;
        return;
      }

      // Dismiss cookie consent banner before starting interaction
      DismissCookieBanner(v_oBrowserWrapper, v_oJobProtocol, sPrefixStringForProtocol);

      // Wait for page to stabilize after cookie dismissal
      Thread.Sleep(2500);

      IWebElement oBodyElement = v_oBrowserWrapper.FindElement(By.TagName("body"));

      IWebElement oRegisterTypeComboBox = null;
      IWebElement oRegisterNumberTextBox = null;
      IWebElement oRecordCountPerPageComboBox = null;
      IWebElement oSearchButton = null;


      // Find "form:registerArt_focus"
      oRegisterTypeComboBox = TryFindElementAtWebPageById(v_oBrowserWrapper, "form:registerArt_focus");
      if (oRegisterTypeComboBox == null)
        throw new Exception("Cannot find combo-box for Register Type (id='form:registerArt_focus').");

      // Find "form:registerNummer"
      oRegisterNumberTextBox = TryFindElementAtWebPageById(v_oBrowserWrapper, "form:registerNummer");
      if (oRegisterNumberTextBox == null)
        throw new Exception("Cannot find text box for Register Number (id='form:registerNummer').");

      // Find "form:ergebnisseProSeite_focus"
      oRecordCountPerPageComboBox = TryFindElementAtWebPageById(v_oBrowserWrapper, "form:ergebnisseProSeite_focus");
      if (oRecordCountPerPageComboBox == null)
        throw new Exception("Cannot find combo-box for Record Count per Page (id='form:ergebnisseProSeite_focus').");

      // Find "form:btnSuche"
      oSearchButton = TryFindElementAtWebPageById(v_oBrowserWrapper, "form:btnSuche");
      if (oSearchButton == null)
        throw new Exception("Cannot find button 'Suchen' (id='form:btnSuche').");


      // Move to "Type of Register" combo-box.
      for (int i = 0; i < 44; i++)
      {
        oBodyElement.SendKeys(Keys.Tab);
      }

      try
      {
        WaitUntilWebElementIsActive(v_oBrowserWrapper, oRegisterTypeComboBox, 30);
      }
      catch (Exception ex)
      {
        v_oJobProtocol.WriteLine(sPrefixStringForProtocol + "Warning: Could not activate register type combo: " + ex.Message);
        // Continue anyway - element might still work
      }

      oRegisterTypeComboBox.SendKeys(Keys.ArrowDown);
      Thread.Sleep(100);
      oRegisterTypeComboBox.SendKeys(Keys.ArrowDown);
      Thread.Sleep(300);


      // Move to "Register number" text-box
      oBodyElement.SendKeys(Keys.Tab);
      
      try
      {
        WaitUntilWebElementIsActive(v_oBrowserWrapper, oRegisterNumberTextBox, 10);
      }
      catch (Exception ex)
      {
        v_oJobProtocol.WriteLine(sPrefixStringForProtocol + "Warning: Could not activate register number field: " + ex.Message);
        // Continue anyway
      }

      oRegisterNumberTextBox.SendKeys(v_sTradeRegisterNumber);


      // Move to "Record Count per Page" combo-box 
      for (int i = 0; i < 7; i++)
      {
        oBodyElement.SendKeys(Keys.Tab);
      }

      try
      {
        WaitUntilWebElementIsActive(v_oBrowserWrapper, oRecordCountPerPageComboBox, 10);
      }
      catch (Exception ex)
      {
        v_oJobProtocol.WriteLine(sPrefixStringForProtocol + "Warning: Could not activate record count combo: " + ex.Message);
        // Continue anyway
      }

      oRecordCountPerPageComboBox.SendKeys(Keys.ArrowDown);
      Thread.Sleep(100);
      oRecordCountPerPageComboBox.SendKeys(Keys.ArrowDown);
      Thread.Sleep(100);
      oRecordCountPerPageComboBox.SendKeys(Keys.ArrowDown);
      Thread.Sleep(300);


      // Press "Suchen" button.
      oSearchButton.SendKeys(Keys.Enter);

      // Wait while Search Page is replaced by Results Page.
      {
        Boolean bSearchPageWasReplacedByANewPage = false;

        for (int i = 0; i < 100; i++)
        {
          IWebElement oTempSearchButton = TryFindElementAtWebPageById(v_oBrowserWrapper, "form:btnSuche");

          if (oTempSearchButton == null)
          {
            bSearchPageWasReplacedByANewPage = true;

            break;
          }

          Thread.Sleep(100);
        }

        if (!bSearchPageWasReplacedByANewPage)
        {
          // Instead of throwing exception, treat this as session expiration
          // This will trigger browser restart and retry
          v_oJobProtocol.WriteLine(sPrefixStringForProtocol + 
            "Warning: Timeout waiting for results page. Treating as session issue.");
          v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = true;
          return;
        }

      }

      // Check if one of special error pages was shown
      //   by web server of DEU Trade Register portal.
      // 
      // <h1>Expired session:</h1>
      // <p>Error ID: 440</p>
      // Also check for Errorcode: 0
      {
        String sHTMLPageSource = v_oBrowserWrapper.PageSource;

        if (sHTMLPageSource.IndexOf("<p>Error ID: 440</p>") != -1 || 
            sHTMLPageSource.IndexOf("Errorcode: 0") != -1)
        {
          v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = true;
        }
      }

      if (!v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
      {
        // // Start monitoring of HTTP Responces.
        // StartMonitoringOfHttpResponces(v_oBrowserWrapper);

        // Download all SI links (.xml files) from result page.
        DownloadFilesFromTableOfSearchResults(
          v_oBrowserWrapper,
          v_oJobProtocol,
          sPrefixStringForProtocol,
          v_sTradeRegisterNumber,
          out v_bErrorOfAType_At_TradeRegisterPortal,
          out v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440);


        // // Stop monitoring of HTTP Responces.
        // StopMonitoringOfHttpResponces(v_oBrowserWrapper);

        if (v_bErrorOfAType_At_TradeRegisterPortal || 
            v_bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
        {
          // Print Statistics
          PrintStatistics(v_oJobProtocol, sPrefixStringForProtocol);

          v_oJobProtocol.WriteLine(
            sPrefixStringForProtocol +
            "Register number \"" + v_sTradeRegisterNumber + "\" downloaded has finished with error!");
        }
        else
        {
          // Print Statistics
          PrintStatistics(v_oJobProtocol, sPrefixStringForProtocol);

          v_oJobProtocol.WriteLine(
            sPrefixStringForProtocol +
            "Register number \"" + v_sTradeRegisterNumber + "\" was downloaded successfully.");
        }
      }
      else
      {
        v_oJobProtocol.WriteLine(
          sPrefixStringForProtocol +
          "Step 'Search for companies by Register Number'. " +
          "An error at DEU Trade Register web server has occured. " +
          "\"Error ID: 440 (Expired session.)\" " +
          "Register number \"" + v_sTradeRegisterNumber + "\" was NOT downloaded. " +
          "I will try to restart web browser " + 
          "and repeat search for this register number again.");
      }


      DateTime oEndTime = DateTime.Now;
      TimeSpan oWorkTimeSpan = oEndTime - oStartTime;

      v_oJobProtocol.WriteLine(
        sPrefixStringForProtocol +
        "End time: " + oEndTime.ToString("yyyy.MM.dd HH:mm:ss"));

      v_oJobProtocol.WriteLine(
        sPrefixStringForProtocol +
        "Processing time span: " + oWorkTimeSpan.ToString());

      v_oJobProtocol.WriteLine();
      v_oJobProtocol.WriteLine();
    }

    void PrintFieldsOfWebElement(
      ClJobProtocol v_oJobProtocol,
      IWebElement v_oWebElement,
      String v_sElementNameForLog)
    {
      v_oJobProtocol.WriteLine("Fields of '" + v_sElementNameForLog + "':");
      v_oJobProtocol.WriteLine("  Tag name................: \"" + v_oWebElement.TagName + "\"");
      v_oJobProtocol.WriteLine("  Id......................: \"" + v_oWebElement.GetAttribute("id") + "\"");
      v_oJobProtocol.WriteLine("  Class...................: \"" + v_oWebElement.GetAttribute("class") + "\"");
      v_oJobProtocol.WriteLine("  IWebElement.Displayed...: \"" + v_oWebElement.Displayed.ToString() + "\"");
      v_oJobProtocol.WriteLine("  IWebElement.Selected....: \"" + v_oWebElement.Selected.ToString() + "\"");
      v_oJobProtocol.WriteLine("  IWebElement.Text........: \"" + v_oWebElement.Text + "\"");
      v_oJobProtocol.WriteLine("");
    }


    public ChromeDriver CreateWebBrowserObject_and_SetSettings(
      ClJobProtocol v_oJobProtocol)
    {
      // Print information to log file.
      DateTime oStartTime = DateTime.Now;

      v_oJobProtocol.WriteLine(
        "Start creating of a new Web Browser process: " + oStartTime.ToString("yyyy.MM.dd HH:mm:ss"));


      // Create and fill options for Google Chrome browser.
      ChromeOptions oChromeOptions = new ChromeOptions();

      // Rotate user agent to avoid detection
      string selectedUserAgent = UserAgents[m_random.Next(UserAgents.Length)];
      oChromeOptions.AddArgument($"--user-agent={selectedUserAgent}");
      v_oJobProtocol.WriteLine("Using User-Agent: " + selectedUserAgent);

      // Configure proxy if available
      if (m_proxyList != null && m_proxyList.Count > 0)
      {
        string currentProxy = m_proxyList[m_currentProxyIndex];
        
        // Parse proxy format: [user:pass@]host:port
        string proxyServer = currentProxy;
        
        if (currentProxy.Contains("@"))
        {
          // Proxy with authentication
          string[] parts = currentProxy.Split('@');
          if (parts.Length == 2)
          {
            string credentials = parts[0]; // user:pass
            proxyServer = parts[1]; // host:port
            
            // Chrome doesn't support proxy auth via command line directly
            // User will need to handle authentication separately or use extension
            v_oJobProtocol.WriteLine("Warning: Proxy with authentication detected. Chrome may require additional setup for auth.");
          }
        }
        
        oChromeOptions.AddArgument($"--proxy-server={proxyServer}");
        v_oJobProtocol.WriteLine($"Using Proxy [{m_currentProxyIndex + 1}/{m_proxyList.Count}]: {currentProxy}");
      }
      else
      {
        v_oJobProtocol.WriteLine("No proxies configured. Using direct connection.");
      }

      // Suppress erors from Selenium driver for Google Chrome.
      //
      // For example:
      //   "[7252:1340:1001/193319.654:ERROR:components\device_event_log\device_event_log_impl.cc:198] [19:33:19.657] FIDO: webauthn_api.cc:121 Windows WebAuthn API failed to load"
      //   "[7252:5376:1001/193323.313:ERROR:google_apis\gcm\engine\registration_request.cc:291] Registration response error message: PHONE_REGISTRATION_ERROR"
      //   "[7252:5376:1001/193323.314:ERROR:google_apis\gcm\engine\mcs_client.cc:702] Failed to log in to GCM, resetting connection."
      //   "[7252:1340:1001/193326.383:ERROR:components\page_load_metrics\browser\page_load_metrics_update_dispatcher.cc:88] Invalid response_start 0.112 s for parse_start 0.11 s"
      oChromeOptions.AddArgument("--log-level=3");

      // Configure automatic downloads without "Save As" dialog
      // Set the default download directory
      oChromeOptions.AddUserProfilePreference("download.default_directory", m_sBrowserDownloadFolder);
      
      // Disable the "Ask where to save each file before downloading" prompt
      oChromeOptions.AddUserProfilePreference("download.prompt_for_download", false);
      
      // Automatically download files to the specified directory
      oChromeOptions.AddUserProfilePreference("download.directory_upgrade", true);
      
      // Enable safe browsing to avoid additional download warnings
      oChromeOptions.AddUserProfilePreference("safebrowsing.enabled", true);
      
      // Disable PDF viewer to force download instead of opening in browser
      oChromeOptions.AddUserProfilePreference("plugins.always_open_pdf_externally", true);

      v_oJobProtocol.WriteLine("Configured Chrome download folder: " + m_sBrowserDownloadFolder);


      // Create process of Web Browser
      ChromeDriver oBrowserWrapper = null;
      
      try
      {
        // Create ChromeDriverService to let Selenium Manager handle driver automatically
        // This ensures the correct ChromeDriver version is downloaded for the installed Chrome
        v_oJobProtocol.WriteLine("Starting Chrome browser (Selenium Manager will auto-download matching ChromeDriver if needed)...");
        
        ChromeDriverService service = ChromeDriverService.CreateDefaultService();
        service.SuppressInitialDiagnosticInformation = true;
        
        oBrowserWrapper = new ChromeDriver(service, oChromeOptions);
        v_oJobProtocol.WriteLine("Chrome browser started successfully.");
      }
      catch (Exception ex)
      {
        v_oJobProtocol.WriteLine("ERROR: Failed to start Chrome browser.");
        v_oJobProtocol.WriteLine("Error message: " + ex.Message);
        v_oJobProtocol.WriteLine();
        v_oJobProtocol.WriteLine("POSSIBLE SOLUTIONS:");
        v_oJobProtocol.WriteLine("1. Delete cached ChromeDriver: Delete folder %USERPROFILE%\\.cache\\selenium");
        v_oJobProtocol.WriteLine("2. Install Google Chrome browser from: https://www.google.com/chrome/");
        v_oJobProtocol.WriteLine("3. Ensure you have internet access (Selenium Manager needs to download ChromeDriver)");
        v_oJobProtocol.WriteLine("4. Check that Chrome is not blocked by antivirus or firewall");
        v_oJobProtocol.WriteLine();
        throw new Exception("Cannot start Chrome browser. Try deleting %USERPROFILE%\\.cache\\selenium folder and run again.", ex);
      }

      Point oBrowserWindowPosition = new Point(0, 0);
      Size oBrowserWindowsSize = new Size(1200, 700);

      oBrowserWrapper.Manage().Window.Position = oBrowserWindowPosition;
      oBrowserWrapper.Manage().Window.Size = oBrowserWindowsSize;


      // Print information to log file.
      DateTime oEndTime = DateTime.Now;
      TimeSpan oWorkTimeSpan = oEndTime - oStartTime;

      v_oJobProtocol.WriteLine(
        "End creating of a new Web Browser process: " + oEndTime.ToString("yyyy.MM.dd HH:mm:ss"));

      v_oJobProtocol.WriteLine(
        "Creating time span: " + oWorkTimeSpan.ToString());

      v_oJobProtocol.WriteLine();
      v_oJobProtocol.WriteLine();


      return oBrowserWrapper;
    }

    void RotateProxy(ClJobProtocol v_oJobProtocol)
    {
      if (m_proxyList != null && m_proxyList.Count > 0)
      {
        m_currentProxyIndex = (m_currentProxyIndex + 1) % m_proxyList.Count;
        v_oJobProtocol.WriteLine($"Rotated to next proxy. Will use proxy [{m_currentProxyIndex + 1}/{m_proxyList.Count}] on next browser restart.");
      }
    }

    public void ExtractData_RangeMode(
      ClJobProtocol v_oJobProtocol,
      String v_sStartRegisterNumber,
      String v_sEndRegisterNumber)
    {
      DateTime oStartTime = DateTime.Now;

      v_oJobProtocol.WriteLine(
        "Processing start time: " + oStartTime.ToString("yyyy.MM.dd HH:mm:ss"));


      uint uiStartRegisterNumber = UInt32.Parse(v_sStartRegisterNumber);
      uint uiEndRegisterNumber = UInt32.Parse(v_sEndRegisterNumber);


      // oBrowserWrapper.Navigate().GoToUrl(sTradeRegisterResultTablePageMore100HitsUrl);
      // bool bMore100HitsState = CheckWhetherMore100ResultsWereFound(oBrowserWrapper);


      ChromeDriver oBrowserWrapper = null;

      oBrowserWrapper = CreateWebBrowserObject_and_SetSettings(v_oJobProtocol);


      v_oJobProtocol.WriteLine(
        "Start processing of Trade Register Number range (" +
        uiStartRegisterNumber.ToString() + " - " + uiEndRegisterNumber.ToString() + ").");
      v_oJobProtocol.WriteLine();


      // Download data
      uint uiSequantialErrorCounter = 0;
      uint uiSuccessfullyDownloadedTradeNumbers_SequentialCounter = 0;

      uint uiRegisterNumber = uiStartRegisterNumber;

      while( uiRegisterNumber <= uiEndRegisterNumber)
      {
        String sTradeRegisterNumber = uiRegisterNumber.ToString();

        Boolean bErrorOfAType_At_TradeRegisterPortal = false;
        Boolean bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = false;


        SearchByRegisterNumberAndDownloadResults_SI(
          oBrowserWrapper,
          v_oJobProtocol,
          sTradeRegisterNumber,
          out bErrorOfAType_At_TradeRegisterPortal,
          out bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440);


        if (bErrorOfAType_At_TradeRegisterPortal ||
            bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
        {
          uiSequantialErrorCounter++;

          v_oJobProtocol.WriteLine(
            "Warning: Sequential error count at DEU TradeRegister web site: " +
            uiSequantialErrorCounter.ToString() + " errors. Restarting browser and continuing...");

          // Restart Google Chrome browser.
          try
          {
            if (oBrowserWrapper != null)
            {
              oBrowserWrapper.Quit();
            }
          }
          catch (Exception exQuit)
          {
            v_oJobProtocol.WriteLine("Warning: Error closing old browser: " + exQuit.Message);
          }
          finally
          {
            oBrowserWrapper = null;
          }

          if (bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
          {
            v_oJobProtocol.WriteLine(
              "Restart Web Browser " +
              "because of expiration of session at Trade Register Web Server.");
          }
          else
          {
            v_oJobProtocol.WriteLine(
              "Restart Web Browser " +
              "because of \"Errorcode: A xxx\" errors at Trade Register Web Server.");
          }

          // Rotate to next proxy (RangeMode error restart)
          RotateProxy(v_oJobProtocol);

          oBrowserWrapper = CreateWebBrowserObject_and_SetSettings(v_oJobProtocol);

          Thread.Sleep(2000);
        }
        else
        {
          // Restart Google Chrome browser
          //   after each 5 successfully downloaded searches
          //   by Register Number.
          //
          // This approach allows to prevent
          //   the following error
          //   at DEU Trade Register web server:
          //   "Error ID: 440 (Expired session.)"
          {
            uiSuccessfullyDownloadedTradeNumbers_SequentialCounter++;

            if (uiSuccessfullyDownloadedTradeNumbers_SequentialCounter >= m_iBrowserRestartInterval)
            {
              // Restart Google Chrome browser.
              try
              {
                if (oBrowserWrapper != null)
                {
                  oBrowserWrapper.Quit();
                }
              }
              catch (Exception exQuit)
              {
                v_oJobProtocol.WriteLine("Warning: Error closing old browser: " + exQuit.Message);
              }
              finally
              {
                oBrowserWrapper = null;
              }

              v_oJobProtocol.WriteLine(
                "Restart Web Browser " +
                "after successful processing of " + m_iBrowserRestartInterval + " Numbers in Trade Register.");

              v_oJobProtocol.WriteLine(
                "This restart should improve stability " +
                "of process of extracion of data.");

              // Rotate to next proxy (RangeMode successful restart)
              RotateProxy(v_oJobProtocol);

              oBrowserWrapper = CreateWebBrowserObject_and_SetSettings(v_oJobProtocol);

              Thread.Sleep(2000);

              uiSuccessfullyDownloadedTradeNumbers_SequentialCounter = 0;
            }
          }


          // Do needed actions if Download was successfull.
          uiRegisterNumber++;


          DEU_TR_ProgramSettings.Update_StartRegisterNumber_InSettingsFile(
            uiRegisterNumber.ToString());


          uiSequantialErrorCounter = 0;
        }
      }


      /*
      String sPrevCheckboxCheckedValue = oBadenWuerttembergCheckBox.GetAttribute("checked");

      oBadenWuerttembergCheckBox.Click();

      String sNewCheckboxCheckedValue = oBadenWuerttembergCheckBox.GetAttribute("checked");
      */


      // oBrowserWrapper.Navigate().GoToUrl(sTradeRegisterSeachPageUrl);

      // Console.WriteLine("Title of current Web Page: \"" + oBrowserWrapper.Title + "\".");
      // Console.WriteLine();

      // // Get HTML source code of current page.
      //
      // String sHTMLPageSource = oBrowserWrapper.PageSource;
      // Console.WriteLine("Source code of current Web Page: \"" + sHTMLPageSource + "\".");
      // Console.WriteLine();


      // String sPrevComboValue = oRegisterTypeComboBox.GetAttribute("value");

      // String sHTMLPageSource = oBrowserWrapper.PageSource;

      // File.WriteAllText(Path.Combine(sLogFolder, "PageSource.txt"), sHTMLPageSource);

      // String sActiveWindowTitle = ImportsFrom_User32_dll.GetActiveWindowTitle();
      // oJobProtocol.WriteLine("Title of active window: \"" + sActiveWindowTitle + "\"");

      // IWebElement oDivForULRegisterArt = oULRegisterArt.FindElement(By.XPath("parent::node()"));

      // IWebElement oActiveElement5 = oBrowserWrapper.SwitchTo().ActiveElement();
      // PrintFieldsOfWebElement(oJobProtocol, oActiveElement5, "ActiveElement5");



      // Close Web Browser.
      try
      {
        if (oBrowserWrapper != null)
        {
          oBrowserWrapper.Quit();
        }
      }
      catch (Exception ex)
      {
        v_oJobProtocol.WriteLine("Warning: Failed to close browser during final cleanup: " + ex.Message);
      }

      // Close CSV writer
      if (m_csvWriter != null)
      {
        m_csvWriter.Close();
        m_csvWriter = null;
        v_oJobProtocol.WriteLine("CSV output file closed.");
      }


      // Print statistics to .log file.
      DateTime oEndTime = DateTime.Now;
      TimeSpan oWorkTimeSpan = oEndTime - oStartTime;

      v_oJobProtocol.WriteLine(
        "Processing of Trade Register Number range (" + 
        uiStartRegisterNumber.ToString() + " - " + uiEndRegisterNumber.ToString() + ") " +
        "was completed successfully.");

      v_oJobProtocol.WriteLine(
        "Processing end time: " + oEndTime.ToString("yyyy.MM.dd HH:mm:ss"));

      v_oJobProtocol.WriteLine(
        "Processing time span: " + oWorkTimeSpan.ToString());

      v_oJobProtocol.WriteLine();
      v_oJobProtocol.WriteLine();
    }


    public void ExtractData_ListMode(
      ClJobProtocol v_oJobProtocol,
      List<String> v_oRegisterNumberList )
    {
      DateTime oStartTime = DateTime.Now;

      v_oJobProtocol.WriteLine(
        "Processing start time: " + oStartTime.ToString("yyyy.MM.dd HH:mm:ss"));


      // oBrowserWrapper.Navigate().GoToUrl(sTradeRegisterResultTablePageMore100HitsUrl);
      // bool bMore100HitsState = CheckWhetherMore100ResultsWereFound(oBrowserWrapper);


      ChromeDriver oBrowserWrapper = null;

      oBrowserWrapper = CreateWebBrowserObject_and_SetSettings(v_oJobProtocol);


      v_oJobProtocol.WriteLine(
        "Start processing of Trade Register Number list (" +
        v_oRegisterNumberList.Count().ToString() + " items).");
      v_oJobProtocol.WriteLine();


      // Download data
      uint uiSequantialErrorCounter = 0;
      uint uiSuccessfullyDownloadedTradeNumbers_SequentialCounter = 0;

      // uint uiRegisterNumber = uiStartRegisterNumber;
      int iCurrentListPosition = 0;

      while (iCurrentListPosition < v_oRegisterNumberList.Count())
      {
        String sTradeRegisterNumber = v_oRegisterNumberList[iCurrentListPosition];

        Boolean bErrorOfAType_At_TradeRegisterPortal = false;
        Boolean bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440 = false;


        SearchByRegisterNumberAndDownloadResults_SI(
          oBrowserWrapper,
          v_oJobProtocol,
          sTradeRegisterNumber,
          out bErrorOfAType_At_TradeRegisterPortal,
          out bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440);


        if (bErrorOfAType_At_TradeRegisterPortal ||
            bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
        {
          uiSequantialErrorCounter++;

          v_oJobProtocol.WriteLine(
            "Warning: Sequential error count at DEU TradeRegister web site: " +
            uiSequantialErrorCounter.ToString() + " errors. Restarting browser and continuing...");

          // Restart Google Chrome browser.
          try
          {
            if (oBrowserWrapper != null)
            {
              oBrowserWrapper.Quit();
            }
          }
          catch (Exception exQuit)
          {
            v_oJobProtocol.WriteLine("Warning: Error closing old browser: " + exQuit.Message);
          }
          finally
          {
            oBrowserWrapper = null;
          }

          if (bDEU_TradeRegister_WebSiteError__ExpiredSession_ErrorID_440)
          {
            v_oJobProtocol.WriteLine(
              "Restart Web Browser " +
              "because of expiration of session at Trade Register Web Server.");
          }
          else
          {
            v_oJobProtocol.WriteLine(
              "Restart Web Browser " +
              "because of \"Errorcode: A xxx\" errors at Trade Register Web Server.");
          }

          // Rotate to next proxy (ListMode error restart)
          RotateProxy(v_oJobProtocol);

          oBrowserWrapper = CreateWebBrowserObject_and_SetSettings(v_oJobProtocol);

          Thread.Sleep(2000);
        }
        else
        {
          // Restart Google Chrome browser
          //   after each 5 successfully downloaded searches
          //   by Register Number.
          //
          // This approach allows to prevent
          //   the following error
          //   at DEU Trade Register web server:
          //   "Error ID: 440 (Expired session.)"
          {
            uiSuccessfullyDownloadedTradeNumbers_SequentialCounter++;

            if (uiSuccessfullyDownloadedTradeNumbers_SequentialCounter >= m_iBrowserRestartInterval)
            {
              // Restart Google Chrome browser.
              try
              {
                if (oBrowserWrapper != null)
                {
                  oBrowserWrapper.Quit();
                }
              }
              catch (Exception exQuit)
              {
                v_oJobProtocol.WriteLine("Warning: Error closing old browser: " + exQuit.Message);
              }
              finally
              {
                oBrowserWrapper = null;
              }

              v_oJobProtocol.WriteLine(
                "Restart Web Browser " +
                "after successful processing of " + m_iBrowserRestartInterval + " Numbers in Trade Register.");

              v_oJobProtocol.WriteLine(
                "This restart should improve stability " +
                "of process of extracion of data.");

              // Rotate to next proxy (ListMode successful restart)
              RotateProxy(v_oJobProtocol);

              oBrowserWrapper = CreateWebBrowserObject_and_SetSettings(v_oJobProtocol);

              Thread.Sleep(2000);

              uiSuccessfullyDownloadedTradeNumbers_SequentialCounter = 0;
            }
          }


          // Do needed actions if Download was successfull.
          iCurrentListPosition++;


          DEU_TR_ProgramSettings.InsertOrUpdate_UnsuccessfullyDownloadedRegisterNumber__in__ListMode_Settings(
            sTradeRegisterNumber, true, "", 0, 0 );


          uiSequantialErrorCounter = 0;
        }
      }


      // Close Web Browser.
      try
      {
        if (oBrowserWrapper != null)
        {
          oBrowserWrapper.Quit();
        }
      }
      catch (Exception ex)
      {
        v_oJobProtocol.WriteLine("Warning: Failed to close browser during final cleanup: " + ex.Message);
      }

      // Close CSV writer
      if (m_csvWriter != null)
      {
        m_csvWriter.Close();
        m_csvWriter = null;
        v_oJobProtocol.WriteLine("CSV output file closed.");
      }


      // Print statistics to .log file.
      DateTime oEndTime = DateTime.Now;
      TimeSpan oWorkTimeSpan = oEndTime - oStartTime;

      v_oJobProtocol.WriteLine(
        "Processing of Trade Register Number list (" +
        v_oRegisterNumberList.Count().ToString() + " items) " +
        "was completed successfully.");

      v_oJobProtocol.WriteLine(
        "Processing end time: " + oEndTime.ToString("yyyy.MM.dd HH:mm:ss"));

      v_oJobProtocol.WriteLine(
        "Processing time span: " + oWorkTimeSpan.ToString());

      v_oJobProtocol.WriteLine();
      v_oJobProtocol.WriteLine();
    }

    public void ExtractData_TableScrapingMode(
      ClJobProtocol v_oJobProtocol,
      String v_sStartRegisterNumber,
      String v_sEndRegisterNumber)
    {
      DateTime oStartTime = DateTime.Now;

      v_oJobProtocol.WriteLine("Processing mode: Table Scraping (no file downloads)");
      v_oJobProtocol.WriteLine(
        "Processing start time: " + oStartTime.ToString("yyyy.MM.dd HH:mm:ss"));
      v_oJobProtocol.WriteLine();

      uint uiStartRegisterNumber = 0;
      uint uiEndRegisterNumber = 0;

      uiStartRegisterNumber = UInt32.Parse(v_sStartRegisterNumber);
      uiEndRegisterNumber = UInt32.Parse(v_sEndRegisterNumber);

      // Create CSV writer with table scraping columns (append mode to preserve existing data)
      bool fileExists = File.Exists(m_sOutputCSVFile) && new FileInfo(m_sOutputCSVFile).Length > 0;
      StreamingCSVWriter tableCsvWriter = new StreamingCSVWriter(m_sOutputCSVFile, ';', '"', fileExists);
      
      // Always set column names (needed for validation even when appending)
      if (!fileExists)
      {
        tableCsvWriter.WriteHeader("Region", "District court", "HR number", "CompanyName", "City", "Status");
      }
      else
      {
        // When appending, still need to set column names for validation
        tableCsvWriter.SetColumnNames("Region", "District court", "HR number", "CompanyName", "City", "Status");
      }

      WebDriver oBrowserWrapper = null;

      try
      {
        oBrowserWrapper = CreateWebBrowserObject_and_SetSettings(v_oJobProtocol);

        uint uiCurrentRegisterNumber = uiStartRegisterNumber;
        uint uiSuccessfullyDownloadedTradeNumbers_SequentialCounter = 0;

        while (uiCurrentRegisterNumber <= uiEndRegisterNumber)
        {
          String sTradeRegisterNumber = uiCurrentRegisterNumber.ToString();

          v_oJobProtocol.WriteLine("Processing HR number: " + sTradeRegisterNumber);

          Boolean bSuccess = SearchAndScrapeTableData(
            oBrowserWrapper,
            v_oJobProtocol,
            sTradeRegisterNumber,
            tableCsvWriter);

          if (bSuccess)
          {
            v_oJobProtocol.WriteLine("Table data scraped successfully for HR " + sTradeRegisterNumber);
            tableCsvWriter.Flush();  // Flush after each page
            uiSuccessfullyDownloadedTradeNumbers_SequentialCounter++;
          }
          else
          {
            v_oJobProtocol.WriteLine("No data found or error for HR " + sTradeRegisterNumber);
          }

          uiCurrentRegisterNumber++;

          // Update StartRegisterNumber in settings file to resume from next number
          DEU_TR_ProgramSettings.Update_StartRegisterNumber_InSettingsFile(
            uiCurrentRegisterNumber.ToString());

          // Browser restart logic
          if (uiSuccessfullyDownloadedTradeNumbers_SequentialCounter >= m_iBrowserRestartInterval)
          {
            v_oJobProtocol.WriteLine(
              "Browser restart interval reached (" + m_iBrowserRestartInterval.ToString() + " searches). " +
              "Restarting browser...");

            try
            {
              if (oBrowserWrapper != null)
              {
                oBrowserWrapper.Quit();
              }
            }
            catch (Exception ex)
            {
              v_oJobProtocol.WriteLine("Warning: Error closing browser: " + ex.Message);
            }
            finally
            {
              oBrowserWrapper = null;
            }

            RotateProxy(v_oJobProtocol);
            oBrowserWrapper = CreateWebBrowserObject_and_SetSettings(v_oJobProtocol);
            Thread.Sleep(2000);
            uiSuccessfullyDownloadedTradeNumbers_SequentialCounter = 0;
          }
        }
      }
      finally
      {
        // Close Web Browser
        try
        {
          if (oBrowserWrapper != null)
          {
            oBrowserWrapper.Quit();
          }
        }
        catch (Exception ex)
        {
          v_oJobProtocol.WriteLine("Warning: Failed to close browser during final cleanup: " + ex.Message);
        }

        // Close CSV writer
        if (tableCsvWriter != null)
        {
          tableCsvWriter.Close();
          v_oJobProtocol.WriteLine("CSV output file closed.");
        }
      }

      DateTime oEndTime = DateTime.Now;
      TimeSpan oWorkTimeSpan = oEndTime - oStartTime;

      v_oJobProtocol.WriteLine(
        "Table scraping completed for range (" +
        uiStartRegisterNumber.ToString() + " - " + uiEndRegisterNumber.ToString() + ")");

      v_oJobProtocol.WriteLine(
        "Processing end time: " + oEndTime.ToString("yyyy.MM.dd HH:mm:ss"));

      v_oJobProtocol.WriteLine(
        "Processing time span: " + oWorkTimeSpan.ToString());

      v_oJobProtocol.WriteLine();
      v_oJobProtocol.WriteLine();
    }

    Boolean SearchAndScrapeTableData(
      WebDriver v_oBrowserWrapper,
      ClJobProtocol v_oJobProtocol,
      String v_sTradeRegisterNumber,
      StreamingCSVWriter v_csvWriter)
    {
      try
      {
        String sTrageRegisterType = "HRB";
        String sPrefixStringForProtocol = sTrageRegisterType + " " + v_sTradeRegisterNumber + " ";

        // Navigate browser to Search Page (identical to SearchByRegisterNumberAndDownloadResults_SI)
        try
        {
          v_oBrowserWrapper.Navigate().GoToUrl(m_sTradeRegisterSeachPageUrl);
          Thread.Sleep(500);
          
          // Check if page loaded successfully or shows "can't be reached"
          string pageTitle = v_oBrowserWrapper.Title.ToLower();
          string pageSource = v_oBrowserWrapper.PageSource.ToLower();
          
          if (pageTitle.Contains("site can't be reached") || 
              pageTitle.Contains("this site can't be reached") ||
              pageSource.Contains("err_connection_refused") ||
              pageSource.Contains("err_connection_timed_out") ||
              pageSource.Contains("err_name_not_resolved"))
          {
            m_iNetworkErrorCounter++;
            int waitSeconds = Math.Min(30 * m_iNetworkErrorCounter, 300); // Max 5 minutes
            
            v_oJobProtocol.WriteLine(sPrefixStringForProtocol + 
              "ERROR: Site can't be reached. Network error counter: " + m_iNetworkErrorCounter + 
              ". Waiting " + waitSeconds + " seconds before retry...");
            
            Thread.Sleep(waitSeconds * 1000);
            return false;
          }
          
          // Reset counter on successful connection
          if (m_iNetworkErrorCounter > 0)
          {
            v_oJobProtocol.WriteLine(sPrefixStringForProtocol + 
              "Connection successful. Resetting network error counter (was: " + m_iNetworkErrorCounter + ").");
            m_iNetworkErrorCounter = 0;
          }
        }
        catch (Exception ex)
        {
          m_iNetworkErrorCounter++;
          int waitSeconds = Math.Min(30 * m_iNetworkErrorCounter, 300); // Max 5 minutes
          
          v_oJobProtocol.WriteLine(sPrefixStringForProtocol + 
            "Network error during navigation: " + ex.Message + 
            ". Error counter: " + m_iNetworkErrorCounter + 
            ". Waiting " + waitSeconds + " seconds before retry...");
          
          Thread.Sleep(waitSeconds * 1000);
          return false;
        }

        // Dismiss cookie consent banner before starting interaction
        DismissCookieBanner(v_oBrowserWrapper, v_oJobProtocol, sPrefixStringForProtocol);

        // Wait for page to stabilize after cookie dismissal
        Thread.Sleep(2500);

        IWebElement oBodyElement = v_oBrowserWrapper.FindElement(By.TagName("body"));
        IWebElement oRegisterTypeComboBox = null;
        IWebElement oRegisterNumberTextBox = null;
        IWebElement oRecordCountPerPageComboBox = null;
        IWebElement oSearchButton = null;

        // Find elements
        oRegisterTypeComboBox = TryFindElementAtWebPageById(v_oBrowserWrapper, "form:registerArt_focus");
        if (oRegisterTypeComboBox == null)
          throw new Exception("Cannot find combo-box for Register Type (id='form:registerArt_focus').");

        oRegisterNumberTextBox = TryFindElementAtWebPageById(v_oBrowserWrapper, "form:registerNummer");
        if (oRegisterNumberTextBox == null)
          throw new Exception("Cannot find text box for Register Number (id='form:registerNummer').");

        oRecordCountPerPageComboBox = TryFindElementAtWebPageById(v_oBrowserWrapper, "form:ergebnisseProSeite_focus");
        if (oRecordCountPerPageComboBox == null)
          throw new Exception("Cannot find combo-box for Record Count per Page (id='form:ergebnisseProSeite_focus').");

        oSearchButton = TryFindElementAtWebPageById(v_oBrowserWrapper, "form:btnSuche");
        if (oSearchButton == null)
          throw new Exception("Cannot find button 'Suchen' (id='form:btnSuche').");

        // Move to "Type of Register" combo-box
        for (int i = 0; i < 44; i++)
        {
          oBodyElement.SendKeys(Keys.Tab);
        }

        try
        {
          WaitUntilWebElementIsActive(v_oBrowserWrapper, oRegisterTypeComboBox, 30);
        }
        catch (Exception ex)
        {
          v_oJobProtocol.WriteLine(sPrefixStringForProtocol + "Warning: Could not activate register type combo: " + ex.Message);
        }

        oRegisterTypeComboBox.SendKeys(Keys.ArrowDown);
        Thread.Sleep(100);
        oRegisterTypeComboBox.SendKeys(Keys.ArrowDown);
        Thread.Sleep(300);

        // Move to "Register number" text-box
        oBodyElement.SendKeys(Keys.Tab);
        
        try
        {
          WaitUntilWebElementIsActive(v_oBrowserWrapper, oRegisterNumberTextBox, 10);
        }
        catch (Exception ex)
        {
          v_oJobProtocol.WriteLine(sPrefixStringForProtocol + "Warning: Could not activate register number field: " + ex.Message);
        }

        oRegisterNumberTextBox.SendKeys(v_sTradeRegisterNumber);

        // Move to "Record Count per Page" combo-box 
        for (int i = 0; i < 7; i++)
        {
          oBodyElement.SendKeys(Keys.Tab);
        }

        try
        {
          WaitUntilWebElementIsActive(v_oBrowserWrapper, oRecordCountPerPageComboBox, 10);
        }
        catch (Exception ex)
        {
          v_oJobProtocol.WriteLine(sPrefixStringForProtocol + "Warning: Could not activate record count combo: " + ex.Message);
        }

        oRecordCountPerPageComboBox.SendKeys(Keys.ArrowDown);
        Thread.Sleep(100);
        oRecordCountPerPageComboBox.SendKeys(Keys.ArrowDown);
        Thread.Sleep(100);
        oRecordCountPerPageComboBox.SendKeys(Keys.ArrowDown);
        Thread.Sleep(300);

        // Press "Suchen" button
        oSearchButton.SendKeys(Keys.Enter);

        // Wait while Search Page is replaced by Results Page
        Boolean bSearchPageWasReplacedByANewPage = false;
        for (int i = 0; i < 100; i++)
        {
          IWebElement oTempSearchButton = TryFindElementAtWebPageById(v_oBrowserWrapper, "form:btnSuche");
          if (oTempSearchButton == null)
          {
            bSearchPageWasReplacedByANewPage = true;
            break;
          }
          Thread.Sleep(100);
        }

        if (!bSearchPageWasReplacedByANewPage)
        {
          v_oJobProtocol.WriteLine(sPrefixStringForProtocol + "Timeout waiting for results page.");
          return false;
        }

        // Check for error pages
        String sHTMLPageSource = v_oBrowserWrapper.PageSource;
        if (sHTMLPageSource.IndexOf("<p>Error ID: 440</p>") != -1 || 
            sHTMLPageSource.IndexOf("Errorcode: 0") != -1)
        {
          v_oJobProtocol.WriteLine(sPrefixStringForProtocol + "Session error detected (Error 440 or Errorcode 0).");
          return false;
        }

        // Find the results table
        IWebElement resultsTable = v_oBrowserWrapper.FindElement(
          By.XPath("//*[@id=\"ergebnissForm:selectedSuchErgebnisFormTable\"]/div[2]/table"));

        // Get all rows (skip header if needed)
        ReadOnlyCollection<IWebElement> rows = resultsTable.FindElements(By.TagName("tr"));

        int recordCount = 0;

        foreach (IWebElement row in rows)
        {
          try
          {
            // Look for the main info cell with class containing the court and HR number
            var infoCells = row.FindElements(By.XPath(".//td[@class='ui-panelgrid-cell fontTableNameSize']"));
            
            if (infoCells.Count == 0)
              continue;

            string infoText = infoCells[0].Text.Trim();
            
            if (string.IsNullOrEmpty(infoText))
              continue;

            // Parse: "Bavaria  District court Nürnberg HRB 30302"
            string region = "";
            string districtCourt = "";
            string hrNumber = "";

            // Extract HR number (e.g., "HRB 30302")
            var hrMatch = System.Text.RegularExpressions.Regex.Match(infoText, @"(HR[AB]\s+\d+)");
            if (hrMatch.Success)
            {
              hrNumber = hrMatch.Groups[1].Value.Trim();

              // Remove HR number from text to parse the rest
              string remainingText = infoText.Substring(0, hrMatch.Index).Trim();

              // Look for "District court" or "Amtsgericht"
              int districtCourtIndex = remainingText.IndexOf("District court");
              if (districtCourtIndex == -1)
                districtCourtIndex = remainingText.IndexOf("Amtsgericht");

              if (districtCourtIndex != -1)
              {
                region = remainingText.Substring(0, districtCourtIndex).Trim();
                districtCourt = remainingText.Substring(districtCourtIndex).Replace("District court", "").Replace("Amtsgericht", "").Trim();
              }
            }

            // Get company name (next row, span with marginLeft20 class)
            var companyNameElements = row.FindElements(By.XPath(".//span[@class='marginLeft20']"));
            string companyName = "";
            if (companyNameElements.Count > 0)
            {
              companyName = companyNameElements[0].Text.Trim();
            }

            // Get city (span with class "verticalText" in sitzSuchErgebnisse cell)
            var cityElements = row.FindElements(By.XPath(".//td[contains(@class, 'sitzSuchErgebnisse')]//span[@class='verticalText ']"));
            string city = "";
            if (cityElements.Count > 0)
            {
              city = cityElements[0].Text.Trim();
            }

            // Get status (another span with class "verticalText")
            var statusElements = row.FindElements(By.XPath(".//span[@class='verticalText']"));
            string status = "";
            // The second verticalText span is usually the status
            if (statusElements.Count > 1)
            {
              status = statusElements[1].Text.Trim();
            }
            else if (statusElements.Count == 1 && string.IsNullOrEmpty(city))
            {
              // If only one verticalText, check if it's status
              string text = statusElements[0].Text.Trim();
              if (text.Contains("registered") || text.Contains("deleted") || text.Contains("gelöscht"))
              {
                status = text;
              }
              else
              {
                city = text;
              }
            }

            // Write to CSV if we have valid data
            if (!string.IsNullOrEmpty(hrNumber) && !string.IsNullOrEmpty(companyName))
            {
              v_csvWriter.WriteRow(new string[] 
              { 
                region, 
                districtCourt, 
                hrNumber, 
                companyName, 
                city, 
                status 
              });

              recordCount++;
              v_oJobProtocol.WriteLine($"  Scraped: {region} | {districtCourt} | {hrNumber} | {companyName} | {city} | {status}");
            }
          }
          catch (Exception exRow)
          {
            // Skip problematic rows
            v_oJobProtocol.WriteLine("Warning: Error parsing row: " + exRow.Message);
          }
        }

        v_oJobProtocol.WriteLine($"Total records scraped from table: {recordCount}");
        return recordCount > 0;
      }
      catch (NoSuchElementException ex)
      {
        v_oJobProtocol.WriteLine("No results found or element not found: " + ex.Message);
        return false;
      }
      catch (Exception ex)
      {
        v_oJobProtocol.WriteLine("Error during table scraping: " + ex.Message);
        return false;
      }
    }

  }
}
