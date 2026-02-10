using APX;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.DevTools.V139.IndexedDB;
using System.Timers;
using System.Xml;

namespace DEU_TradeRegister_WebExtractor
{
  internal class Program
  {
    static String s_sLogFolder = "";

    static int Main(string[] args)
    {
      String sTradeRegisterSeachPageUrl = "";
      String sOutputFolder = "";
      String sBrowserDownloadFolder = "";
      String sOutputCSVFile = "";
      String sBrowserRestartInterval = "";
      String sProcessingMode = "";
      int iBrowserRestartInterval = 5;

      ClJobProtocol oJobProtocol = new ClJobProtocol();


      try
      {

        // String sMessageText = File.ReadAllText(@"D:\temp\2025.10.08\DEU TR Extractor 2025-10-08 17-54-50.log");

        // char[] achCharsForTriming = new char[] { ' ', '\t', '\n', '\r' };

        // String sMessageText2 = sMessageText.Trim(achCharsForTriming);


        // int[][] aiMatrix = new int[][] { 
        //   new int[] { 1, 2, 3 }, 
        //   new int[] { 4, 5, 6 },
        //   new int[] { 7, 8, 9 },
        //   new int[] { 10, 11, 12 }
        // };
        //
        // int iMatrixLength = aiMatrix.Length;


        String sProgramVersion = Process.GetCurrentProcess().MainModule.FileVersionInfo.FileVersion;


        Console.WriteLine("Program for web-extarction of data from German Trade Register.");
        Console.WriteLine("Version: " + sProgramVersion);
        Console.WriteLine();
        Console.WriteLine("https://www.handelsregister.de");
        // Console.WriteLine("Created by Dmitry F. Haiduchonak - 2025");
        Console.WriteLine();


        if (args.Length > 0 )
        {
          String sArgument1 = args[0];

          if(sArgument1 != null &&
             sArgument1 == "-pause" )
          {
            // Sleep for 10 seconds.
            Thread.Sleep(10 * 1000);
          }
        }

        try
        {
          IntPtr iConsoleHWnd = Process.GetCurrentProcess().MainWindowHandle;

          ImportsFromWindowsAPI.SetWindowPosition(
            iConsoleHWnd,
            0, 701, 1200, 330,
            "Cannot change position of console window of this process.");
        }
        catch( Exception ex)
        {
          String sLogString =
            "Message: \"" + ex.Message + "\".";

          Console.WriteLine("Not a critiacl error: I cannot change position of current Console window." );
          Console.WriteLine();
          Console.WriteLine(sLogString);
          Console.WriteLine();
          Console.WriteLine();
        }


        DEU_TR_ProgramSettings.Read_CommonSettings_Settings(
          out sTradeRegisterSeachPageUrl,
          out s_sLogFolder,
          out sOutputFolder,
          out sBrowserDownloadFolder,
          out sOutputCSVFile,
          out sBrowserRestartInterval,
          out sProcessingMode);


        oJobProtocol.CreateLogFile(s_sLogFolder, "DEU TR Extractor");
        oJobProtocol.SetAutoFlush(true);

        oJobProtocol.SetPrintToConsole(false);
        oJobProtocol.WriteLine("Program for web-extarction of data from German Trade Register.");
        oJobProtocol.WriteLine("Version: " + sProgramVersion);
        oJobProtocol.WriteLine();
        oJobProtocol.WriteLine("https://www.handelsregister.de");
        // oJobProtocol.WriteLine("Created by Dmitry F. Haiduchonak - 2025");
        oJobProtocol.WriteLine();
        oJobProtocol.SetPrintToConsole(true);
      }
      catch (Exception ex)
      {
        String sLogString =
          "Exception. Class name: " +
          ex.GetType().FullName + ". Message: \"" +
          ex.Message + "\"" + Environment.NewLine +
          "Stack trace:" + Environment.NewLine +
          ex.StackTrace;

        Console.WriteLine(sLogString);
        Console.WriteLine();

        Console.Out.WriteLine("Press Enter to continue: ");
        Console.In.ReadLine();

        return 1;
      }


      try
      {
        DEU_TR_ProgramSettings.Update_ListMode_Statistics();

        Int32.TryParse(sBrowserRestartInterval, out iBrowserRestartInterval);

        // Read proxy list from settings
        List<String> proxyList = DEU_TR_ProgramSettings.ReadProxyList();
        
        if (proxyList.Count > 0)
        {
          oJobProtocol.WriteLine($"Loaded {proxyList.Count} proxies from settings file.");
        }
        else
        {
          oJobProtocol.WriteLine("No proxies configured. Using direct connection.");
        }

        DEU_TradeRegister_WebExtractor oDEU_TradeRegister_WebExtractor =
          new DEU_TradeRegister_WebExtractor(
          sTradeRegisterSeachPageUrl,
          s_sLogFolder,
          sOutputFolder,
          sBrowserDownloadFolder,
          sOutputCSVFile,
          iBrowserRestartInterval,
          proxyList);

        if ( sProcessingMode == "1" )
        {
          String sStartRegisterNumber = "";
          String sEndRegisterNumber = "";

          DEU_TR_ProgramSettings.Read_RangeMode_Settings(
            out sStartRegisterNumber,
            out sEndRegisterNumber);

          oDEU_TradeRegister_WebExtractor.ExtractData_RangeMode(
            oJobProtocol,
            sStartRegisterNumber,
            sEndRegisterNumber);
        }
        else 
        {
          if (sProcessingMode == "2")
          {
            List<String> oRegisterNumberList = new List<String>();

            DEU_TR_ProgramSettings.Read_ListMode_Settings(
              out oRegisterNumberList);

            oDEU_TradeRegister_WebExtractor.ExtractData_ListMode(
              oJobProtocol,
              oRegisterNumberList);
          }
          else if (sProcessingMode == "3")
          {
            String sStartRegisterNumber = "";
            String sEndRegisterNumber = "";

            DEU_TR_ProgramSettings.Read_RangeMode_Settings(
              out sStartRegisterNumber,
              out sEndRegisterNumber);

            oDEU_TradeRegister_WebExtractor.ExtractData_TableScrapingMode(
              oJobProtocol,
              sStartRegisterNumber,
              sEndRegisterNumber);
          }
          else
          {
            String sSettingsXml_FileName = "";

            sSettingsXml_FileName = DEU_TR_ProgramSettings.GetSettingsXmlFileName();

            throw new Exception(
              "Invalid value of ProcessingMode parameter in Settings file: " +
              "\"" + sProcessingMode + "\". Valid values are: 1, 2, or 3. " +
              "(\"" + sSettingsXml_FileName + "\")");
          }
        }
      }
      catch (WebDriverException oWebDriverException)
      {
        String sLogString =
          "Exception. Class name: " +
          oWebDriverException.GetType().FullName + ". Message: \"" +
          oWebDriverException.Message + "\"" + Environment.NewLine +
          "Stack trace:" + Environment.NewLine +
          oWebDriverException.StackTrace;

        oJobProtocol.WriteLine(sLogString);
        oJobProtocol.WriteLine();


        String sErrorMessage = oWebDriverException.Message;

        if (sErrorMessage.IndexOf("The HTTP request to the remote WebDriver server for URL http://localhost:") != -1)
        {
          Process oCurrentProcess = Process.GetCurrentProcess();
          String sCurrentProcessExeFileName = oCurrentProcess.MainModule.FileName;

          oJobProtocol.WriteLine(
            "I am trying to solve problem " + 
            "related to \"hang\" of server part of Selenium Framework:");
          oJobProtocol.WriteLine(
            "restart current process (\"" + oCurrentProcess.MainModule.ModuleName + "\") " +
            "with \"-pause\" argument and exit.");
          oJobProtocol.WriteLine();

          // Restart itself with "-pause" argument.
          Process oNewProcess = Process.Start(sCurrentProcessExeFileName, "-pause");

          // Exit from current process.
          return 2;
        }
      }
      catch (Exception ex)
      {
        String sLogString =
          "Exception. Class name: " +
          ex.GetType().FullName + ". Message: \"" +
          ex.Message + "\"" + Environment.NewLine +
          "Stack trace:" + Environment.NewLine +
          ex.StackTrace;

        oJobProtocol.WriteLine(sLogString);
        oJobProtocol.WriteLine();
      }


      oJobProtocol.CloseLogFile();
      oJobProtocol = null;


      Console.Out.WriteLine("Press Enter to continue: ");
      Console.In.ReadLine();


      return 0;
    }


    /*        
    NetworkResponseHandler oNetworkResponseHandler = new NetworkResponseHandler();

    ClMyResponseHandler oMyResponseHandler = new ClMyResponseHandler();

    oNetworkResponseHandler.ResponseMatcher = oMyResponseHandler.MatchResponse;
    oNetworkResponseHandler.ResponseTransformer = oMyResponseHandler.TransformResponse;

    oBrowserWrapper.Manage().Network.AddResponseHandler(oNetworkResponseHandler);

    oBrowserWrapper.Manage().Network.NetworkResponseReceived += Network_NetworkResponseReceived;

    oBrowserWrapper.Manage().Network.StartMonitoring();
    */

    private static void Network_NetworkResponseReceived(object sender, NetworkResponseReceivedEventArgs e)
    {
      Console.Out.WriteLine();
      Console.Out.WriteLine(",,,,,,,,,");
      Console.Out.WriteLine("I am inside Network_NetworkResponseReceived().");

      // String sTextOfResponse = e.ResponseBody;

      if (e.ResponseResourceType == "Image")
      {
        DateTime oCurrentTime = DateTime.Now;

        String sCurrentTimeStr =
          oCurrentTime.ToString("yyyy-MM-dd HH-mm-ss") +
          "_" + oCurrentTime.Ticks.ToString();

        String sFileName = "ResponseData_" + sCurrentTimeStr + ".txt";

        String sFullFileName = Path.Combine(s_sLogFolder, sFileName);

        // Console.Out.WriteLine(sFileName);
        Console.Out.WriteLine("Saving file: \"" + sFullFileName + "\".");

        // File.WriteAllText(sFullFileName, sTextOfResponse);
        File.WriteAllBytes(sFullFileName, e.ResponseContent.ReadAsByteArray());
      }

      Console.Out.WriteLine(",,,,,,,,,");
      Console.Out.WriteLine();
    }
  }


  public class ClMyResponseHandler
  {
    public bool MatchResponse(HttpResponseData v_oInputData)
    {
      // Console.Out.WriteLine();
      // Console.Out.WriteLine(";;;;;;;;;");
      // Console.Out.WriteLine("I am inside MatchResponse().");
      // Console.Out.WriteLine(";;;;;;;;;");
      // Console.Out.WriteLine();

      return true;
    }

    public HttpResponseData TransformResponse(HttpResponseData v_oInputData)
    {
      // Console.Out.WriteLine();
      // Console.Out.WriteLine("?????????");
      // Console.Out.WriteLine("I am inside TransformResponse().");
      // Console.Out.WriteLine("?????????");
      // Console.Out.WriteLine();

      return v_oInputData;
    }
  }
}
