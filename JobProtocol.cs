using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APX
{
  public class ClJobProtocol
  {
    public ClJobProtocol()
    {
      m_oLogFileWriter = null;
      m_sLogFileName = "";
      m_bPrintToConsole = true;
    }

    ~ClJobProtocol()
    {
      if(m_oLogFileWriter != null)
      {
        m_oLogFileWriter.Close();
        m_oLogFileWriter = null;
      }
    }

    public void CreateLogFile(String v_sFolderName, String v_sFilePrefixName)
    {
      if (v_sFolderName == null)
      {
        throw new Exception(
          "Cannot create a log file for Job Protocol. " +
          "Input parameter v_sFolderName is 'null'.");
      }

      if (v_sFilePrefixName == null)
      {
        throw new Exception(
          "Cannot create a log file for Job Protocol. " +
          "Input parameter v_sFilePrefixName is 'null'.");
      }

      if (m_oLogFileWriter != null)
      {
        throw new Exception(
          "Cannot create a log file for Job Protocol. " +
          "Log file is already open. (\"" + m_sLogFileName + "\") " +
          "Plase, call CloseLogFile() method before call of this method.");
      }

      DateTime oCurrentDateTime = DateTime.Now;

      String sLogFileNameWOFolder =
        v_sFilePrefixName + " " +
        oCurrentDateTime.Year.ToString("D4") + "-" +
        oCurrentDateTime.Month.ToString("D2") + "-" +
        oCurrentDateTime.Day.ToString("D2") + " " +
        oCurrentDateTime.Hour.ToString("D2") + "-" +
        oCurrentDateTime.Minute.ToString("D2") + "-" +
        oCurrentDateTime.Second.ToString("D2") +
        ".log";

      String sLogFileName = Path.Combine(v_sFolderName, sLogFileNameWOFolder);

      m_oLogFileWriter = new StreamWriter(sLogFileName, false, Encoding.Unicode);

      m_sLogFileName = sLogFileName;
    }

    public void CloseLogFile()
    {
      if (m_oLogFileWriter != null)
      {
        m_oLogFileWriter.Close();
        m_oLogFileWriter = null;

        m_sLogFileName = "";
      }
    }

    public bool GetLogFileIsOpen()
    {
      bool bResult = false;

      if (m_oLogFileWriter != null)
        bResult = true;

      return bResult;
    }

    public bool GetPrintToConsole()
    {
      return m_bPrintToConsole;
    }

    public void SetPrintToConsole(bool v_bPrintToConsole)
    {
      m_bPrintToConsole = v_bPrintToConsole;
    }

    public void WriteLine(String v_sInputLine)
    {
      if (v_sInputLine == null)
      {
        throw new Exception(
          "Cannot write 'String' to Job Protocol. " +
          "Input parameter v_sInputLine is 'null'.");
      }

      if (m_oLogFileWriter != null )
      {
        m_oLogFileWriter.WriteLine(v_sInputLine);
      }

      if(m_bPrintToConsole)
      {
        Console.Out.WriteLine(v_sInputLine);
      }  
    }

    public void WriteLine()
    {
      if (m_oLogFileWriter != null)
      {
        m_oLogFileWriter.WriteLine();
      }

      if (m_bPrintToConsole)
      {
        Console.Out.WriteLine();
      }
    }

    public void Write(String v_sInputLine)
    {
      if (v_sInputLine == null)
      {
        throw new Exception(
          "Cannot write 'String' to Job Protocol. " +
          "Input parameter v_sInputLine is 'null'.");
      }

      if (m_oLogFileWriter != null)
      {
        m_oLogFileWriter.Write(v_sInputLine);
      }

      if (m_bPrintToConsole)
      {
        Console.Out.Write(v_sInputLine);
      }
    }

    public void Flush()
    {
      if (m_oLogFileWriter != null)
      {
        m_oLogFileWriter.Flush();
      }

      if (m_bPrintToConsole)
      {
        Console.Out.Flush();
      }
    }

    public bool GetAutoFlush()
    {
      if (m_oLogFileWriter == null)
      {
        throw new Exception(
          "Log file was not created for Job Protocol.");
      }

      return m_oLogFileWriter.AutoFlush;
    }

    public void SetAutoFlush(bool v_bAutoFlush)
    {
      if (m_oLogFileWriter == null)
      {
        throw new Exception(
          "Log file was not created for Job Protocol.");
      }

      m_oLogFileWriter.AutoFlush = v_bAutoFlush;
    }

    private StreamWriter m_oLogFileWriter;
    private String m_sLogFileName;
    private bool m_bPrintToConsole;
  };

}
