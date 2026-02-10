using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DEU_TradeRegister_WebExtractor
{
  public class StreamingCSVWriter : IDisposable
  {
    private StreamWriter m_writer;
    private char m_delimiter;
    private char m_quoteChar;
    private bool m_headerWritten;
    private List<string> m_columnNames;
    private bool m_disposed;

    public StreamingCSVWriter(string filePath, char delimiter = ';', char quoteChar = '"', bool append = false)
    {
      m_delimiter = delimiter;
      m_quoteChar = quoteChar;
      m_headerWritten = false;
      m_columnNames = new List<string>();
      m_disposed = false;

      // Check if file exists and we're appending
      bool fileExists = File.Exists(filePath);
      
      // Create file with UTF-8 BOM encoding for Excel compatibility
      // If appending and file exists, don't write BOM again
      if (append && fileExists)
      {
        m_writer = new StreamWriter(filePath, true, new UTF8Encoding(false));
        m_headerWritten = true; // Skip header if appending to existing file
      }
      else
      {
        m_writer = new StreamWriter(filePath, false, new UTF8Encoding(true));
      }
      
      // Disable buffering for immediate writes
      m_writer.AutoFlush = true;
    }

    public void WriteHeader(params string[] columns)
    {
      if (m_headerWritten)
        throw new InvalidOperationException("Header has already been written.");

      m_columnNames.AddRange(columns);
      
      string headerLine = string.Join(m_delimiter.ToString(), 
        columns.Select(col => EscapeCSVField(col)));
      
      m_writer.WriteLine(headerLine);
      m_headerWritten = true;
    }

    public void SetColumnNames(params string[] columns)
    {
      // Set column names without writing header (for append mode)
      if (m_columnNames.Count > 0)
        throw new InvalidOperationException("Column names have already been set.");
      
      m_columnNames.AddRange(columns);
    }

    public void WriteRow(params string[] values)
    {
      if (!m_headerWritten)
        throw new InvalidOperationException("Header must be written before writing rows.");

      if (values.Length != m_columnNames.Count)
        throw new ArgumentException(
          $"Row must have {m_columnNames.Count} values, but {values.Length} were provided.");

      string rowLine = string.Join(m_delimiter.ToString(), 
        values.Select(val => EscapeCSVField(val ?? string.Empty)));
      
      m_writer.WriteLine(rowLine);
      // AutoFlush is true, so data is immediately written to disk
    }

    public void WriteRow(Dictionary<string, string> fieldValues)
    {
      if (!m_headerWritten)
        throw new InvalidOperationException("Header must be written before writing rows.");

      string[] values = new string[m_columnNames.Count];
      for (int i = 0; i < m_columnNames.Count; i++)
      {
        string columnName = m_columnNames[i];
        values[i] = fieldValues.ContainsKey(columnName) ? fieldValues[columnName] : string.Empty;
      }

      WriteRow(values);
    }

    public void WriteCompanyData(CompanyTradeRegisterData data)
    {
      WriteRow(
        data.LocalCourtId,
        data.LocalCourtName,
        data.TypeOfRegister,
        data.RegistrationNumber,
        data.RegistrationDate_FormattedDate,
        data.RegistrationDate_UnformattedDate,
        data.CompanyName,
        data.LegalForm,
        data.CityName_WhereMainOfficeIsLocated,
        data.PostalAddressType,
        data.StreetName,
        data.HouseNumber,
        data.PostalCode,
        data.CityName,
        data.CountryName,
        data.TypeOfRepresentation,
        data.ShareCapital_GmbH_TotalAmountOfMoney,
        data.ShareCapital_GmbH_OutdatedCurrency,
        data.ShareCapital_GmbH_ActualCurrency,
        data.ShareCapital_AG_TotalAmountOfMoney,
        data.ShareCapital_AG_ActualCurrency,
        data.BranchText
      );
    }

    private string EscapeCSVField(string field)
    {
      if (string.IsNullOrEmpty(field))
        return string.Empty;

      // Check if field needs quoting
      bool needsQuoting = field.Contains(m_delimiter) || 
                          field.Contains(m_quoteChar) || 
                          field.Contains('\n') || 
                          field.Contains('\r');

      if (needsQuoting)
      {
        // Escape quote characters by doubling them
        string escaped = field.Replace(m_quoteChar.ToString(), 
                                      new string(m_quoteChar, 2));
        return m_quoteChar + escaped + m_quoteChar;
      }

      return field;
    }

    public void Flush()
    {
      m_writer?.Flush();
    }

    public void Close()
    {
      Dispose();
    }

    public void Dispose()
    {
      if (!m_disposed)
      {
        if (m_writer != null)
        {
          m_writer.Flush();
          m_writer.Close();
          m_writer.Dispose();
          m_writer = null;
        }
        m_disposed = true;
      }
    }
  }
}
