// DEU_TradeRegister_Converter.cpp : This file contains the 'main' function. Program execution begins and ends there.
//


// Enable Memory Leak Detection in debug mode.
#ifdef _DEBUG
#define _CRTDBG_MAP_ALLOC
#include <stdlib.h>
#if defined WIN32
#include <crtdbg.h>
#endif
#endif 

#include <stdio.h>

#include <string>
using std::string;

#include <io.h>
#include <fcntl.h>
#include <locale>
using std::locale;

#include <vector>
using std::vector;

#include <list>
using std::list;

#include <iostream>
#include <fstream>
#include <sstream>
using std::cout;
using std::wcout;
using std::wfstream;
using std::wofstream;
using std::ios;
using std::wostringstream;
using std::endl;

#include "pugixml.hpp"

#include "ENVException.h"
using FLX::EGeneric;
using FLX::g_pszCannotCreateFileErrNo;
using FLX::g_pszCannotWriteToFileErrNo;
using FLX::ThrowWin32FileException;
using FLX::ThrowAnsiFileException;
using FLX::ThrowWin32ExceptionWW;
using FLX::TranslateFmt;
using FLX::TranslateFmtW;

#include "GPStrUtils.h"
using FLX::IntToStr;
using FLX::Int64ToStr;
using FLX::DoubleToStr;
using FLX::ParseCSVStringW2;

#include "GPFileUtils.h"
using FLX::MakePathW;

#include "LCLCodePageExt.h"
using FLX::ClCodePageExt;

#include "IOFileDEL.h"
using FLX::ClFileDEL;

#include "APXProtocol.h"
using APX::ClJobProtocol;

#include "TestPugiXML.h"


/*
void TestJobProtocol()
{
  ClJobProtocol oJobProtocol;

  // oJobProtocol.Assign("C:\\Work\\2025.12.20\\", "TestProt", true);

  oJobProtocol.Assign(L"C:\\Work\\2025.12.20\\", L"TestProt", L"Суффикс", true);

  oJobProtocol.SetOnScreen(true);
  oJobProtocol.SetUnicodeForConsoleWindowIsEnabled(true);

  oJobProtocol << "Hello!";
  oJobProtocol.endl();

  oJobProtocol << L"Привет 1.\n";
  oJobProtocol.endl();

  oJobProtocol << L"Числа: " << 1 << ", " << 2 << ", " << 3 << L"(три)." << "\n";
  oJobProtocol.flush();
  oJobProtocol.endl();
  oJobProtocol.endl();

  oJobProtocol << "Tests for specific data types:";
  oJobProtocol.endl();

  string sString("std::string");
  oJobProtocol << "  ClJobProtocol::operator<<(const string& v_sString)........: \"" << sString << "\"";
  oJobProtocol.endl();

  string svString("std::string_view");
  oJobProtocol << "  ClJobProtocol::operator<<(const string_view& v_svString)..: \"" << svString << "\"";
  oJobProtocol.endl();

  w_string wsString(L"w_string  АБВГДЕЁ");
  oJobProtocol << "  ClJobProtocol::operator<<(const w_string& v_wsString).....: L\"" << wsString << "\"";
  oJobProtocol.endl();

  wstring_view wsvString(L"std::wstring_view  АБВГДЕЁ");
  oJobProtocol << "  ClJobProtocol::operator<<(const wstring_view& v_wsvString): L\"" << wsvString << "\"";
  oJobProtocol.endl();


  const char* szString = "const char*";
  oJobProtocol << "  ClJobProtocol::operator<<(const char* v_szString).........: \"" << szString << "\"";
  oJobProtocol.endl();

  char chChar = 'C';
  oJobProtocol << "  ClJobProtocol::operator<<(char v_chChar)..................: \"" << chChar << "\"";
  oJobProtocol.endl();

  unsigned char uchChar = 'C';
  oJobProtocol << "  ClJobProtocol::operator<<(unsigned char v_chChar).........: \"" << uchChar << "\"";
  oJobProtocol.endl();

  const wchar_t* wwszString = L"const wchar_t*  АБВГДЕЁ";
  oJobProtocol << "  ClJobProtocol::operator<<(const wchar_t* v_wwszString)....: \"" << wwszString << "\"";
  oJobProtocol.endl();

  wchar_t wwchChar = L'Ё';
  oJobProtocol << "  ClJobProtocol::operator<<(wchar_t v_wwchChar).............: \"" << wwchChar << "\"";
  oJobProtocol.endl();

  short iShortValue = -16;
  oJobProtocol << "  ClJobProtocol::operator<<(short v_iValue).................: \"" << iShortValue << "\"";
  oJobProtocol.endl();

  unsigned short uiShortValue = 16;
  oJobProtocol << "  ClJobProtocol::operator<<(unsigned short v_uiValue).......: \"" << uiShortValue << "\"";
  oJobProtocol.endl();

  int iInteger = -32;
  oJobProtocol << "  ClJobProtocol::operator<<(int v_iInteger).................: \"" << iInteger << "\"";
  oJobProtocol.endl();

  unsigned int uiInteger = 32;
  oJobProtocol << "  ClJobProtocol::operator<<(unsigned int v_uiInteger).......: \"" << uiInteger << "\"";
  oJobProtocol.endl();

  FLX::INT64 iInt64 = -64;
  oJobProtocol << "  ClJobProtocol::operator<<(FLX::INT64 v_iValue)............: \"" << iInt64 << "\"";
  oJobProtocol.endl();

  FLX::UINT64 uiInt64 = 64;
  oJobProtocol << "  ClJobProtocol::operator<<(FLX::UINT64 v_uiValue)..........: \"" << uiInt64 << "\"";
  oJobProtocol.endl();

  float fValue = -1.1f;
  oJobProtocol << "  ClJobProtocol::operator<<(float v_fValue).................: \"" << fValue << "\"";
  oJobProtocol.endl();

  double dValue = -1.01;
  oJobProtocol << "  ClJobProtocol::operator<<(double v_dValue)................: \"" << dValue << "\"";
  oJobProtocol.endl();

  long double ldValue = -1.001;
  oJobProtocol << "  ClJobProtocol::operator<<(long double v_ldValue)..........: \"" << ldValue << "\"";
  oJobProtocol.endl();

  bool bValueTrue = true;
  bool bValueFalse = false;
  oJobProtocol << "  ClJobProtocol::operator<<(bool v_bValue)..................: \"" << 
    bValueTrue << "\", \"" << bValueFalse << "\"";
  oJobProtocol.endl();

  std::exception oStdException("Test message for 'std::exception'.");
  oJobProtocol << "  ClJobProtocol::operator<<(const std::exception& v_oValue).: \"" << oStdException << "\"";
  oJobProtocol.endl();

  EGeneric oException(ERX_LOGICAL, "Test message for 'FLX::EGeneric'.");
  oJobProtocol << "  ClJobProtocol::operator<<(const FLX::EGeneric& v_oValue).: \"" << oException << "\"";
  oJobProtocol.endl();


  oJobProtocol.SetUnicodeForConsoleWindowIsEnabled(false);

 
  oJobProtocol.Close();
}
*/

void TestUnicodeOutputInCpp()
{
  // Set output codepage to Unicode for Console Window
  _setmode(_fileno(stdout), _O_U16TEXT);

  wprintf(L"Привет 1.\n");
  std::wcout << L"Привет 2.";


  // Set output codepage to ANSI for Console Window
  _setmode(_fileno(stdout), _O_TEXT);

  printf("Hello 1.\n");
  std::cout << "Hello 2.\n";


  // The best way to write Unicode characters to a file on disk by fstream class
  //   is to switch it into Binary mode and write binary codes of Unicode characters.
  // In this approach one should also write BOM.
  // 
  // https://ru.stackoverflow.com/questions/289383/unicode-%D0%BF%D1%80%D0%B8-%D0%B7%D0%B0%D0%BF%D0%B8%D1%81%D0%B8-%D0%B2-%D1%84%D0%B0%D0%B9%D0%BB


  string sFileName = "C:\\Work\\2025.12.20\\3.txt";

  FILE* poOutputFile = NULL;

  {
    errno_t iErrNo = 0;

    iErrNo = fopen_s(&poOutputFile, sFileName.c_str(), "w, ccs=UTF-16LE");
    if (iErrNo != 0 || poOutputFile == NULL)
      throw EGeneric(ERX_CANNOT_CREATE_FILE, TranslateFmt(g_pszCannotCreateFileErrNo, sFileName.c_str(), iErrNo));
  }

  wostringstream oWideCharOutputStringStream;

  oWideCharOutputStringStream << L"Привет.\n";
  oWideCharOutputStringStream << 11;
  oWideCharOutputStringStream << endl;

  wstring wsOutputString = oWideCharOutputStringStream.str();

  wcout << wsOutputString;

  if (fputws(wsOutputString.c_str(), poOutputFile) == EOF)
  {
    errno_t iErrNo = errno;

    throw EGeneric(ERX_IO_ERROR, TranslateFmt(g_pszCannotWriteToFileErrNo, sFileName.c_str(), iErrNo));
  }

  fclose(poOutputFile);
}


// Searches in target directory for all files
//   whose names are match specified file name mask.
// Fills output list with full names of found files.
//
//   v_wsTargetDirectory  - should contain the name of direcory,
//                            where files will be searched.
// 
//   v_wsFileNameMask     - mask of the name of files that should be found
//                            (e.g. "*.xml").
// 
//   v_oFoundFileNameList - list of full names of found files.
//
void FindAllFilesInDirectory_UsingWildcards(
  const w_string& v_wsTargetDirectory,
  const w_string& v_wsFileNameMask,
  list<w_string>& v_oFoundFileNameList // out parameter
  )
{
  HANDLE hFind = NULL;
  WIN32_FIND_DATAW oFindFileDataW;
  

  v_oFoundFileNameList.clear();

  memset(&oFindFileDataW, 0, sizeof(oFindFileDataW));

  w_string wsFullFileNameMask = MakePathW(v_wsTargetDirectory, v_wsFileNameMask);
  LPCWSTR lpFileName = wsFullFileNameMask.c_str();

  hFind = FindFirstFileExW(
    lpFileName, 
    FindExInfoStandard, 
    &oFindFileDataW,
    FindExSearchNameMatch,
    NULL, 
    0);
  if (hFind != INVALID_HANDLE_VALUE)
  {
    if ((oFindFileDataW.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
    {
      w_string wsFoundFileFullFileName =
        MakePathW(v_wsTargetDirectory, oFindFileDataW.cFileName);

      v_oFoundFileNameList.push_back(wsFoundFileFullFileName);
    }


    bool bNoMoreFiles = false;

    do
    {
      BOOL iResult = FindNextFileW(hFind, &oFindFileDataW);
      if (iResult != 0)
      {
        if ((oFindFileDataW.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
        {
          w_string wsFoundFileFullFileName =
            MakePathW(v_wsTargetDirectory, oFindFileDataW.cFileName);

          v_oFoundFileNameList.push_back(wsFoundFileFullFileName);
        }
      }
      else
      {
        DWORD iLastError = GetLastError();

        if (iLastError == ERROR_NO_MORE_FILES)
        {
          bNoMoreFiles = true;
        }
        else
        {
          ThrowWin32ExceptionWW(L"Error in FindAllFilesInDirectory_UsingWildcards() function.");
        }
      }
    } while (!bNoMoreFiles);

    FindClose(hFind);
  }
}

// Searches in target directory for all directories
//   whose names are match specified directory name mask.
// Fills output list with full names of found directories.
//
//   v_wsTargetDirectory   - should contain the name of direcory,
//                             where files will be searched.
// 
//   v_wsDirectoryNameMask - mask of the name of directories that should be found
//                            (e.g. "2025-01-11 *").
// 
//   v_oFoundDirectoryNameList - list of full names of found directories.
//
void FindAllDirectoriesInDirectory_UsingWildcards(
  const w_string& v_wsTargetDirectory,
  const w_string& v_wsDirectoryNameMask,
  list<w_string>& v_oFoundDirectoryNameList // out parameter
)
{
  HANDLE hFind = NULL;
  WIN32_FIND_DATAW oFindFileDataW;


  v_oFoundDirectoryNameList.clear();

  memset(&oFindFileDataW, 0, sizeof(oFindFileDataW));

  w_string wsFullDirectoryNameMask = MakePathW(v_wsTargetDirectory, v_wsDirectoryNameMask);
  LPCWSTR lpFileName = wsFullDirectoryNameMask.c_str();

  hFind = FindFirstFileExW(
    lpFileName,
    FindExInfoStandard,
    &oFindFileDataW,
    FindExSearchNameMatch,
    NULL,
    0);
  if (hFind != INVALID_HANDLE_VALUE)
  {
    if ((oFindFileDataW.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
      w_string wsDirectoryName = oFindFileDataW.cFileName;

      if (wsDirectoryName != L"." && wsDirectoryName != L"..")
      {
        w_string wsFoundFileFullDirectoryName =
          MakePathW(v_wsTargetDirectory, wsDirectoryName);

        v_oFoundDirectoryNameList.push_back(wsFoundFileFullDirectoryName);
      }
    }


    bool bNoMoreDirectories = false;

    do
    {
      BOOL iResult = FindNextFileW(hFind, &oFindFileDataW);
      if (iResult != 0)
      {
        if ((oFindFileDataW.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
          w_string wsDirectoryName = oFindFileDataW.cFileName;

          if (wsDirectoryName != L"." && wsDirectoryName != L"..")
          {
            w_string wsFoundFileFullDirectoryName =
              MakePathW(v_wsTargetDirectory, wsDirectoryName);

            v_oFoundDirectoryNameList.push_back(wsFoundFileFullDirectoryName);
          }
        }
      }
      else
      {
        DWORD iLastError = GetLastError();

        if (iLastError == ERROR_NO_MORE_FILES)
        {
          bNoMoreDirectories = true;
        }
        else
        {
          ThrowWin32ExceptionWW(L"Error in FindAllDirectoriesInDirectory_UsingWildcards() function.");
        }
      }
    } while (!bNoMoreDirectories);

    FindClose(hFind);
  }
}

class ClCompanyTradeRegisterData
{
public:

  ClCompanyTradeRegisterData();

  void Clear();


  // * Registartion data of company

  // Local Court Id where this company is registered
  //   (for example: "D2601V").
  w_string m_wsLocalCourtId;

  // Real name of registartion court from XML comment
  //   (for exmaple: "<!--Registergericht Amtsgericht München-->").
  w_string m_wsLocalCourtName;

  // Type of register
  //   (for example: "HRB").
  w_string m_wsTypeOfRegister;

  // Registration number
  //   (for example: "116828").
  w_string m_wsRegistrationNumber;



  // * Registration date of company ("Statutes date")
  //    (i.e. date of registration of the company in the German Trade Register).

  // Formatted date ("yyyy-MM-dd").
  w_string m_wsRegistrationDate_FormattedDate;

  // Unformatted date (i.e. free text).
  w_string m_wsRegistrationDate_UnformattedDate;



  // * Other data of company

  // The name of company (legal entity).
  w_string m_wsCompanyName;

  // The type of Legal Form.
  w_string m_wsLegalForm;

  // The name of city where the main office of the company is located.
  w_string m_wsCityName_WhereMainOfficeIsLocated;



  // * Postal address of the main office of the company.

  // Type of postal address.
  w_string m_wsPostalAddressType;

  // Street name.
  w_string m_wsStrretName;

  // House number.
  w_string m_wsHouseNumber;

  // Postal Code (PLZ).
  w_string m_wsPostalCode;

  // City name.
  w_string m_wsCityName;

  // Country name (may absent in XML document).
  w_string m_wsCountryName;



  // Type of represenatation ("allgemeineVertretungsregelung") of company.
  w_string m_wsTypeOfRepresentation;



  // * Share capital of company (for GmbH):

  // Total amount of money.
  w_string m_wsShareCapital_GmbH_TotalAmountOfMoney;

  // Currency code - Variant 1
  //   (outdated currency, e.g. "DEM").
  w_string m_wsShareCapital_GmbH_OutdatedCurrency;

  // Currency code - Variant 2
  //   (actual currency, e.g. "EUR").
  w_string m_wsShareCapital_GmbH_ActualCurrency;



  // * Share capital of company (for AG):

  // Total amount of money.
  w_string m_wsShareCapital_AG_TotalAmountOfMoney;

  // Currency code
  //   (actual currency, e.g. "EUR").
  w_string m_wsShareCapital_AG_ActualCurrency;



  // Branch of company.
  w_string m_wsBranchText;
};

ClCompanyTradeRegisterData::ClCompanyTradeRegisterData()
{
  Clear();
}

void 
ClCompanyTradeRegisterData::Clear()
{
  m_wsLocalCourtId.clear();
  m_wsLocalCourtName.clear();
  m_wsTypeOfRegister.clear();
  m_wsRegistrationNumber.clear();


  m_wsRegistrationDate_FormattedDate.clear();
  m_wsRegistrationDate_UnformattedDate.clear();


  m_wsCompanyName.clear();


  m_wsLegalForm.clear();


  m_wsCityName_WhereMainOfficeIsLocated.clear();


  m_wsPostalAddressType.clear();
  m_wsStrretName.clear();
  m_wsHouseNumber.clear();
  m_wsPostalCode.clear();
  m_wsCityName.clear();
  m_wsCountryName.clear();


  m_wsTypeOfRepresentation.clear();


  m_wsShareCapital_GmbH_TotalAmountOfMoney.clear();
  m_wsShareCapital_GmbH_OutdatedCurrency.clear();
  m_wsShareCapital_GmbH_ActualCurrency.clear();


  m_wsShareCapital_AG_TotalAmountOfMoney.clear();
  m_wsShareCapital_AG_ActualCurrency.clear();


  m_wsBranchText.clear();
}

typedef list<w_string> ClInvalidFileNameList;
typedef list<ClCompanyTradeRegisterData> ClCompanyTradeRegisterDataList;

void
LoadXMLFileForCompany(
  const w_string& v_cwsXMLFileName,
  ClCompanyTradeRegisterData& v_oCompanyTradeRegisterData, // "out" parameter
  ClInvalidFileNameList v_oRegistrationData_CannotBeExtracted_FileNameList, // "in" and "out" parameter
  ClInvalidFileNameList v_oRegistrationDate_CannotBeExtracted_FileNameList, // "in" and "out" parameter
  ClInvalidFileNameList v_oOtherDataOfCompany_CannotBeExtracted_FileNameList, // "in" and "out" parameter
  ClInvalidFileNameList v_oPostalAddress_CannotBeExtracted_FileNameList, // "in" and "out" parameter
  ClInvalidFileNameList v_oTypeOfRepresentation_CannotBeExtracted_FileNameList, // "in" and "out" parameter
  ClInvalidFileNameList v_oShareCapital_CannotBeExtracted_FileNameList, // "in" and "out" parameter
  ClInvalidFileNameList v_oBranchText_CannotBeExtracted_FileNameList, // "in" and "out" parameter
  ClJobProtocol& v_oJobProtocol)
{
  v_oJobProtocol << L"Loading file \"" << v_cwsXMLFileName << L"\": ";
  v_oJobProtocol.flush();


  ClCodePageExt oCPConverter;

  bool bRegistrationData_CannotBeExtracted = false;
  bool bRegistrationDate_CannotBeExtracted = false;
  bool bOtherDataOfCompany_CannotBeExtracted = false;
  bool bPostalAddress_CannotBeExtracted = false;
  bool bTypeOfRepresentation_CannotBeExtracted = false;
  bool bShareCapital_CannotBeExtracted = false;
  bool bBranchText_CannotBeExtracted = false;

  v_oCompanyTradeRegisterData.Clear();

  wstring wwsXMLFileName = W_StrToWStr(v_cwsXMLFileName);

  pugi::xml_document oXmlDocument;
  pugi::xml_parse_result oParseResult = oXmlDocument.load_file(wwsXMLFileName.c_str());

  if (!oParseResult)
  {
    wstring wwsDescription = oCPConverter.ConvertStringToWWString(string(oParseResult.description()));
    long long llOffset = oParseResult.offset;

    throw EGeneric(ERX_LOGICAL,
      TranslateFmtW(
        L"XML file parsing error.\nFile: \"%s\".\nOffset: %I64d\nDescription: \"%s\".",
        wwsXMLFileName.c_str(),
        llOffset,
        wwsDescription.c_str()));
  }


  // * Registartion data of company


  // Local Court Id where this company is registered
  //   (for example: "D2601V").
  pugi::xpath_node oLocalCourtId_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:grunddaten/tns:verfahrensdaten/tns:instanzdaten/tns:auswahl_instanzbehoerde/tns:gericht/code");
  if (oLocalCourtId_Node && oLocalCourtId_Node.node())
  {
    wstring wwsNodeValue = oLocalCourtId_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsLocalCourtId = wwsNodeValue;
  }
  else
  {
    bRegistrationData_CannotBeExtracted = true;
  }


  // .. w_string m_wsLocalCourtName;


  // Type of register
  //   (for example: "HRB").
  pugi::xpath_node oTypeOfRegister_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:grunddaten/tns:verfahrensdaten/tns:instanzdaten/tns:aktenzeichen/tns:auswahl_aktenzeichen/tns:aktenzeichen.strukturiert/tns:register/code");
  if (oTypeOfRegister_Node && oTypeOfRegister_Node.node())
  {
    wstring wwsNodeValue = oTypeOfRegister_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsTypeOfRegister = wwsNodeValue;
  }
  else
  {
    bRegistrationData_CannotBeExtracted = true;
  }

  // Registration number
  //   (for example: "116828").
  pugi::xpath_node oRegistrationNumber_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:grunddaten/tns:verfahrensdaten/tns:instanzdaten/tns:aktenzeichen/tns:auswahl_aktenzeichen/tns:aktenzeichen.strukturiert/tns:laufendeNummer");
  if (oRegistrationNumber_Node && oRegistrationNumber_Node.node())
  {
    wstring wwsNodeValue = oRegistrationNumber_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsRegistrationNumber = wwsNodeValue;
  }
  else
  {
    bRegistrationData_CannotBeExtracted = true;
  }


  if (bRegistrationData_CannotBeExtracted)
    v_oRegistrationData_CannotBeExtracted_FileNameList.push_back(v_cwsXMLFileName);

  
  // * Registration date of company ("Statutes date")
  //    (i.e. date of registration of the company in the German Trade Register).

  // Formatted date ("yyyy-MM-dd").
  pugi::xpath_node oRegistrationDate_FormattedDate_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:satzungsdatum/tns:aktuellesSatzungsdatum");
  if (oRegistrationDate_FormattedDate_Node && oRegistrationDate_FormattedDate_Node.node())
  {
    wstring wwsNodeValue = oRegistrationDate_FormattedDate_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsRegistrationDate_FormattedDate = wwsNodeValue;
  }
  else
  {
    bRegistrationDate_CannotBeExtracted = true;
  }

  // Unformatted date (i.e. free text).
  pugi::xpath_node oRegistrationDate_UnformattedDate_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:satzungsdatum/tns:satzungsdatumFreitext");
  if (oRegistrationDate_UnformattedDate_Node && oRegistrationDate_UnformattedDate_Node.node())
  {
    wstring wwsNodeValue = oRegistrationDate_UnformattedDate_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsRegistrationDate_UnformattedDate = wwsNodeValue;
  }
  else
  {
    bRegistrationDate_CannotBeExtracted = true;
  }


  if (bRegistrationDate_CannotBeExtracted)
    v_oRegistrationDate_CannotBeExtracted_FileNameList.push_back(v_cwsXMLFileName);



  // * Other data of company

  // The name of company (legal entity).
  pugi::xpath_node oCompanyName_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:bezeichnung/tns:bezeichnung.aktuell");
  if (oCompanyName_Node && oCompanyName_Node.node())
  {
    wstring wwsNodeValue = oCompanyName_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsCompanyName = wwsNodeValue;
  }
  else
  {
    bOtherDataOfCompany_CannotBeExtracted = true;
  }


  // The type of Legal Form.
  pugi::xpath_node oLegalForm_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:angabenZurRechtsform/tns:rechtsform/code");
  if (oLegalForm_Node && oLegalForm_Node.node())
  {
    wstring wwsNodeValue = oLegalForm_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsLegalForm = wwsNodeValue;
  }
  else
  {
    bOtherDataOfCompany_CannotBeExtracted = true;
  }


  // The name of city where the main office of the company is located.
  pugi::xpath_node oCityName_WhereMainOfficeIsLocated_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:sitz/tns:ort");
  if (oCityName_WhereMainOfficeIsLocated_Node && 
      oCityName_WhereMainOfficeIsLocated_Node.node())
  {
    wstring wwsNodeValue = oCityName_WhereMainOfficeIsLocated_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsCityName_WhereMainOfficeIsLocated = wwsNodeValue;
  }
  else
  {
    bOtherDataOfCompany_CannotBeExtracted = true;
  }


  if (bOtherDataOfCompany_CannotBeExtracted)
    v_oOtherDataOfCompany_CannotBeExtracted_FileNameList.push_back(v_cwsXMLFileName);



  // * Postal address of the main office of the company.

  // Type of postal address.
  pugi::xpath_node oPostalAddressType_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:anschrift/tns:anschriftstyp/code");
  if (oPostalAddressType_Node && oPostalAddressType_Node.node())
  {
    wstring wwsNodeValue = oPostalAddressType_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsPostalAddressType = wwsNodeValue;
  }
  else
  {
    bPostalAddress_CannotBeExtracted = true;
  }


  // Street name.
  pugi::xpath_node oStrretName_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:anschrift/tns:strasse");
  if (oStrretName_Node && oStrretName_Node.node())
  {
    wstring wwsNodeValue = oStrretName_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsStrretName = wwsNodeValue;
  }
  else
  {
    bPostalAddress_CannotBeExtracted = true;
  }


  // House number.
  pugi::xpath_node oHouseNumber_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:anschrift/tns:hausnummer");
  if (oHouseNumber_Node && oHouseNumber_Node.node())
  {
    wstring wwsNodeValue = oHouseNumber_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsHouseNumber = wwsNodeValue;
  }
  else
  {
    bPostalAddress_CannotBeExtracted = true;
  }


  // Postal Code (PLZ).
  pugi::xpath_node oPostalCode_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:anschrift/tns:postleitzahl");
  if (oPostalCode_Node && oPostalCode_Node.node())
  {
    wstring wwsNodeValue = oPostalCode_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsPostalCode = wwsNodeValue;
  }
  else
  {
    bPostalAddress_CannotBeExtracted = true;
  }


  // City name.
  pugi::xpath_node oCityName_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:anschrift/tns:ort");
  if (oCityName_Node && oCityName_Node.node())
  {
    wstring wwsNodeValue = oCityName_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsCityName = wwsNodeValue;
  }
  else
  {
    bPostalAddress_CannotBeExtracted = true;
  }


  // Country name (may absent in XML document).
  pugi::xpath_node oCountryName_Node = oXmlDocument.select_node(
    L"/tns:nachricht.reg.0400003/tns:fachdatenRegister/tns:basisdatenRegister/tns:rechtstraeger/tns:anschrift/tns:staat/code");
  if (oCountryName_Node && oCountryName_Node.node())
  {
    wstring wwsNodeValue = oCountryName_Node.node().text().get();

    v_oCompanyTradeRegisterData.m_wsCountryName = wwsNodeValue;
  }
  else
  {
    bPostalAddress_CannotBeExtracted = true;
  }


  if (bPostalAddress_CannotBeExtracted)
    v_oPostalAddress_CannotBeExtracted_FileNameList.push_back(v_cwsXMLFileName);



  v_oJobProtocol << L"OK";
  v_oJobProtocol.endl();
}

void
SaveCompanyTradeRegisterDataListToFile(
  const ClCompanyTradeRegisterDataList& v_coCompanyTradeRegisterDataList,
  const w_string v_cwsOutputFileName,
  ClJobProtocol& v_oJobProtocol)
{
  v_oJobProtocol << L"Saving file \"" << v_cwsOutputFileName << L"\" (one dot per each 10 000 records): ";
  v_oJobProtocol.flush();


  ClCompanyTradeRegisterDataList::const_iterator citCompanyTradeRegisterData;

  ClFileDEL oOutputFile;
  ClCodePageExt oCPConverter;
  string sOutputFileName;

  sOutputFileName = oCPConverter.ConvertW_StringToString(v_cwsOutputFileName);


  oOutputFile.m_oFields.Add("LocalCourtId");
  oOutputFile.m_oFields.Add("LocalCourtName");
  oOutputFile.m_oFields.Add("TypeOfRegister");
  oOutputFile.m_oFields.Add("RegistrationNumber");


  oOutputFile.m_oFields.Add("RegistrationDate_FormattedDate");
  oOutputFile.m_oFields.Add("RegistrationDate_UnformattedDate");


  oOutputFile.m_oFields.Add("CompanyName");


  oOutputFile.m_oFields.Add("LegalForm");


  oOutputFile.m_oFields.Add("CityName_WhereMainOfficeIsLocated");


  oOutputFile.m_oFields.Add("PostalAddressType");
  oOutputFile.m_oFields.Add("StrretName");
  oOutputFile.m_oFields.Add("HouseNumber");
  oOutputFile.m_oFields.Add("PostalCode");
  oOutputFile.m_oFields.Add("CityName");
  oOutputFile.m_oFields.Add("CountryName");


  oOutputFile.m_oFields.Add("TypeOfRepresentation");


  oOutputFile.m_oFields.Add("ShareCapital_GmbH_TotalAmountOfMoney");
  oOutputFile.m_oFields.Add("ShareCapital_GmbH_OutdatedCurrency");
  oOutputFile.m_oFields.Add("ShareCapital_GmbH_ActualCurrency");


  oOutputFile.m_oFields.Add("ShareCapital_AG_TotalAmountOfMoney");
  oOutputFile.m_oFields.Add("ShareCapital_AG_ActualCurrency");


  oOutputFile.m_oFields.Add("BranchText");


  oOutputFile.SetSettings(';', '\"');
  oOutputFile.Create(sOutputFileName, eUTF16LE);
  oOutputFile.AppenHeaderRecord();

  for (
    citCompanyTradeRegisterData = v_coCompanyTradeRegisterDataList.begin();
    citCompanyTradeRegisterData != v_coCompanyTradeRegisterDataList.end();
    citCompanyTradeRegisterData++)
  {
    const ClCompanyTradeRegisterData& coCompanyTradeRegisterData = 
      *citCompanyTradeRegisterData;

    oOutputFile.m_oFields["LocalCourtId"].SetStringW(coCompanyTradeRegisterData.m_wsLocalCourtId);
    oOutputFile.m_oFields["LocalCourtName"].SetStringW(coCompanyTradeRegisterData.m_wsLocalCourtName);
    oOutputFile.m_oFields["TypeOfRegister"].SetStringW(coCompanyTradeRegisterData.m_wsTypeOfRegister);
    oOutputFile.m_oFields["RegistrationNumber"].SetStringW(coCompanyTradeRegisterData.m_wsRegistrationNumber);


    oOutputFile.m_oFields["RegistrationDate_FormattedDate"].SetStringW(
      coCompanyTradeRegisterData.m_wsRegistrationDate_FormattedDate);
    oOutputFile.m_oFields["RegistrationDate_UnformattedDate"].SetStringW(
      coCompanyTradeRegisterData.m_wsRegistrationDate_UnformattedDate);


    oOutputFile.m_oFields["CompanyName"].SetStringW(coCompanyTradeRegisterData.m_wsCompanyName);


    oOutputFile.m_oFields["LegalForm"].SetStringW(coCompanyTradeRegisterData.m_wsLegalForm);


    oOutputFile.m_oFields["CityName_WhereMainOfficeIsLocated"].SetStringW(
      coCompanyTradeRegisterData.m_wsCityName_WhereMainOfficeIsLocated);


    oOutputFile.m_oFields["PostalAddressType"].SetStringW(
      coCompanyTradeRegisterData.m_wsPostalAddressType);
    oOutputFile.m_oFields["StrretName"].SetStringW(coCompanyTradeRegisterData.m_wsStrretName);
    oOutputFile.m_oFields["HouseNumber"].SetStringW(coCompanyTradeRegisterData.m_wsHouseNumber);
    oOutputFile.m_oFields["PostalCode"].SetStringW(coCompanyTradeRegisterData.m_wsPostalCode);
    oOutputFile.m_oFields["CityName"].SetStringW(coCompanyTradeRegisterData.m_wsCityName);
    oOutputFile.m_oFields["CountryName"].SetStringW(coCompanyTradeRegisterData.m_wsCountryName);


    oOutputFile.m_oFields["TypeOfRepresentation"].SetStringW(
      coCompanyTradeRegisterData.m_wsTypeOfRepresentation);


    oOutputFile.m_oFields["ShareCapital_GmbH_TotalAmountOfMoney"].SetStringW(
      coCompanyTradeRegisterData.m_wsShareCapital_GmbH_TotalAmountOfMoney);
    oOutputFile.m_oFields["ShareCapital_GmbH_OutdatedCurrency"].SetStringW(
      coCompanyTradeRegisterData.m_wsShareCapital_GmbH_OutdatedCurrency);
    oOutputFile.m_oFields["ShareCapital_GmbH_ActualCurrency"].SetStringW(
      coCompanyTradeRegisterData.m_wsShareCapital_GmbH_ActualCurrency);


    oOutputFile.m_oFields["ShareCapital_AG_TotalAmountOfMoney"].SetStringW(
      coCompanyTradeRegisterData.m_wsShareCapital_AG_TotalAmountOfMoney);
    oOutputFile.m_oFields["ShareCapital_AG_ActualCurrency"].SetStringW(
      coCompanyTradeRegisterData.m_wsShareCapital_AG_ActualCurrency);


    oOutputFile.m_oFields["BranchText"].SetStringW(coCompanyTradeRegisterData.m_wsBranchText);


    oOutputFile.Append();

    if (oOutputFile.GetRecordNum() % 10000 == 0)
    {
      v_oJobProtocol << L".";
      v_oJobProtocol.flush();
    }
  }

  oOutputFile.Close();


  v_oJobProtocol << L"OK";
  v_oJobProtocol.endl();
}


void
SaveCompanyTradeRegisterDataListToFile(
  const ClInvalidFileNameList& v_coInvalidFileNameList,
  const w_string v_cwsOutputFileName,
  ClJobProtocol& v_oJobProtocol)
{
  v_oJobProtocol << L"Saving file \"" << v_cwsOutputFileName << L"\" (one dot per each 10 000 records): ";
  v_oJobProtocol.flush();


  ClInvalidFileNameList::const_iterator citInvalidFileName;

  ClFileDEL oOutputFile;
  ClCodePageExt oCPConverter;
  string sOutputFileName;

  sOutputFileName = oCPConverter.ConvertW_StringToString(v_cwsOutputFileName);


  oOutputFile.m_oFields.Add("InputFileName");

  oOutputFile.SetSettings(';', '\"');
  oOutputFile.Create(sOutputFileName, eUTF16LE);
  oOutputFile.AppenHeaderRecord();

  for (
    citInvalidFileName = v_coInvalidFileNameList.begin();
    citInvalidFileName != v_coInvalidFileNameList.end();
    citInvalidFileName++)
  {
    const w_string& cwsInvalidFileName = *citInvalidFileName;

    oOutputFile.m_oFields["InputFileName"].SetStringW(cwsInvalidFileName);

    oOutputFile.Append();

    if (oOutputFile.GetRecordNum() % 10000 == 0)
    {
      v_oJobProtocol << L".";
      v_oJobProtocol.flush();
    }
  }

  oOutputFile.Close();


  v_oJobProtocol << L"OK";
  v_oJobProtocol.endl();
}

void
ConvertData(
  w_string v_wsInputDirectory,
  w_string v_wsOutputDirectory,
  w_string v_wsErrorStatisticsDirectory,
  ClJobProtocol& v_oJobProtocol )
{
  ClCompanyTradeRegisterDataList oCompanyTradeRegisterDataList;

  ClInvalidFileNameList oRegistrationData_CannotBeExtracted_FileNameList;
  ClInvalidFileNameList oRegistrationDate_CannotBeExtracted_FileNameList;
  ClInvalidFileNameList oOtherDataOfCompany_CannotBeExtracted_FileNameList;
  ClInvalidFileNameList oPostalAddress_CannotBeExtracted_FileNameList;
  ClInvalidFileNameList oTypeOfRepresentation_CannotBeExtracted_FileNameList;
  ClInvalidFileNameList oShareCapital_CannotBeExtracted_FileNameList;
  ClInvalidFileNameList oBranchText_CannotBeExtracted_FileNameList;
  

  list<w_string> oFoundDirectoryNameList;
  list<w_string>::const_iterator citTRNumber_DirectoryFullName;

  FindAllDirectoriesInDirectory_UsingWildcards(
    v_wsInputDirectory,
    L"*",
    oFoundDirectoryNameList);

  for (citTRNumber_DirectoryFullName = oFoundDirectoryNameList.begin();
    citTRNumber_DirectoryFullName != oFoundDirectoryNameList.end();
    citTRNumber_DirectoryFullName++)
  {
    w_string wsFileNameMask;
    list<w_string> oFoundFileNameList;
    list<w_string>::const_iterator citTRFile_FullFileName;

    const w_string& wsTargetDirectory = *citTRNumber_DirectoryFullName;
    wsFileNameMask = L"*.xml";

    FindAllFilesInDirectory_UsingWildcards(
      wsTargetDirectory,
      wsFileNameMask,
      oFoundFileNameList);

    for (citTRFile_FullFileName = oFoundFileNameList.begin();
      citTRFile_FullFileName != oFoundFileNameList.end();
      citTRFile_FullFileName++)
    {
      const w_string& cwsTRFile_FullFileName = *citTRFile_FullFileName;
      ClCompanyTradeRegisterData oCompanyTradeRegisterData;

      LoadXMLFileForCompany(
        cwsTRFile_FullFileName, 
        oCompanyTradeRegisterData,
        oRegistrationData_CannotBeExtracted_FileNameList,
        oRegistrationDate_CannotBeExtracted_FileNameList,
        oOtherDataOfCompany_CannotBeExtracted_FileNameList,
        oPostalAddress_CannotBeExtracted_FileNameList,
        oTypeOfRepresentation_CannotBeExtracted_FileNameList,
        oShareCapital_CannotBeExtracted_FileNameList,
        oBranchText_CannotBeExtracted_FileNameList,
        v_oJobProtocol );

      oCompanyTradeRegisterDataList.push_back(oCompanyTradeRegisterData);
    }
  }

  v_oJobProtocol.endl();
  v_oJobProtocol.endl();

  SaveCompanyTradeRegisterDataListToFile(
    oCompanyTradeRegisterDataList,
    MakePathW(v_wsOutputDirectory, L"Companies.txt"),
    v_oJobProtocol);

  SaveCompanyTradeRegisterDataListToFile(
    oRegistrationData_CannotBeExtracted_FileNameList,
    MakePathW(v_wsErrorStatisticsDirectory, L"RegistrationData_CannotBeExtracted.txt"),
    v_oJobProtocol );

  SaveCompanyTradeRegisterDataListToFile(
    oRegistrationDate_CannotBeExtracted_FileNameList,
    MakePathW(v_wsErrorStatisticsDirectory, L"RegistrationDate_CannotBeExtracted.txt"),
    v_oJobProtocol);

  SaveCompanyTradeRegisterDataListToFile(
    oOtherDataOfCompany_CannotBeExtracted_FileNameList,
    MakePathW(v_wsErrorStatisticsDirectory, L"OtherDataOfCompany_CannotBeExtracted.txt"),
    v_oJobProtocol);

  SaveCompanyTradeRegisterDataListToFile(
    oPostalAddress_CannotBeExtracted_FileNameList,
    MakePathW(v_wsErrorStatisticsDirectory, L"PostalAddress_CannotBeExtracted.txt"),
    v_oJobProtocol);

  SaveCompanyTradeRegisterDataListToFile(
    oTypeOfRepresentation_CannotBeExtracted_FileNameList,
    MakePathW(v_wsErrorStatisticsDirectory, L"TypeOfRepresentation_CannotBeExtracted.txt"),
    v_oJobProtocol);

  SaveCompanyTradeRegisterDataListToFile(
    oShareCapital_CannotBeExtracted_FileNameList,
    MakePathW(v_wsErrorStatisticsDirectory, L"ShareCapital_CannotBeExtracted.txt"),
    v_oJobProtocol);

  SaveCompanyTradeRegisterDataListToFile(
    oBranchText_CannotBeExtracted_FileNameList,
    MakePathW(v_wsErrorStatisticsDirectory, L"BranchText_CannotBeExtracted.txt"),
    v_oJobProtocol);


  int i = 0;
}



int main()
{
  int iProgramError = 0;

  w_string wsLogFileDirectory;
  w_string wsInputDirectory;
  w_string wsOutputDirectory;
  w_string wsErrorStatisticsDirectory;


  ClJobProtocol oJobProtocol;


  try
  {
    // TestJobProtocol();
    // Test_CreateXMLFile();



    // Initialize program arguments.
    wsLogFileDirectory = L"C:\\Work\\2026.01.04\\Log";
    wsInputDirectory = L"C:\\Work\\2026.01.04\\Input";
    wsOutputDirectory = L"C:\\Work\\2026.01.04\\Output";
    wsErrorStatisticsDirectory = L"C:\\Work\\2026.01.04\\ErrorStatistics";

    // Create .log file.
    oJobProtocol.Assign(wsLogFileDirectory, L"DEU TR Converter", true);
    oJobProtocol.SetOnScreen(true);
    oJobProtocol.SetUnicodeForConsoleWindowIsEnabled(true);


    // Print program description.
    oJobProtocol << L"Program for convertation of .xml files" << L"\n";
    oJobProtocol << L"  (that were extarcted from German Trade Register web site)" << L"\n";
    oJobProtocol << L"  into Companies.txt and Persons.txt files." << L"\n";
    oJobProtocol << L"\n";
    oJobProtocol << L"  Version: 1.0, Build: 26010." << L"\n";
    oJobProtocol << L"\n";
    oJobProtocol << L"  Copyright (c) 2001 - 2026  ACS Informatik GmbH, Munich, Germany." << L"\n";
    oJobProtocol << L"\n";
    oJobProtocol << L"\n";
    oJobProtocol.flush();
  }
  catch (EGeneric& ex)
  {
    wcout << L"== Error ==\n";
    wcout << L"\n";
    wcout << L"Message (FLX::EGeneric): \"" << W_StrToWStr(ex.GetErrorMessageW()) << L"\"\n";
    wcout << L"\n";

    iProgramError = 1;
  }
  catch (std::exception& std_ex)
  {
    cout << "== Error std::exception ==\n";
    cout << "\n";
    cout << "Message (std::exception): \"" << std_ex.what() << "\"\n";
    cout << "\n";

    iProgramError = 2;
  }
  // catch (System::Exception^ clr_ex)
  // {
  //   wstring wwsMessage = ManagedStrToWString(clr_ex->Message);
  //
  //  wprintf(L"%s\n\"%s\"\n", L"System::Exception", wwsMessage.c_str());
  // }
  catch( ... )
  {
    wcout << L"== General error ==\n";
    wcout << L"\n";

    iProgramError = 4;
  }


  if (iProgramError == 0)
  {
    try
    {
      // TestJobProtocol();
      // Test_CreateXMLFile();

      ConvertData(
        wsInputDirectory,
        wsOutputDirectory,
        wsErrorStatisticsDirectory,
        oJobProtocol );


      oJobProtocol.endl();
      oJobProtocol << L"Processing was successfully completed.\n";
      oJobProtocol.endl();
      oJobProtocol.endl();
    }
    catch (EGeneric& ex)
    {
      oJobProtocol << L"== Error ==\n";
      oJobProtocol << L"\n";
      oJobProtocol << L"Message (FLX::EGeneric): \"" << W_StrToWStr(ex.GetErrorMessageW()) << L"\"\n";
      oJobProtocol << L"\n";

      iProgramError = 11;
    }
    catch (std::exception& std_ex)
    {
      oJobProtocol << "== Error std::exception ==\n";
      oJobProtocol << "\n";
      oJobProtocol << "Message (std::exception): \"" << std_ex.what() << "\"\n";
      oJobProtocol << "\n";

      iProgramError = 12;
    }
    // catch (System::Exception^ clr_ex)
    // {
    //   wstring wwsMessage = ManagedStrToWString(clr_ex->Message);
    //
    //  wprintf(L"%s\n\"%s\"\n", L"System::Exception", wwsMessage.c_str());
    // }
    catch (...)
    {
      oJobProtocol << L"== General error ==\n";
      oJobProtocol << L"\n";
      iProgramError = 14;
    }
  }

  oJobProtocol.Close();

  return iProgramError;
}
