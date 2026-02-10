# DEU Trade Register Web Extractor

Web scraping tool for extracting company data from the German Trade Register (Handelsregister).

## Features

- **Mode 1**: Range mode - Download XML files for a range of trade register numbers
- **Mode 2**: List mode - Download XML files for a specific list of trade register numbers  
- **Mode 3**: Table scraping mode - Scrape table data directly without downloading files

## Output

The tool extracts the following company information:
- Local Court ID and Name
- Type of Register (HRA/HRB)
- Registration Number and Date
- Company Name, Legal Form, City
- Postal Address
- Type of Representation
- Share Capital (GmbH and AG)
- Business Branch/Activity
- **Company Status** (registered, deleted, etc.)

## Requirements

- .NET 8.0 SDK
- Google Chrome browser
- Internet connection (for ChromeDriver auto-download)

## Configuration

Edit `DEU_TradeRegister_WebExtractor_Settings.xml`:
- Set `ProcessingMode`: 1, 2, or 3
- Configure output paths
- Set range or list of register numbers

## Running

```bash
dotnet run
```

## Notes

- ChromeDriver is automatically managed by Selenium Manager
- Browser restart interval can be configured to avoid session timeouts
- Proxy support available in Mode 3

## Recent Fixes

- ✅ Fixed Status column parsing in Mode 3 (table scraping)
- ✅ Updated to Selenium WebDriver 4.28.0
- ✅ Improved ChromeDriver version management
