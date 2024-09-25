# mkmgo-momworks

A collection of APIs designed to assist with data filtering, mapping, and processing. This project is specifically created to help my mom with her work.

## API Overview

### 1. Export Sasaran Kejar

**Endpoint:** `POST /momworks/sasaran/kejar`

#### Request

You can use the following `curl` command to export the sasaran kejar data:

```bash
curl --location 'http://localhost:8080/momworks/sasaran/kejar' \
--form 'myFile=@"postman-cloud:///xxxxxxx-xxxxxxxx-xxxxxxxx-xxxxxxxx"' \
--form 'sheetName="Sheet1"' \
--form 'kejarType="baduta"'
```

#### Parameters

- **myFile**: The Excel file to be uploaded (must be in `.xlsx` format).
- **sheetName**: The name of the sheet in the Excel file to process (e.g., `"Sheet1"`).
- **kejarType**: Specify either `"bayi"` (infant) or `"baduta"` (toddler) based on the data in the Excel file.

#### Description

This API accepts an uploaded `imunisasi` (immunization) Excel file for either `bayi` or `baduta`. It filters and validates the data, then creates a new Excel file for `sasaran kejar`. The exported file contains the data of `bayi` or `baduta` who have yet to receive complete immunization.

--- 