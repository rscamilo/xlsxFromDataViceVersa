# js-xlsx-client-side: Client-Side XLSX Generation and Parsing in JavaScript

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)  This library provides two powerful JavaScript functions for working with XLSX (Excel) files directly in the browser, **without any server-side dependencies**.  It leverages `JSZip` for handling the XLSX file structure. This allows for offline functionality, improved performance, and reduced server load.

## Features

*   **`xlsxFromData(data, templateBase64 = null)`:** Convert a JavaScript array to a Base64-encoded XLSX file.
    *   Supports optional XLSX templates for preserving formatting, styles, and formulas.  This allows you to create beautifully formatted reports.
    *   Intelligently appends data to existing templates, handling rows that may already contain data.  New data is added *after* existing data.
    *   Uses chunking to efficiently process very large datasets, preventing memory issues.
*   **`dataFromXlsx(xlsxBase64, options = {})`:** Parse a Base64-encoded XLSX file into a JavaScript array.
    *   Handles shared strings (a common XLSX optimization) correctly.
    *   Provides flexible options for automatically converting Excel date/time values to JavaScript `Date` objects and formatting them into user-friendly strings.  You can specify date/time columns by:
        *   **Column index (number):**  e.g., `1` (for the first column), `2` (for the second), etc.
        *   **Column letter (string):** e.g., `"A"`, `"B"`, `"AB"`, etc.
        *   **Column header name (string):** e.g., `"Date of Birth"`, `"Timestamp"`, `"Order Date"`.  This is case-sensitive.
    *   Automatically removes empty rows and columns, providing clean and concise data.

## Why Client-Side?

*   **Offline Functionality:** Your application can generate and parse XLSX files even without an internet connection.  This is perfect for field work, mobile applications, or situations with unreliable connectivity.
*   **Performance:** Processing happens directly in the user's browser, eliminating server round-trips and providing a much faster user experience.  No waiting for uploads and downloads!
*   **Reduced Server Load:** Your server doesn't need to handle the complex task of XLSX processing, freeing up resources and reducing hosting costs.
*   **Enhanced User Experience:** Instant feedback and a responsive interface make your application feel smoother and more interactive.
*   **Data Privacy:** Sensitive data never leaves the user's browser, enhancing privacy and security.

## Use Cases

*   **Reporting and Dashboards:** Generate dynamic reports and export data directly from web applications. Users can download data in a familiar format.
*   **Data Entry Forms:** Export form data to XLSX for easy sharing, analysis, and integration with other tools.
*   **Offline Applications:** Enable spreadsheet interaction in offline-capable web apps, perfect for scenarios where internet access is limited.
*   **Data Migration:** Facilitate data transfer between web applications and spreadsheet-based systems.  Import and export data seamlessly.
*   **Data Analysis:**  Allow users to upload their own XLSX files for analysis and visualization within your web application.

## Installation

This is a client-side library, so you don't need a package manager like npm (although you *could* publish it to npm for convenience!). Simply include `index.js` in your HTML:

<script src="index.js"></script>
<script src="[https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js](https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js)"></script>

