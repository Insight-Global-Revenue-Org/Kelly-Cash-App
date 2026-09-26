# Reconciliation Model Structure

```text
KellyCashApp
├── Configuration
│   └── Settings.cs
├── Models
│   ├── MicrosoftVmsMatch.cs
│   └── OirMatch.cs
├── Processors
│   ├── Allegis
│   │   ├── CDWPayment.cs
│   │   ├── CushmanWakefieldPayment.cs
│   │   ├── MicrosoftPayment.cs
│   │   ├── MicrosoftVms.cs
│   │   └── SamsungPayment.cs
│   ├── Guidant
│   │   └── GuidantPayment.cs
│   ├── Leidos
│   │	└── LeidosPayment.cs
│   ├── Kelly Services
│   │   ├── JohnsonJohnsonPayment.cs
│   │   └── KellyPayment.cs
│   ├── Monument
│   │   └── MonumentPayment.cs
│   └── Randstad
│       ├── NikeTracker.cs
│       └── RandstadPayment.cs
├── Reporting Workflows
│   ├── OIR.cs
│   └── UAC.cs
├── Services
│   ├── Analytics.cs
│   ├── ConsoleUi.cs
│   ├── FileSelector.cs
│   ├── OirImporter.cs
│   └── Rename.cs
├── Program.cs
└── Settings.cs
```

# Reconciliation Model — Development Log

> *This log contains the extracted commit metadata throughout this models development as I worked on it through the Summer of 2026. Proprietary source code, business data, credentials, internal file locations, and confidential implementation details are intentionally excluded. This Documentation Writeup contains no programmatic solution, no source code, and no reasonable identifiers for propriatary information.*
>
> **May 2026 — September 2026**

---

## September 2026

### September 16, 2026

- `4187c33` — Updated `ExtractContractorName` helper function
- `a5b1218` — Added and tested contractor-name regex parsing
- `55f4b27` — Updated column widths and corrected currency formatting
- `1107885` — Added payment reference number support, refactored worksheet output, and implemented fallback end-client string matching
- `2ba7202` — Added comma-delimited support for multiple identifier values
- `9031f99` — Added Experis reporting path configuration and end-client identifier matching
- `c8822dc` — Added Experis payment processor with EML/HTML remittance parsing
- `a0ad9bc` — Restricted Nike invoice index-matching parameters

---

## August 2026

### August 27, 2026

- `6c28f23` — Updated J&J payment processor candidate-matching logic

### August 26, 2026

- `239827a` — Expanded candidate matching pool for VMS fee-month matching
- `91bcac0` — Updated fee-row formatting behavior
- `a408593` — Added Kenvue payment processor with fee-description parsing, OIR index matching, and payment routing

### August 20, 2026

- `3d3f8d6` — Increased payment matching tolerance after restricting date/time matching parameters
- `ca1e9dd` — Added nullable `FeeMonth` storage to J&J VMS match model and updated dictionary matching parameters

### August 19, 2026

- `bf1b2f5` — Replaced iterative string matching with an explicit pattern array
- `7c063b8` — Refactored `ExtractWorkerName` into a sequence of maintainable explicit parsing patterns
- `0b8aaea` — Updated regex parsing to explicitly recognize month and four-digit year combinations
- `4a1ef6f` — Improved regular-expression parsing and null-value formatting behavior
- `77f71d0` — Added J&J VMS matching with 5% invoice tolerance and conditional formatting
- `341c8de` — Added Johnson & Johnson VMS index matching with in-memory contractor-name lookup

### August 16, 2026

- `d5b0613` — Added CenterPoint payment processor and initial VMS integration structure
- `5750a98` — Refactored Microsoft payment processing to automatically load the configured VMS report
- `e3d6045` — Refactored internal helper naming and column-width configuration
- `084539d` — Added reusable UAC notation-column width helper
- `e980a8b` — Improved Unapplied Cash column normalization and standardized column widths

### August 10, 2026

- `2112571` — Updated Randstad OIR index to support multiple invoices in memory
- `cfef8c5` — Improved invoice-grouping logic in the Randstad payment processor

---

## July 2026

### July 14, 2026

- `0851c47` — Updated project file-tree documentation
- `42eeda9` — Added exception handling to regex parsing workflow
- `c345ff3` — Restructured PDF format identifiers for Randstad and Leidos processors
- `255e925` — Added Leidos payment processor with initial regex-based parsing
- `eec65a0` — Moved project README to repository root
- `454b744` — Added project README documentation
- `5485193` — Removed temporary Monument diagnostic logging
- `8b41db5` — Added ±2-day and 10% variance fallback for Monument invoice dictionary matching
- `32ffa76` — Updated workflow to rediscover source columns after worksheet column insertion
- `6551c1f` — Separated Kelly payment aggregates into labor and expense categories using tuple-based grouping
- `d5760c1` — Enforced exact expense matching and added duplicate-invoice assignment protection
- `e274766` — Refactored `PauseAndReset` functionality into the UAC reporting workflow
- `a32cc10` — Refactored `PauseAndReset` functionality into the OIR reporting workflow

### July 12, 2026

- `2ca90a7` — Additional workflow cleanup and code documentation
- `f2d5b56` — Removed manual invoice-list positioning from OIR/UAC workflows
- `79d2494` — Completed integration of the new Console UI service into OIR and Full Cash workflows
- `0e8f1a1` — Split Kelly labor and expense invoice-matching rules
- `b543665` — Moved loading-display execution to the primary application thread
- `d22e690` — Delegated standard Kelly remittance workflow to the `KellyPayment` processor

### July 10, 2026

- `80585f0` — Converted row identifiers to dynamic values

### July 7, 2026

- `5971f4b` — Completed Console UI service cleanup
- `466f8e3` — Added persistent imported OIR status across application menus
- `d2a0530` — Added workbook and stream disposal before returning to the application menu
- `548c989` — Added fallback invoice handling and corrected date parsing
- `2644572` — Removed redundant diagnostic logging and improved code documentation
- `3f4cdb4` — Added two-pass expense matching with ±7-day range and 5% tolerance fallback
- `50473f3` — Corrected Hours Type sorting and added ordered output-row processing
- `03a02d2` — Added dynamic header-row detection
- `909c90f` — Added delegated Guidant processor with cached session import and conditional routing
- `7bd1bbc` — Refactored console rendering into a dedicated `ConsoleUi` helper and resolved redraw artifacts

### July 6, 2026

- `d246d13` — Updated application to .NET 10

---

## June 2026

### June 25, 2026

- `1a6bc83` — Added merge rules to combination-based invoice matching
- `66489f4` — Simplified combination matching by removing full-group matching attempt
- `059a0d9` — Added multiple-invoice detection to combination matching
- `d9b0386` — Added grouped SOW invoice matching with combination fallback
- `637279a` — Updated client-project matching restrictions
- `5368dbe` — Added client-project support
- `36f99ed` — Added Nike Tracker client-project matching for Randstad payments
- `74523d1` — Added Beeline identifier support to Randstad payment headers

### June 24, 2026

- `eadacde` — Added ±2-day tolerance for fuzzy matching in primary Kelly reconciliation logic

### June 18, 2026

- `9170054` — Updated project configuration to .NET 10

### June 17, 2026

- `af04949` — Updated Randstad parsing to support negative credit line items and ±2-day OIR matching tolerance

### June 16, 2026

- `d776bde` — Added delegated Randstad payment processor with PDF parsing support
- `24c42de` — Added CDW, Samsung, and Cushman & Wakefield payment-processing classes and application routing
- `b985cac` — Updated Microsoft customer header identification

### June 15, 2026

- `c0cedd1` — Adjusted threshold used to identify likely applied/closed transactions
- `f42801b` — Added dynamic header-row detection for new Open Invoice Reports

### June 11, 2026

- `9b1f12c` — Added delegated Microsoft VMS processor with cached session import
- `d9fd881` — Corrected Notes column width following sales-tax processing
- `2611fa5` — Corrected sales-tax calculation when retrieving invoices from the OIR dictionary
- `92b92a8` — Updated pageable interface page size

### June 10, 2026

- `731db0b` — Added conditional Allegis identification to primary remittance workflow
- `37cb756` — Resolved recurring console UI rendering artifacts

### June 8, 2026

- `81ca2a3` — Created Allegis namespace and Microsoft payment processor with duplicate-key protection for OIR dictionary indexing

### June 7, 2026

- `57c20bb` — Added local application-storage directory
- `d61bedf` — Restructured application into Models, Services, Reporting Workflows, Configuration, and Processors
- `7c843bc` — Replaced ClosedXML OIR import with a delegated streaming OpenXML parser

### June 6, 2026

- `ab31b3a` — Corrected Monument aggregate matching to consolidate split line-item payments

### June 4, 2026

- `547842d` — Corrected settings-menu UI rendering artifacts
- `e06e3de` — Added Monument payment processor
- `b6c96d3` — Added delegated Monument processor with conditional application routing

### June 3, 2026

- `d56cdd1` — Expanded internal code documentation
- `f962cb5` — Removed redundant application logging
- `db43d85` — Added configurable file selection and Windows File Explorer integration
- `fbeee06` — Improved menu rendering and selection behavior
- `a2974ac` — Corrected string processing in delegated analytics service
- `131708b` — Added OIR, UAC, and remittance analytics logging with persistent AppData storage and pageable history views
- `a675d88` — Added persistent contractor-name mapping and configurable settings submenu

---

## May 2026

### May 9, 2026

- `1e15512` — Added Johnson & Johnson fixed-fee remittance processing
- `46b10d8` — Added delegated UAC pull-forward workflow
- `ee33c37` — Added delegated OIR pull-forward workflow

### May 5, 2026

- `2764e3d` — Expanded documentation across primary processing workflows
- `1913214` — Refactored console UI rendering, added input handling, and persisted OIR data in memory
- `7b07f06` — Added Name and Concat matching keys and corrected worksheet column-index shifting

### May 2, 2026

- `4e5554b` — Corrected string handling and expanded code documentation
- `fc604d2` — Added dynamically generated output filenames using location description and total remittance amount
- `6da2f75` — Added dynamic row counting and payment-row aggregation
- `7a2eff3` — Added column normalization to standardize remittance formats before aggregation
- `b48b98c` — Corrected invoice line-item calculations
- `f433896` — Added payment-details output formatting

---

