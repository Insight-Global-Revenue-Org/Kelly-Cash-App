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