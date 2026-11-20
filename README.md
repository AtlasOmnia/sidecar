# Sidecar by Atlas Omnia

Bring your own Hermes agent into native Microsoft Excel, Word, and PowerPoint.

This repository contains three independently installable Hermes plugins and their native Office add-ins:

- [`excel/`](excel/) — Sidecar for Excel
- [`word/`](word/) — Sidecar for Word
- [`powerpoint/`](powerpoint/) — Sidecar for PowerPoint

Start with the README inside the app you want to install. Each app includes its own plugin manifest, setup instructions, tests, and Apache-2.0 license.

## Sidecar for LibreOffice

Using LibreOffice instead? [Sidecar for LibreOffice](https://github.com/AtlasOmnia/sidecar-libreoffice) lives in its own repository. Hermes reads and edits the Calc, Writer and Impress documents you already have open, with an optional chat panel in the LibreOffice sidebar. It runs on macOS and Windows and is installed and updated separately from the Microsoft Office plugins here.

## Scope

Sidecar connects native desktop Office applications to the Hermes profile where its matching plugin is enabled. It does not send email, manage invitations, or provide a universal custom undo system; use the native Office Undo control where applicable.

## License and affiliation

Each component is licensed under Apache-2.0. Sidecar by Atlas Omnia is an independent project and is not affiliated with or endorsed by Microsoft or Nous Research. Microsoft Office product names are trademarks of Microsoft Corporation. LibreOffice is a registered trademark of The Document Foundation; Sidecar for LibreOffice is likewise not affiliated with or endorsed by The Document Foundation.
