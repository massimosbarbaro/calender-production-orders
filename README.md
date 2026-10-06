# Calandre: production orders and running conditions of calendering lines

*Commesse di produzione e condizioni di marcia delle calandre*

**Visual Basic 6** · 2002 · version 1.0.0  
Author: **Massimo Sbarbaro** ([ORCID 0009-0006-8965-9013](https://orcid.org/0009-0006-8965-9013))

## Overview

Calandre follows the production jobs running on the calenders of a plastic-film plant. For every job it records order, product, customer specification, finish, thickness, height, length, kg per metre, reels, beam, work centre, packaging, shift, start and end time and operator. It also stores the customer and product specifications and the machine running conditions used for each product. Orders are imported from the MFG/PRO ERP.

I designed and programmed this application in 2002. It is published here in 2026 as part of the documented record of my software work, with a permanent DOI on Zenodo.

## Main functions

- Job in production with all technical parameters (`frmID`).
- Customer specifications and product specifications (`frmCliente`, `frmProdotto`).
- Machine running conditions per product (`frmMarcia`).
- Import of production orders from the ERP through Reflection and FTP (`frmImportMFG`).

## Data

Microsoft Access database (`Calandre.mdb`) through DAO.

## Technology

DAO 3.51, Microsoft Access 8 object library, Crystal Reports, Winsock, Common Controls, Tab control, Masked Edit, DBGrid.

## Repository contents

| Path | Content |
|---|---|
| `src/` | Visual Basic 6 project (`.vbp`), forms (`.frm` with their binary resources `.frx`) and modules (`.bas`), as listed in the project file. |
| `config-example/` | Templates of the `.ini` configuration files read at start-up, with placeholder values. |

## What is not included

Crystal Reports layouts (`.rpt`), compiled executables, installers, scripts for the host connection and the production databases are **not** included, because they contain operational data of the organisations for which the program was built. Configuration files are published as templates in `config-example/` with placeholder values: server addresses, accounts and passwords have been removed. Names of client organisations and products have been removed from captions and comments; local paths have been normalised.

## Related repositories

- [compound-batch-dosing-vb6](https://github.com/massimosbarbaro/compound-batch-dosing-vb6)
- [warehouse-rack-locations-vb6](https://github.com/massimosbarbaro/warehouse-rack-locations-vb6)

## How to cite

Use the citation metadata in [`CITATION.cff`](CITATION.cff) (GitHub: *Cite this repository*). Each release is archived on Zenodo with its own DOI.

> Sbarbaro, Massimo. *Calandre: production orders and running conditions of calendering lines (Visual Basic 6, 2002)*. Software, version 1.0.0. GitHub: https://github.com/massimosbarbaro/calender-production-orders-vb6

## License

Released under the [MIT License](LICENSE). © 2002 Massimo Sbarbaro.
