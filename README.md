# CTS Cloud Waste & Governance Report (GUI Edition)

A PowerShell + Windows Forms tool for identifying cost waste and governance-tagging gaps across **Azure Government** subscriptions. It scans virtual machines, orphaned managed disks, snapshots, and orphaned public IP addresses, estimates their monthly cost, and presents everything in a tabbed dark-mode dashboard with CSV export.

![Version](https://img.shields.io/badge/version-2.0-blue) ![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE) ![Cloud](https://img.shields.io/badge/cloud-Azure%20Government-0089D6)

---

## Features

- **Tabbed GUI** — separate views for VMs, Orphan Disks, Snapshots, and Orphan IPs, each with color-coded waste flags.
- **Cost estimation** — pulls live VM pricing from the Azure Retail Prices API and applies built-in disk/snapshot/IP rate tables.
- **Governance tagging audit** — checks each VM for `CKID`, `EMASS`, and `VASI` tags (with common casing variants) and classifies it as *Fully Tagged*, *Billing Only*, *ATO Only*, or *Untagged*.
- **Waste detection** — flags deallocated/stopped VMs, unattached disks, stale or source-less snapshots, and unassigned static IPs.
- **Live summary panel** — running totals for CKID coverage, compute vs. disk spend, and total estimated monthly waste.
- **CSV export** — writes one file per resource category for reporting or ticketing.
- **Offline pricing mode** — `-SkipPricing` skips the Retail Prices API for faster runs or air-gapped environments.

---

## Requirements

- Windows with **PowerShell 5.1+** (Windows Forms is required, so this does not run on PowerShell Core / Linux).
- Azure PowerShell modules:
  - `Az.Compute`
  - `Az.Network`
  - `Az.Accounts` (dependency)
- Read access (e.g. **Reader** role) on the target Azure Government subscriptions.
- Outbound HTTPS to `prices.azure.com` for live pricing (not needed with `-SkipPricing`).

Install the modules if needed:

```powershell
Install-Module Az.Compute, Az.Network -Scope CurrentUser
```

---

## Usage

```powershell
# Standard run with live pricing
.\Get-CloudWasteReport-GUI.ps1

# Skip the Retail Prices API (faster / offline)
.\Get-CloudWasteReport-GUI.ps1 -SkipPricing
```

On launch the script connects to **Azure Government** (`AzureUSGovernment`), enumerates your enabled subscriptions, and opens the GUI.

1. Pick a subscription from the dropdown.
2. Click **SCAN**.
3. Review the tabs and the summary panel.
4. Click **EXPORT CSV** to save results.

> If you are not already signed in, an interactive `Connect-AzAccount` prompt appears. If your current context is pointed at a non-Government cloud, the script disconnects and reconnects to Azure Government automatically.

---

## What Gets Scanned

| Tab | What it shows | Waste logic |
|-----|---------------|-------------|
| **VMs** | Every VM with size, OS, power state, tags, disk count/size, and cost breakdown | `WASTE` when deallocated/stopped (still billed for disks); `CHECK` for other non-running states |
| **Orphan Disks** | Managed disks not attached to any VM | Always `WASTE` |
| **Snapshots** | All snapshots with age and whether the source disk still exists | `WASTE` if source is gone or age > 90 days; otherwise `REVIEW` |
| **Orphan IPs** | Public IPs with no IP config and no NAT gateway | `WASTE` if static (billed); `REVIEW` if dynamic |

**Actual cost model:** `Actual = Compute (only if running) + Disk (always billed)`. Deallocated VMs still incur disk charges, which is why they surface as waste.

---

## Governance Tags

The tool looks for these tags on each resource and tolerates common casing/spelling variants:

- **CKID** — `CKID`, `ckid`, `CkId`, `Ckid`
- **EMASS** — `EMASS`, `eMASS`, `EMASNumber`, `EMASS_Number`, etc.
- **VASI** — `VASI`, `vasi`, `VASINumber`, `VASI_Number`, etc.

VM governance status:

| Status | Meaning |
|--------|---------|
| Fully Tagged | CKID **and** EMASS present |
| Billing Only | CKID present, EMASS missing |
| ATO Only | EMASS present, CKID missing |
| Untagged | Neither present |

---

## Cost Data Notes

- **VM pricing** is retrieved live from the Azure Retail Prices API, filtered to on-demand (non-Spot, non-Low-Priority) rates, and multiplied by 730 hours/month. Windows vs. Linux is chosen from the OS disk type.
- **Disk, snapshot, and static-IP costs** use built-in rate tables in the script and are **estimates**. Review these tables (`$script:DiskTierRates`, `Get-SnapshotMonthlyCost`, `$script:StaticIPMonthlyCost`) and update them to match your commercial/agreement pricing.
- All figures are approximations intended for prioritization, not billing reconciliation.

---

## Output Files

Exporting writes up to four CSVs alongside the name you choose (empty categories are skipped):

```
CloudWaste_<Subscription>_<YYYY-MM-DD>_VMs.csv
CloudWaste_<Subscription>_<YYYY-MM-DD>_OrphanDisks.csv
CloudWaste_<Subscription>_<YYYY-MM-DD>_Snapshots.csv
CloudWaste_<Subscription>_<YYYY-MM-DD>_OrphanIPs.csv
```

---

## Notes & Limitations

- Currently targets the **Azure Government** environment; change `Connect-AzAccount -Environment` if you need commercial Azure.
- Scans one subscription per **SCAN** click. Switch the dropdown and re-scan for others.
- Built-in rate tables drift over time — validate periodically.
- Requires a desktop session (Windows Forms); it will not run headless.

---

## Project Structure

```
.
├── Get-CloudWasteReport-GUI.ps1   # Main GUI script
└── README.md
```

---

## License

Add your license here (e.g. MIT). Create a `LICENSE` file in the repo root.
