# RVTools to BOM Pricing Analysis

A two-step process to convert RVTools VM inventory data into detailed Oracle Cloud Infrastructure (OCI) pricing analysis.

## Overview

This toolkit consists of two Python scripts that work together:
1. **RVTools Data Extractor** - Processes RVTools ZIP exports into clean CSV format
2. **VM BOM Generator** - Converts VM specifications into detailed Oracle Cloud pricing

## Prerequisites

```bash
pip install pandas chardet openpyxl
```

## Quick Start

### Step 1: Extract RVTools Data
```bash
python rvtools_extractor.py rvtools_export.zip
```

### Step 2: Generate Pricing Analysis
```bash
python vm_bom.py rvtools_essential_poweredon_20250609_143022.csv --excel
```

## Detailed Usage

### Script 1: RVTools Data Extractor

**Purpose**: Converts RVTools ZIP export into clean, standardized CSV format suitable for pricing analysis.

#### Basic Usage
```bash
python rvtools_extractor.py <rvtools_export.zip>
```

#### Advanced Options
```bash
# Include all VMs (default: powered on only)
python rvtools_extractor.py rvtools_export.zip --all

# Output all available data (default: essential columns only)
python rvtools_extractor.py rvtools_export.zip --comprehensive

# Both options combined
python rvtools_extractor.py rvtools_export.zip --all --comprehensive
```

#### Output Files
- **Essential mode** (default): `rvtools_essential_poweredon_YYYYMMDD_HHMMSS.csv`
- **Comprehensive mode**: `rvtools_comprehensive_poweredon_YYYYMMDD_HHMMSS.csv`
- **All VMs**: Includes `_all` instead of `_poweredon` in filename

#### What It Does
- Extracts all CSV files from RVTools ZIP export
- Merges VM data by UUID across multiple tables
- Converts storage units (MiB/MB → GB) for consistency
- Aggregates multiple records per VM (disks, NICs, etc.)
- Filters to essential columns: VM Name, vCPUs, RAM, Disk, OS, Power State

### Script 2: VM BOM Generator

**Purpose**: Converts VM specifications into detailed Oracle Cloud Infrastructure pricing with cost breakdown.

#### Basic Usage
```bash
python vm_bom.py <csv_file>
```

#### Advanced Options
```bash
# Generate Excel report with detailed analysis
python vm_bom.py vm_inventory.csv --excel

# Enable debug output for troubleshooting
python vm_bom.py vm_inventory.csv --debug

# Both options combined
python vm_bom.py vm_inventory.csv --debug --excel
```

#### Pricing Components Calculated
- **Compute**: OCPU costs (1 OCPU = 2 vCPUs, minimum 1)
- **Memory**: RAM costs in GB
- **Storage**: Block volume storage costs
- **Storage Performance**: VPU costs (10 VPUs per GB)
- **OS Licensing**: Windows Server licensing (Linux is free)

#### Output Formats

**Console Report**:
- Executive summary with total costs
- Detailed VM breakdown table
- Component cost analysis
- Powered-off VMs summary

**Excel Report** (with `--excel` flag):
- **Cost Summary** sheet: VM-by-VM costs
- **Detailed Analysis** sheet: Component-level breakdown
- **Component Breakdown** sheet: Cost by component type
- **Powered Off VMs** sheet: Excluded VMs

## Complete Workflow Example

### 1. Extract RVTools Data
```bash
# Extract essential data for powered-on VMs only (recommended)
python rvtools_extractor.py customer_rvtools_export.zip

# Output: rvtools_essential_poweredon_20250609_143022.csv
```

### 2. Generate Pricing Analysis
```bash
# Create detailed pricing analysis with Excel export
python vm_bom.py rvtools_essential_poweredon_20250609_143022.csv --excel

# Output: 
# - Console report with summary
# - rvtools_essential_poweredon_20250609_143022_detailed_analysis.xlsx
```

### 3. Review Results
The console output provides immediate insights:
```
EXECUTIVE SUMMARY
Total VMs Analyzed: 45 (powered on)
Monthly Total Cost: €12,847.50
Annual Total Cost: €154,170.00
Average Cost per VM: €285.50/month
```

## Troubleshooting

### Common Issues

**"No valid VM specifications found"**
- Check CSV column mapping in debug mode: `--debug`
- Ensure RVTools export contains VM data
- Verify CSV is not corrupted

**"Cannot find columns for: ['vm_name', 'cpu_cpus']"**
- RVTools export may be incomplete
- Try using `--comprehensive` flag in Step 1
- Check available columns in debug output

**Excel export fails**
- Install openpyxl: `pip install openpyxl`
- Check file permissions in output directory

### Debug Mode
Enable debug output to see detailed processing:
```bash
python vm_bom.py vm_inventory.csv --debug
```

This shows:
- Column mapping detection
- VM processing details  
- OCPU calculations
- Pricing breakdowns per VM

## Pricing Information

**Current Oracle Cloud Pricing (EUR)**:
- OCPU: €0.0279/hour (€20.76/month)
- Memory: €0.00186/GB/hour (€1.38/GB/month)
- Block Storage: €0.023715/GB/month
- Storage VPUs: €0.001581/VPU/month
- Windows License: €0.08556/OCPU/hour (€63.66/OCPU/month)

**Assumptions**:
- 24/7 operation (744 hours/month)
- Standard block storage performance
- 10 VPUs per GB storage allocation

## File Compatibility

### Supported RVTools Exports
- RVTools 4.x ZIP exports
- CSV format with semicolon delimiters
- UTF-8 or Windows-1252 encoding

### Expected CSV Columns
The BOM generator automatically maps these common column patterns:
- VM Name: `cpu_vm`, `vm name`, `vm_name`
- vCPUs: `cpu_cpus`, `vcpu`, `cpus`
- Memory: `memory_size gb`, `mem_size gb`, `memory gb`
- Disk: `disk_capacity gb`, `disk_total gb`, `disk gb`
- OS: `cpu_os according to the configuration file`
- Power State: `cpu_powerstate`, `powerstate`

## Output Customization

### Essential vs Comprehensive Data
- **Essential** (default): VM Name, vCPUs, RAM, Disk, OS, Power State
- **Comprehensive**: All available RVTools data (network, snapshots, tools, etc.)

### Filtering Options
- **Powered On** (default): Only includes running VMs in cost calculations
- **All VMs**: Includes all VMs but excludes powered-off ones from pricing

## Support

For issues or questions:
1. Run with `--debug` flag for detailed output
2. Check that RVTools export is complete and recent
3. Verify CSV column headers match expected patterns
4. Ensure all required Python packages are installed