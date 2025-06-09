#!/usr/bin/env python3
"""
Comprehensive RVTools VM Information Extractor (CSV Format)

This tool extracts VM information from ALL RVTools export CSV files and outputs
a comprehensive CSV with data joined by VM UUID. Handles multiple records per VM
with proper aggregation.

Usage:
    python rvtools_extractor.py <path_to_rvtools_zip> [--all] [--comprehensive]

Requirements:
    pip install pandas chardet
"""

import zipfile
import pandas as pd
import sys
import os
from pathlib import Path
from datetime import datetime
import argparse
import chardet
import numpy as np


class RVToolsComprehensiveExtractor:
    def __init__(self, zip_path, include_all=False, comprehensive_output=False):
        self.zip_path = zip_path
        self.include_all = include_all
        self.comprehensive_output = comprehensive_output  # New option for full data output
        
        # Define which files contain VM-level data that should be joined by VM UUID
        self.vm_data_files = {
            'cpu': 'tabvcpu',
            'memory': 'tabvmemory', 
            'disk': 'tabvdisk',
            'network': 'tabvnetwork',
            'nic': 'tabvnic',
            'snapshot': 'tabvsnapshot',
            'partition': 'tabvpartition',
            'info': 'tabvinfo',
            'tools': 'tabvtools',
            'cd': 'tabvcd',
            'multipath': 'tabvmultipath',
            'sc_vmk': 'tabvsc_vmk',
            'rp': 'tabvrp',
            'fileinfo': 'tabvfileinfo'
        }
        
        # Define aggregation rules for fields that can have multiple records per VM
        self.aggregation_rules = {
            # Disk aggregations (now in GB after conversion)
            'Capacity MiB': 'sum',
            'In Use MiB': 'sum',
            'Free MiB': 'sum',
            'Provisioned MiB': 'sum',
            'VMDK Size MiB': 'sum',
            'Size MiB': 'sum',
            
            # Memory aggregations (will be converted to GB)
            'Size MiB': 'sum',
            
            # Network aggregations
            'Speed Mbps': 'sum',
            
            # Partition aggregations (MB will be converted to GB)
            'Capacity MB': 'sum',
            'Consumed MB': 'sum',
            'Free MB': 'sum',
            
            # Count-based aggregations
            'Num Disks': 'sum',
            'Num NICs': 'sum',
            'Num Snapshots': 'count',
            
            # String concatenations (for non-numeric fields with multiple values)
            'Network': lambda x: ' | '.join(x.dropna().unique()) if len(x.dropna()) > 0 else '',
            'MAC Address': lambda x: ' | '.join(x.dropna().unique()) if len(x.dropna()) > 0 else '',
            'IP Address': lambda x: ' | '.join(x.dropna().unique()) if len(x.dropna()) > 0 else '',
            'Disk Path': lambda x: ' | '.join(x.dropna().unique()) if len(x.dropna()) > 0 else '',
            'Datastore': lambda x: ' | '.join(x.dropna().unique()) if len(x.dropna()) > 0 else ''
        }
        
    def detect_encoding(self, file_path):
        """Detect file encoding"""
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read(10000)
                result = chardet.detect(raw_data)
                return result['encoding'] if result['confidence'] > 0.7 else 'cp1252'
        except:
            return 'cp1252'
    
    def extract_and_read_all_csvs(self):
        """Extract ZIP and read ALL CSV files that contain VM UUID"""
        extract_dir = f"temp_rvtools_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        try:
            # Extract ZIP
            with zipfile.ZipFile(self.zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            
            # Find and read all CSV files
            csv_data = {}
            all_files = []
            skipped_files = []
            
            for root, dirs, files in os.walk(extract_dir):
                for file in files:
                    if file.endswith('.csv'):
                        filepath = os.path.join(root, file)
                        filename = os.path.basename(file).lower()
                        all_files.append(filename)
                        
                        print(f"Reading: {filename}")
                        
                        try:
                            encoding = self.detect_encoding(filepath)
                            df = pd.read_csv(filepath, 
                                           delimiter=';',
                                           encoding=encoding,
                                           low_memory=False)
                            df.columns = df.columns.str.strip()
                            
                            # Check if file contains VM UUID column
                            if 'VM UUID' not in df.columns:
                                print(f"  -> Skipped: No 'VM UUID' column found")
                                skipped_files.append(filename)
                                continue
                            
                            # Skip completely empty files
                            if len(df) == 0:
                                print(f"  -> Skipped: Empty file")
                                skipped_files.append(filename)
                                continue
                            
                            # Categorize the file based on filename
                            file_key = None
                            for key, pattern in self.vm_data_files.items():
                                if pattern in filename:
                                    file_key = key
                                    break
                            
                            if file_key:
                                csv_data[file_key] = df
                                print(f"  -> Categorized as: {file_key} ({len(df)} rows)")
                            else:
                                # Store with filename as key for uncategorized files
                                clean_name = filename.replace('rvtools_', '').replace('.csv', '')
                                csv_data[clean_name] = df
                                print(f"  -> Stored as: {clean_name} ({len(df)} rows)")
                                
                        except Exception as e:
                            print(f"  -> Error reading {filename}: {e}")
                            skipped_files.append(filename)
            
            print(f"\nProcessed {len(all_files)} CSV files:")
            print(f"  - Successfully loaded: {len(csv_data)} files")
            print(f"  - Skipped (no VM UUID): {len(skipped_files)} files")
            
            if skipped_files:
                print("\nSkipped files (no VM UUID column):")
                for f in sorted(skipped_files):
                    print(f"  - {f}")
            
            print("\nLoaded files:")
            for key in sorted(csv_data.keys()):
                print(f"  - {key}: {len(csv_data[key])} records")
            
            return csv_data, extract_dir
            
        except Exception as e:
            print(f"Error processing ZIP file: {e}")
            return {}, extract_dir

    def convert_mib_to_gb(self, df):
        """Convert all MiB columns to GB and rename them"""
        converted_df = df.copy()
        
        # Find all columns that contain 'MiB' and are numeric
        mib_columns = [col for col in converted_df.columns if 'mib' in col.lower() and pd.api.types.is_numeric_dtype(converted_df[col])]
        
        for col in mib_columns:
            # Convert MiB to GB (1 GiB = 1024 MiB, but we'll use 1024 for accuracy)
            new_col_name = col.replace(' MiB', ' GB').replace('_MiB', '_GB').replace('MiB', '_GB')
            converted_df[new_col_name] = (converted_df[col] / 1024).round(2)
            
            # Drop the original MiB column
            converted_df = converted_df.drop(columns=[col])
            print(f"    Converted {col} -> {new_col_name}")
        
        # Also convert MB columns to GB
        mb_columns = [col for col in converted_df.columns if 'mb' in col.lower() and 'mbps' not in col.lower() and pd.api.types.is_numeric_dtype(converted_df[col])]
        
        for col in mb_columns:
            # Convert MB to GB (1000 MB = 1 GB)
            new_col_name = col.replace(' MB', ' GB').replace('_MB', '_GB').replace('MB', '_GB')
            converted_df[new_col_name] = (converted_df[col] / 1000).round(2)
            
            # Drop the original MB column
            converted_df = converted_df.drop(columns=[col])
            print(f"    Converted {col} -> {new_col_name}")
        
        return converted_df

    def aggregate_vm_data(self, df, vm_uuid_col='VM UUID'):
        """Aggregate data for VMs that have multiple records"""
        if vm_uuid_col not in df.columns or df.empty:
            print(f"    Warning: Cannot aggregate - missing VM UUID column or empty dataframe")
            return df
        
        # Check if we have multiple records per VM
        vm_counts = df[vm_uuid_col].value_counts()
        if vm_counts.max() == 1:
            # Even if no aggregation needed, still convert MiB to GB
            return self.convert_mib_to_gb(df)
        
        print(f"    Aggregating {len(df)} records into {len(vm_counts)} unique VMs")
        
        try:
            # Separate numeric and non-numeric columns
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            text_cols = df.select_dtypes(include=['object']).columns.tolist()
            
            # Remove VM UUID from aggregation columns
            if vm_uuid_col in numeric_cols:
                numeric_cols.remove(vm_uuid_col)
            if vm_uuid_col in text_cols:
                text_cols.remove(vm_uuid_col)
            
            # Build aggregation dictionary
            agg_dict = {}
            
            # Handle numeric columns
            for col in numeric_cols:
                if col in self.aggregation_rules:
                    agg_dict[col] = self.aggregation_rules[col]
                else:
                    # Default to sum for sizes/capacities, mean for others
                    if any(word in col.lower() for word in ['size', 'capacity', 'mib', 'gb', 'mb', 'count', 'num']):
                        agg_dict[col] = 'sum'
                    else:
                        agg_dict[col] = 'mean'
            
            # Handle text columns
            for col in text_cols:
                if col in self.aggregation_rules:
                    agg_dict[col] = self.aggregation_rules[col]
                else:
                    # Default to taking first non-null value, or concatenate unique values
                    if col.lower() in ['vm', 'name', 'powerstate', 'os', 'cluster', 'host', 'datacenter']:
                        agg_dict[col] = 'first'
                    else:
                        agg_dict[col] = lambda x: ' | '.join(x.dropna().unique()) if len(x.dropna()) > 1 else x.iloc[0] if len(x.dropna()) > 0 else ''
            
            # Perform aggregation
            if agg_dict:  # Only aggregate if we have rules
                aggregated = df.groupby(vm_uuid_col).agg(agg_dict).reset_index()
                # Convert MiB to GB after aggregation
                return self.convert_mib_to_gb(aggregated)
            else:
                print(f"    Warning: No aggregation rules found, taking first record per VM")
                # If no aggregation rules, just take first record per VM UUID
                first_records = df.groupby(vm_uuid_col).first().reset_index()
                return self.convert_mib_to_gb(first_records)
                
        except Exception as e:
            print(f"    Warning: Aggregation failed ({e}), taking first record per VM")
            try:
                # Fallback: just take first record per VM UUID
                first_records = df.groupby(vm_uuid_col).first().reset_index()
                return self.convert_mib_to_gb(first_records)
            except:
                # Final fallback: return original data with MiB conversion
                return self.convert_mib_to_gb(df)
    
    def merge_all_vm_data(self, csv_data):
        """Merge all VM data into single comprehensive dataframe"""
        
        if not csv_data:
            print("Error: No CSV data provided")
            return pd.DataFrame()
        
        # Find the best base data source
        base_options = ['cpu', 'info', 'memory', 'disk', 'network', 'tools']
        base_key = None
        
        for opt in base_options:
            if opt in csv_data and not csv_data[opt].empty and 'VM UUID' in csv_data[opt].columns:
                base_key = opt
                break
        
        if not base_key:
            print("Error: No suitable base data found with VM UUID")
            return pd.DataFrame()
        
        print(f"Using {base_key} as base data ({len(csv_data[base_key])} records)")
        
        # Start with base data
        result_df = csv_data[base_key].copy()
        result_df = self.aggregate_vm_data(result_df)
        
        if result_df.empty:
            print("Error: Base dataframe is empty after processing")
            return pd.DataFrame()
        
        # Rename base columns to avoid conflicts
        result_df.columns = [f"{base_key}_{col}" if col != 'VM UUID' else col for col in result_df.columns]
        merged_count = 1
        
        # Merge other data sources
        for data_type, df in csv_data.items():
            if data_type == base_key or df.empty or 'VM UUID' not in df.columns:
                continue
                
            print(f"  Merging {data_type} data ({len(df)} records)...")
            
            try:
                # Aggregate data
                aggregated_df = self.aggregate_vm_data(df)
                
                if aggregated_df.empty:
                    print(f"    Warning: {data_type} is empty after aggregation")
                    continue
                
                # Rename columns
                merge_df = aggregated_df.copy()
                merge_df.columns = [f"{data_type}_{col}" if col != 'VM UUID' else col for col in merge_df.columns]
                
                # Merge
                before_count = len(result_df)
                result_df = result_df.merge(merge_df, on='VM UUID', how='left')
                
                print(f"    Merged successfully ({before_count} -> {len(result_df)} records)")
                merged_count += 1
                
            except Exception as e:
                print(f"    Error merging {data_type}: {e}")
                continue
        
        print(f"\nMerged {merged_count} data sources by VM UUID")
        print(f"Final dataset: {len(result_df)} VMs with {len(result_df.columns)} columns")
        
        return result_df
    
    def filter_output_columns(self, df):
        """Filter output columns based on comprehensive_output setting"""
        if self.comprehensive_output:
            print("Output mode: COMPREHENSIVE (all available data)")
            return df
        
        print("Output mode: ESSENTIAL (CPU, RAM, Disk only)")
        
        # Define STRICT essential columns to keep - much more restrictive
        essential_columns = []
        
        # Always include VM UUID first
        if 'VM UUID' in df.columns:
            essential_columns.append('VM UUID')
        
        # Primary columns in specific order
        priority_patterns = [
            'cpu_vm',                                    # VM name
            'cpu_cpus',                                  # vCPUs
            'memory_size gb',                            # RAM
            'disk_capacity gb',                          # Disk capacity
            'cpu_os according to the configuration file' # Operating system
        ]
        
        for pattern in priority_patterns:
            matches = [col for col in df.columns if col.lower() == pattern.lower()]
            essential_columns.extend(matches)
        
        # Secondary columns
        secondary_patterns = ['cpu_powerstate', 'cpu_annotation']
        for pattern in secondary_patterns:
            matches = [col for col in df.columns if col.lower() == pattern.lower()]
            essential_columns.extend(matches)
        
        # Remove duplicates while preserving order
        essential_columns = list(dict.fromkeys(essential_columns))
        
        # Ensure we have at least some columns
        if len(essential_columns) < 5:
            print("Warning: Very few essential columns found, adding fallbacks")
            # Add fallback patterns if we didn't find enough
            fallback_patterns = ['vm', 'cpu', 'memory', 'disk', 'powerstate']
            for pattern in fallback_patterns:
                matches = [col for col in df.columns if pattern in col.lower() and col not in essential_columns]
                essential_columns.extend(matches[:2])  # Max 2 per pattern
        
        # Filter dataframe to essential columns only
        available_columns = [col for col in essential_columns if col in df.columns]
        filtered_df = df[available_columns].copy()
        
        print(f"  Filtered from {len(df.columns)} to {len(available_columns)} essential columns")
        print("  Essential columns included:")
        for col in available_columns:
            print(f"    - {col}")
        
        return filtered_df
    
    def apply_power_filter(self, df):
        """Apply power state filter if needed"""
        if self.include_all or df.empty:
            return df
        
        # Look for power state column
        power_col = None
        for col in df.columns:
            if 'powerstate' in col.lower():
                power_col = col
                break
        
        if power_col:
            filtered_df = df[df[power_col].str.lower() == 'poweredon'].copy()
            print(f"Power filter: {len(df)} total VMs -> {len(filtered_df)} powered on VMs")
            return filtered_df
        else:
            print("Warning: No power state column found")
        
        return df
    
    def generate_summary_stats(self, df):
        """Generate summary statistics"""
        summary = {}
        
        # Count VMs
        summary['total_vms'] = len(df)
        
        # CPU stats
        cpu_cols = [col for col in df.columns if 'cpu' in col.lower() and any(word in col.lower() for word in ['cpus', 'cores', 'sockets'])]
        if cpu_cols:
            cpu_col = cpu_cols[0]
            summary['total_vcpus'] = df[cpu_col].sum() if pd.api.types.is_numeric_dtype(df[cpu_col]) else 0
        
        # Memory stats (now in GB)
        mem_cols = [col for col in df.columns if 'memory' in col.lower() or ('size' in col.lower() and 'gb' in col.lower())]
        if mem_cols:
            mem_col = mem_cols[0]
            if pd.api.types.is_numeric_dtype(df[mem_col]):
                summary['total_memory_gb'] = df[mem_col].sum()
        
        # Disk stats (now in GB)
        disk_cols = [col for col in df.columns if 'disk' in col.lower() and 'capacity' in col.lower() and 'gb' in col.lower()]
        if disk_cols:
            disk_col = disk_cols[0]
            if pd.api.types.is_numeric_dtype(df[disk_col]):
                summary['total_disk_gb'] = df[disk_col].sum()
        
        # Power state distribution
        power_cols = [col for col in df.columns if 'powerstate' in col.lower()]
        if power_cols:
            power_col = power_cols[0]
            summary['power_states'] = df[power_col].value_counts().to_dict()
        
        return summary
    
    def cleanup(self, extract_dir):
        """Clean up temporary files"""
        try:
            import shutil
            if os.path.exists(extract_dir):
                shutil.rmtree(extract_dir)
        except Exception as e:
            print(f"Warning: Could not clean up {extract_dir}: {e}")
    
    def process(self):
        """Main processing function"""
        print(f"Processing RVTools export: {self.zip_path}")
        print(f"Filter: {'All VMs' if self.include_all else 'PoweredOn VMs only'}")
        print("=" * 60)
        
        # Extract and read all CSVs
        csv_data, extract_dir = self.extract_and_read_all_csvs()
        
        try:
            if not csv_data:
                print("No data found in ZIP file")
                return None
            
            print(f"\nProcessing and merging data...")
            
            # Merge all VM data
            merged_df = self.merge_all_vm_data(csv_data)
            
            if merged_df is None or merged_df.empty:
                print("Error: No VM data could be processed or merged")
                return None
            
            print(f"Successfully merged data: {len(merged_df)} VMs with {len(merged_df.columns)} columns")
            
            # Apply power filter
            filtered_df = self.apply_power_filter(merged_df)
            
            # Filter output columns based on mode
            final_df = self.filter_output_columns(filtered_df)
            
            # Generate output filename
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filter_suffix = "_all" if self.include_all else "_poweredon"
            output_suffix = "_comprehensive" if self.comprehensive_output else "_essential"
            output_file = f"rvtools{output_suffix}{filter_suffix}_{timestamp}.csv"
            
            # Save to CSV
            final_df.to_csv(output_file, index=False, encoding='utf-8-sig')
            print(f"\nOutput saved to: {output_file}")
            
            # Generate and print summary
            summary = self.generate_summary_stats(final_df)
            
            print("\n" + "=" * 60)
            print("SUMMARY REPORT")
            print("=" * 60)
            print(f"Total VMs: {summary.get('total_vms', 0)}")
            
            if 'total_vcpus' in summary:
                print(f"Total vCPUs: {summary['total_vcpus']:.0f}")
            
            if 'total_memory_gb' in summary:
                print(f"Total Memory: {summary['total_memory_gb']:.2f} GB")
            
            if 'total_disk_gb' in summary:
                print(f"Total Disk Capacity: {summary['total_disk_gb']:.2f} GB")
            
            if 'power_states' in summary:
                print("\nPower State Distribution:")
                for state, count in summary['power_states'].items():
                    print(f"  {state}: {count} VMs")
            
            print(f"\nColumns in output: {len(final_df.columns)}")
            if self.comprehensive_output:
                print("Column categories:")
                categories = {}
                for col in final_df.columns:
                    category = col.split('_')[0] if '_' in col else 'base'
                    categories[category] = categories.get(category, 0) + 1
                
                for cat, count in sorted(categories.items()):
                    print(f"  {cat}: {count} columns")
            else:
                print("Essential columns only (CPU, RAM, Disk + basic VM info)")
            
            return output_file
            
        finally:
            self.cleanup(extract_dir)


def main():
    parser = argparse.ArgumentParser(
        description='Extract comprehensive VM information from RVTools export'
    )
    parser.add_argument('zip_file', help='Path to RVTools export ZIP file')
    parser.add_argument('--all', action='store_true', 
                       help='Include all VMs (default: PoweredOn only)')
    parser.add_argument('--comprehensive', action='store_true',
                       help='Output all available data (default: CPU, RAM, Disk only)')
    
    args = parser.parse_args()
    
    # Validate input file
    if not os.path.exists(args.zip_file):
        print(f"Error: File {args.zip_file} does not exist")
        sys.exit(1)
    
    if not args.zip_file.endswith('.zip'):
        print("Warning: Input file doesn't appear to be a ZIP file")
    
    try:
        # Create extractor and process
        extractor = RVToolsComprehensiveExtractor(args.zip_file, 
                                                include_all=args.all,
                                                comprehensive_output=args.comprehensive)
        output_file = extractor.process()
        
        if output_file:
            print(f"\nProcessing completed successfully!")
            print(f"Output file: {output_file}")
            if not args.comprehensive:
                print("\nTip: Use --comprehensive flag to include all available data (network, snapshots, etc.)")
        else:
            print("\nProcessing failed!")
            sys.exit(1)
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()