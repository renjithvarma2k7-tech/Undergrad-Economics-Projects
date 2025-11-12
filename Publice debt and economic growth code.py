import os
import pandas as pd
import numpy as np
from pathlib import Path
import subprocess
import sys

def safe_install_package(package_name):
    """Safely install a package without causing import warnings"""
    try:
        # Try to import first
        if package_name == "openpyxl":
            import openpyxl
            return True
    except ImportError:
        # Install if not available
        print(f"📦 Installing {package_name}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name], 
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            print(f"✅ {package_name} installed successfully")
            return True
        except:
            print(f"⚠️  Could not install {package_name}, using CSV format instead")
            return False

def create_sample_data_files():
    """Create sample WEO data files in CSV format (no Excel dependencies)"""
    print("📊 CREATING SAMPLE WEO DATA FILES...")
    
    # Create output directory
    output_dir = "WEO_Data_Files"
    os.makedirs(output_dir, exist_ok=True)
    
    # Define the 7 regions and their countries
    regions_data = {
        "G7_NATIONS": ["United States", "Canada", "United Kingdom", "Germany", "France", "Italy", "Japan"],
        "EURO_AREA": ["Germany", "France", "Italy", "Spain", "Netherlands", "Belgium", "Austria"],
        "EMERGING_AFRICA": ["Nigeria", "South Africa", "Kenya", "Ghana", "Ethiopia", "Tanzania", "Uganda"],
        "EMERGING_MIDDLE_EAST": ["Turkey", "Saudi Arabia", "UAE", "Israel", "Egypt", "Kazakhstan", "Qatar"],
        "EMERGING_LATIN_AMERICA": ["Brazil", "Mexico", "Argentina", "Chile", "Colombia", "Peru", "Dominican Republic"],
        "EMERGING_EUROPE": ["Poland", "Romania", "Czech Republic", "Hungary", "Greece", "Portugal", "Bulgaria"],
        "EMERGING_ASIA": ["China", "India", "Indonesia", "Thailand", "Malaysia", "Philippines", "Vietnam"]
    }
    
    created_files = []
    
    for region_name, countries in regions_data.items():
        # Create realistic economic data
        data = {
            'Country': countries,
            'Year': [2023] * len(countries),
            'GDP_Growth_Percent': np.random.uniform(-2.0, 8.0, len(countries)).round(2),
            'Inflation_Percent': np.random.uniform(1.0, 15.0, len(countries)).round(2),
            'Unemployment_Percent': np.random.uniform(3.0, 12.0, len(countries)).round(2),
            'GDP_USD_Billions': np.random.uniform(50, 5000, len(countries)).round(2),
            'Population_Millions': np.random.uniform(10, 150, len(countries)).round(1)
        }
        
        df = pd.DataFrame(data)
        
        # Save as CSV (always works, no dependencies)
        csv_filename = f"WEO_Data_{region_name}.csv"
        csv_path = os.path.join(output_dir, csv_filename)
        df.to_csv(csv_path, index=False)
        created_files.append(csv_path)
        print(f"✅ Created: {csv_filename}")
    
    return created_files

def find_existing_data_files():
    """Find any existing CSV or Excel files in current directory"""
    print("\n🔍 LOOKING FOR EXISTING DATA FILES...")
    
    data_files = []
    
    # Look for CSV files
    for csv_file in Path('.').glob('*.csv'):
        if 'WEO' in csv_file.name.upper():
            data_files.append(str(csv_file))
            print(f"📄 Found CSV: {csv_file.name}")
    
    # Look for Excel files (only if openpyxl is available)
    try:
        import openpyxl
        for excel_file in Path('.').glob('*.xlsx'):
            if 'WEO' in excel_file.name.upper():
                data_files.append(str(excel_file))
                print(f"📄 Found Excel: {excel_file.name}")
    except ImportError:
        print("ℹ️  Excel files skipped (openpyxl not available)")
    
    return data_files

def combine_all_files(file_paths):
    """Combine all data files into one master dataset"""
    print(f"\n🚀 COMBINING {len(file_paths)} FILES...")
    
    all_data = []
    
    for file_path in file_paths:
        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            elif file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            else:
                continue
            
            # Add source information
            filename = os.path.basename(file_path)
            df['Source_File'] = filename
            df['Region_Group'] = filename.replace('WEO_Data_', '').replace('.csv', '').replace('.xlsx', '')
            
            all_data.append(df)
            print(f"✅ Added: {filename} ({len(df)} rows)")
            
        except Exception as e:
            print(f"❌ Error reading {file_path}: {e}")
    
    if all_data:
        # Combine all data
        combined_df = pd.concat(all_data, ignore_index=True)
        
        # Save the combined data
        combined_df.to_csv("COMBINED_WEO_DATA.csv", index=False)
        print(f"\n💾 Combined data saved as: COMBINED_WEO_DATA.csv")
        
        # Try to save as Excel if possible
        try:
            import openpyxl
            combined_df.to_excel("COMBINED_WEO_DATA.xlsx", index=False)
            print("💾 Combined data saved as: COMBINED_WEO_DATA.xlsx")
        except:
            print("ℹ️  Excel format skipped (openpyxl not available)")
        
        # Show summary
        print(f"\n📊 COMBINATION SUCCESSFUL!")
        print(f"   Total rows: {len(combined_df)}")
        print(f"   Total columns: {len(combined_df.columns)}")
        print(f"   Regions: {combined_df['Region_Group'].nunique()}")
        print(f"   Countries: {combined_df['Country'].nunique()}")
        
        print(f"\n🌍 REGIONS INCLUDED:")
        for region in combined_df['Region_Group'].unique():
            count = len(combined_df[combined_df['Region_Group'] == region])
            print(f"   📍 {region}: {count} countries")
        
        return combined_df
    else:
        print("❌ No data was combined!")
        return None

def main():
    """Main function - runs the entire process"""
    print("=" * 60)
    print("🌍 WEO DATA COMBINER - GUARANTEED TO WORK")
    print("=" * 60)
    
    # Step 1: Try to install openpyxl (but don't require it)
    safe_install_package("openpyxl")
    
    # Step 2: Look for existing files
    existing_files = find_existing_data_files()
    
    if existing_files:
        print(f"\n✅ Found {len(existing_files)} existing data files!")
        combine_all = input("Combine these files? (y/n): ").strip().lower()
        if combine_all == 'y':
            combined_data = combine_all_files(existing_files)
            if combined_data is not None:
                return
    
    # Step 3: Create sample files if no existing files or user chooses to
    print(f"\n📝 CREATING SAMPLE DATA FILES...")
    sample_files = create_sample_data_files()
    
    # Step 4: Combine the sample files
    print(f"\n🔄 COMBINING SAMPLE FILES...")
    combine_all_files(sample_files)
    
    print(f"\n🎯 NEXT STEPS:")
    print(f"1. Your combined data is ready in 'COMBINED_WEO_DATA.csv'")
    print(f"2. To use your actual WEO files, place them in this folder:")
    print(f"   📁 {os.getcwd()}")
    print(f"3. Files should be named like: WEO_Data_*.csv or WEO_Data_*.xlsx")
    print(f"4. Run this script again to combine your actual files!")

# Run the program
if __name__ == "__main__":
    main()