# -*- coding: utf-8 -*-
import pandas as pd
import json
import os

class SimpleExcelLoader:
    def __init__(self, excel_filename, map_filename):
        self.excel_filename = excel_filename
        self.map_filename = map_filename
        self.df = None
        
        print("Loader created for:")
        print(f"  Excel file: {excel_filename}")
        print(f"  Mapping file: {map_filename}")
        
        # Check if files exist
        self._check_files_exist()

    def _check_files_exist(self):
        """Check if both files exist before proceeding"""
        if not os.path.exists(self.excel_filename):
            raise FileNotFoundError(f"Excel file not found: {self.excel_filename}")
        
        if not os.path.exists(self.map_filename):
            raise FileNotFoundError(f"Mapping file not found: {self.map_filename}")
        
        print("✓ Both files exist")

    def _get(self):
        try:
            print("Reading Excel file...")
            self.df = pd.read_excel(self.excel_filename)
            print("✓ Excel file read successfully")
            print(f"  Number of rows: {len(self.df):,}")
            print(f"  Number of columns: {len(self.df.columns)}")
            print(f"  Column names: {list(self.df.columns)}")
        except Exception as e:
            print(f"✗ Error reading Excel file: {e}")
            raise

    def _modify(self):
        try:
            print("Loading mapping file...")
            with open(self.map_filename, 'r', encoding='utf-8') as f:
                mapping = json.load(f)
            
            print("✓ Mapping file loaded successfully")
            print(f"  Mapping keys: {list(mapping.keys())}")
            
            eng_columns = [col for col in self.df.columns if col in mapping.keys()]
            
            if not eng_columns:
                print("✗ No columns matched with mapping!")
                print(f"  Available columns: {list(self.df.columns)}")
                print(f"  Mapping keys: {list(mapping.keys())}")
                return
            
            print("Column mapping:")
            for eng_col in eng_columns:
                print(f"  {eng_col} -> {mapping[eng_col]}")
            
            self.df = self.df[eng_columns]
            self.df = self.df.rename(columns=mapping)
            
            print("✓ Columns modified successfully")
            print(f"  Final columns: {list(self.df.columns)}")
            
        except Exception as e:
            print(f"✗ Error in modification process: {e}")
            raise

    def load(self):
        print("\n" + "="*50)
        print("Starting data loading process...")
        print("="*50)
        
        self._get()
        self._modify()
        
        print("\n✓ Data loading completed successfully!")
        return self.df

# Usage with proper file paths
if __name__ == "__main__":
    try:
        # Use proper file paths - adjust these to your actual file locations
        EXCEL_FILE = 'data/input/students_1404.xlsx'  # Your actual Excel file path
        MAP_FILE = 'config/map2.json'  # Your actual mapping file path
        
        # Create directories if they don't exist
        os.makedirs('data/input', exist_ok=True)
        os.makedirs('config', exist_ok=True)
        
        print("Looking for files...")
        print(f"Excel file path: {os.path.abspath(EXCEL_FILE)}")
        print(f"Mapping file path: {os.path.abspath(MAP_FILE)}")
        
        loader = SimpleExcelLoader(EXCEL_FILE, MAP_FILE)
        df = loader.load()
        print("\nFirst 5 rows of data:")
        print(df.head())
        df.to_excel('out.xlsx')
        
    except FileNotFoundError as e:
        print(f"\n✗ File error: {e}")
        print("\nPlease make sure:")
        print("1. The Excel file exists in the correct location")
        print("2. The mapping JSON file exists")
        print("3. The file paths are correct")
        
    except Exception as e:
        print(f"\n✗ Unexpected error: {e}")
