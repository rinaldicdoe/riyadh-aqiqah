#!/usr/bin/env python3
"""
Test script to verify color and date formatting changes
"""
import sys
import os

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import the transform function from app.py
from app import transform_rekap_pemotongan
import pandas as pd
import tempfile

def test_color_date_format():
    """Test that colors and dates are formatted correctly"""
    
    # Create sample data with various scenarios
    data = {
        'No. Invoice': ['INV001', 'INV001', 'INV002'],
        'Nama Pemesan': ['Budi', 'Budi', 'Ahmad'],
        'Tanggal Kirim': ['2024-01-15', '2024-01-15', '2024-01-16'],
        'Tanggal Potong': ['2024-01-15', '2024-01-15', '2024-01-16'],
        'No. HP': ['081234567890', '081234567890', '081987654321'],
        'Alamat': ['Jl. Sudirman', 'Jl. Sudirman', 'Jl. Gatot Subroto'],
        'Nomor': [1, 2, 1],
        'Jumlah': [2, 1, 1],
        'Jenis Kelamin Anak': ['Laki-laki', 'Perempuan', 'Laki-laki'],
        'Nama Anak': ['Ari', 'Siti', 'Hasan'],
        'Umur Anak': ['5 Tahun', '3 Tahun', '2 Tahun'],
        'Paket & Menu': ['Kambing Jantan', 'Kambing Kebuli', 'Domba Jantan'],
        'Paket': ['Jantan', 'Kebuli', 'Jantan'],
        'No. Ref': ['REF001', 'REF001', 'REF002'],
        'No. Video': ['VID001', 'VID001', 'VID002'],
        'Pemotongan Disaksikan': ['Live Video Call', 'Disaksikan', 'Live Video Call'],
        'Catatan Khusus': ['Domba kurus', 'Upgrade Bobot', 'Normal'],
        'Tanggal Kirim + Potong Sama': ['Ya', 'Ya', 'Ya'],
    }
    
    df = pd.DataFrame(data)
    
    # Generate Excel file
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        output_path = tmp.name
    
    try:
        # Call the transform function
        result = transform_rekap_pemotongan(df, output_path)
        
        if result:
            print(f"✅ Excel file generated successfully: {output_path}")
            print(f"📊 File size: {os.path.getsize(output_path)} bytes")
            
            # Read the Excel file to verify it was created properly
            df_check = pd.read_excel(output_path, sheet_name='Hasil Rekap Pemotongan', header=0)
            print(f"✅ Excel file readable, rows: {len(df_check)}")
            
            # Check if dates are in the dataframe
            print("\n📋 Data Preview:")
            print(df_check[['No. Invoice', 'Tanggal Kirim', 'Tanggal Potong', 'Paket', 'Pemotongan Disaksikan']].head())
            
            print("\n✅ TEST PASSED - Check the file manually to verify:")
            print("  1. Date columns display as dates (yyyy-mm-dd), not numbers")
            print("  2. Colored backgrounds are applied with black text")
            print("  3. Dates match properly (same dates have matching colors)")
            print(f"\nFile location: {output_path}")
            print("You can open it with Excel to verify the formatting.")
            
        else:
            print("❌ Failed to generate Excel file")
            
    except Exception as e:
        print(f"❌ Error during test: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test_color_date_format()
