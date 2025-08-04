#!/usr/bin/env python3
"""
Test script for the ReportOrchestrator class
This script tests the orchestrator functionality without running selenium
"""

import os
import sys
import pandas as pd
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(levelname)s] %(asctime)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def create_test_excel_file():
    """Create a test Excel file with sample data"""
    logging.info("Creating test Excel file...")
    
    # Sample data structure based on the expected Excel format
    test_data = {
        'FACILITY_NUMBER': ['FAC001', 'FAC002', 'FAC003', 'FAC004', 'FAC005'],
        'PRODUCT': ['VEHICLE LOAN-REGISTERED', 'Three Wheeler-Lease-Registered', 'TRACTOR LEASE', 'CASH IN HAND', 'IJARAH LEASE'],
        'FACILITY_AMT': [50000, 25000, 75000, 10000, 60000],
        'ASSET_DESCRIPTION': ['Toyota Car', 'Three Wheeler', 'Tractor', 'Cash', 'Vehicle'],
        'MAKE_DESCRIPTION': ['Toyota', 'Bajaj', 'Mahindra', 'N/A', 'Honda'],
        'MODEL_DESCRIPTION': ['Corolla', 'Auto Rickshaw', 'Tractor 475', 'N/A', 'City'],
        'YOM': [2020, 2021, 2019, 2022, 2020],
        'ASSET_VALUATION': [48000, 24000, 72000, 9500, 58000],
        'VALUATION': [50000, 25000, 75000, 10000, 60000],
        'ROUNDED_VALUE': [50000, 25000, 75000, 10000, 60000],
        'PRE_APPROVED_USER': ['User1', 'User2', 'User1', 'User3', 'User2'],
        'PRE_APPROVED_DATE': ['2024-01-15', '2024-01-16', '2024-01-17', '2024-01-18', '2024-01-19'],
        'REPORT_REVIEW_DATE': ['', '', '', '', '']
    }
    
    df = pd.DataFrame(test_data)
    
    # Create data directory if it doesn't exist
    data_dir = 'data'
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)
    
    # Save test file
    test_file_path = os.path.join(data_dir, 'test_evaluation_report.xlsx')
    df.to_excel(test_file_path, index=False)
    logging.info(f"Test Excel file created: {test_file_path}")
    
    return test_file_path

def test_orchestrator_functions():
    """Test the orchestrator functions"""
    logging.info("Testing orchestrator functions...")
    
    # Import the orchestrator
    try:
        from main_orchestrator import ReportOrchestrator
        logging.info("✅ Successfully imported ReportOrchestrator")
    except ImportError as e:
        logging.error(f"❌ Failed to import ReportOrchestrator: {e}")
        return False
    
    # Create orchestrator instance
    try:
        orchestrator = ReportOrchestrator()
        logging.info("✅ Successfully created ReportOrchestrator instance")
    except Exception as e:
        logging.error(f"❌ Failed to create ReportOrchestrator instance: {e}")
        return False
    
    # Test directory creation
    for directory in [orchestrator.data_dir, orchestrator.data_bin_dir, orchestrator.reports_dir]:
        if os.path.exists(directory):
            logging.info(f"✅ Directory exists: {directory}")
        else:
            logging.error(f"❌ Directory missing: {directory}")
            return False
    
    # Create test Excel file
    test_file = create_test_excel_file()
    
    # Test file finding
    try:
        found_file = orchestrator.find_downloaded_file()
        if found_file:
            logging.info(f"✅ Found downloaded file: {found_file}")
        else:
            logging.error("❌ No downloaded file found")
            return False
    except Exception as e:
        logging.error(f"❌ Error finding downloaded file: {e}")
        return False
    
    # Test report generation (without selenium)
    try:
        auto_finance_pdf = orchestrator.generate_auto_finance_report(test_file)
        if auto_finance_pdf and os.path.exists(auto_finance_pdf):
            logging.info(f"✅ Auto Finance report generated: {auto_finance_pdf}")
        else:
            logging.error("❌ Auto Finance report generation failed")
            return False
    except Exception as e:
        logging.error(f"❌ Error generating Auto Finance report: {e}")
        return False
    
    try:
        three_wheeler_pdf = orchestrator.generate_three_wheeler_report(test_file)
        if three_wheeler_pdf and os.path.exists(three_wheeler_pdf):
            logging.info(f"✅ Three Wheeler report generated: {three_wheeler_pdf}")
        else:
            logging.error("❌ Three Wheeler report generation failed")
            return False
    except Exception as e:
        logging.error(f"❌ Error generating Three Wheeler report: {e}")
        return False
    
    # Test dated folder creation and file moving
    try:
        dated_folder = orchestrator.create_dated_folder_and_move_reports(auto_finance_pdf, three_wheeler_pdf)
        if dated_folder and os.path.exists(dated_folder):
            logging.info(f"✅ Dated folder created: {dated_folder}")
            
            # Check if reports are in the folder
            auto_finance_in_folder = os.path.join(dated_folder, os.path.basename(auto_finance_pdf))
            three_wheeler_in_folder = os.path.join(dated_folder, os.path.basename(three_wheeler_pdf))
            
            if os.path.exists(auto_finance_in_folder):
                logging.info(f"✅ Auto Finance report moved to: {auto_finance_in_folder}")
            else:
                logging.error(f"❌ Auto Finance report not found in dated folder")
                return False
                
            if os.path.exists(three_wheeler_in_folder):
                logging.info(f"✅ Three Wheeler report moved to: {three_wheeler_in_folder}")
            else:
                logging.error(f"❌ Three Wheeler report not found in dated folder")
                return False
        else:
            logging.error("❌ Failed to create dated folder")
            return False
    except Exception as e:
        logging.error(f"❌ Error creating dated folder: {e}")
        return False
    
    # Test dataset moving to bin
    try:
        success = orchestrator.move_dataset_to_bin(test_file)
        if success:
            logging.info("✅ Dataset moved to data_bin successfully")
        else:
            logging.error("❌ Failed to move dataset to data_bin")
            return False
    except Exception as e:
        logging.error(f"❌ Error moving dataset to bin: {e}")
        return False
    
    logging.info("🎉 All tests passed successfully!")
    return True

def cleanup_test_files():
    """Clean up test files"""
    logging.info("Cleaning up test files...")
    
    # Remove test Excel file from data_bin
    test_file_in_bin = os.path.join('data_bin', 'test_evaluation_report.xlsx')
    if os.path.exists(test_file_in_bin):
        os.remove(test_file_in_bin)
        logging.info("Removed test file from data_bin")
    
    # Remove test Excel file from data
    test_file_in_data = os.path.join('data', 'test_evaluation_report.xlsx')
    if os.path.exists(test_file_in_data):
        os.remove(test_file_in_data)
        logging.info("Removed test file from data")
    
    # Remove test PDF files from reports
    import glob
    for pdf_file in glob.glob('reports/*/test_*.pdf'):
        os.remove(pdf_file)
        logging.info(f"Removed test PDF: {pdf_file}")

def main():
    """Main test function"""
    print("🧪 Testing ReportOrchestrator functionality...")
    print("=" * 50)
    
    try:
        success = test_orchestrator_functions()
        
        if success:
            print("\n✅ All tests passed! The orchestrator is working correctly.")
            print("🚀 You can now run the complete workflow with:")
            print("   python main_orchestrator.py")
            print("   or double-click run_annual_review.bat")
        else:
            print("\n❌ Some tests failed. Please check the logs above.")
            sys.exit(1)
            
    except Exception as e:
        print(f"\n❌ Test failed with error: {e}")
        sys.exit(1)
    finally:
        # Clean up test files
        cleanup_test_files()

if __name__ == "__main__":
    main() 