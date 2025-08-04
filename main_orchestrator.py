"""
Main orchestrator for annual credit review automation.
Coordinates data download, report generation, archiving, and emailing.
"""
import os
import time
import shutil
import glob
import subprocess
import sys
from datetime import datetime
import logging
import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from io import BytesIO

# Configure logging to file and console
LOG_FILE = 'annual_review.log'
logging.basicConfig(
    level=logging.INFO,
    format='[%(levelname)s] %(asctime)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler(LOG_FILE, mode='a', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

# Import email functionality
try:
    from email_config import EmailSender
    EMAIL_ENABLED = True
except ImportError:
    EMAIL_ENABLED = False
    logging.warning("Email functionality not available - email_config.py not found")

class ReportOrchestrator:
    """
    Orchestrates the annual credit review workflow: downloads data, generates reports, archives results, and sends emails.
    """
    def __init__(self):
        """
        Initialize directories and logging for the orchestrator.
        """
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.data_dir = os.path.abspath(os.path.join(self.base_dir, 'data'))
        self.data_bin_dir = os.path.abspath(os.path.join(self.base_dir, 'data_bin'))
        self.reports_dir = os.path.abspath(os.path.join(self.base_dir, 'reports'))
        self.download_attempts = []  # Store (attempt_number, timestamp) for logging
        
        # Create directories if they don't exist
        for directory in [self.data_dir, self.data_bin_dir, self.reports_dir]:
            if not os.path.exists(directory):
                os.makedirs(directory)
                logging.info(f"Created directory: {directory}")
    
    def run_selenium_download(self):
        """
        Run the selenium script to download data. Returns True if successful, False otherwise.
        """
        logging.info("Starting selenium data download process...")
        try:
            # Run the selenium script
            result = subprocess.run([sys.executable, 'selenium_erp_annual_review.py'], 
                                  capture_output=True, text=True, timeout=300)
            
            if result.returncode == 0:
                logging.info("Selenium download completed successfully")
                return True
            else:
                logging.error(f"Selenium download failed: {result.stderr}\nSTDOUT: {result.stdout}")
                return False
        except subprocess.TimeoutExpired:
            logging.error("Selenium download timed out")
            return False
        except Exception as e:
            logging.error(f"Error running selenium script: {e}")
            return False

    def download_with_retries(self, max_attempts=5, wait_after_fail=900):
        """Try to download the dataset up to max_attempts, then wait and try again, then fail with log."""
        self.download_attempts = []
        for attempt in range(1, max_attempts + 1):
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            logging.info(f"Download attempt {attempt} at {timestamp}")
            self.download_attempts.append((attempt, timestamp))
            success = self.run_selenium_download()
            if success:
                # Check if file exists
                excel_file = self.find_downloaded_file()
                if excel_file:
                    logging.info(f"Download succeeded on attempt {attempt} at {timestamp}")
                    return excel_file
                else:
                    logging.warning(f"No Excel file found after download attempt {attempt}")
            else:
                logging.warning(f"Download attempt {attempt} failed")
            # Wait a short time before next attempt
            time.sleep(5)
        # After max_attempts, wait 15 minutes and try again
        logging.warning(f"All {max_attempts} attempts failed. Waiting 15 minutes before trying again...")
        time.sleep(wait_after_fail)
        # One more attempt after waiting
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        logging.info(f"Download attempt {max_attempts+1} at {timestamp} (after 15 min wait)")
        self.download_attempts.append((max_attempts+1, timestamp))
        success = self.run_selenium_download()
        if success:
            excel_file = self.find_downloaded_file()
            if excel_file:
                logging.info(f"Download succeeded on attempt {max_attempts+1} at {timestamp}")
                return excel_file
            else:
                logging.warning(f"No Excel file found after download attempt {max_attempts+1}")
        else:
            logging.warning(f"Download attempt {max_attempts+1} failed")
        # If still not successful, log and print all attempts
        logging.error("Data downloading from ERP is failed. I have tried 5 attempts on the following timestamps:")
        for attempt, ts in self.download_attempts:
            logging.error(f"Attempt {attempt}: {ts}")
        print("\nData downloading from ERP is failed. I have tried 5 attempts on the following timestamps:")
        for attempt, ts in self.download_attempts:
            print(f"Attempt {attempt}: {ts}")
        return None

    def find_downloaded_file(self):
        """Find the most recently downloaded Excel file in the data directory."""
        logging.info("Searching for downloaded Excel file...")
        
        # Look for Excel files in the data directory
        excel_pattern = os.path.join(self.data_dir, "*.xlsx")
        excel_files = glob.glob(excel_pattern)
        
        if not excel_files:
            logging.error("No Excel files found in data directory")
            return None
        
        # Get the most recent file
        latest_file = max(excel_files, key=os.path.getctime)
        logging.info(f"Found downloaded file: {latest_file}")
        return latest_file
    
    def run_report_script(self, script_name, excel_file, output_pdf):
        """Run a report .py script with the given dataset and output PDF as arguments."""
        logging.info(f"Running script: {script_name} with dataset: {excel_file} and output: {output_pdf}")
        try:
            result = subprocess.run([
                sys.executable, script_name, excel_file, output_pdf
            ], capture_output=True, text=True, timeout=600)
            if result.returncode == 0:
                logging.info(f"Script {script_name} executed successfully. Output: {result.stdout}")
                return output_pdf
            else:
                logging.error(f"Script {script_name} failed. Stdout: {result.stdout}\nStderr: {result.stderr}")
                return None
        except Exception as e:
            logging.error(f"Failed to execute script {script_name}: {e}")
            return None

    def create_dated_folder_and_move_reports(self, auto_finance_pdf, three_wheeler_pdf):
        """
        Create a dated folder and move reports there. Avoids overwriting files.
        """
        try:
            current_date = datetime.now().strftime("%Y-%m-%d")
            dated_folder = os.path.join(self.reports_dir, f"{current_date}_Annual_review_reports")
            if not os.path.exists(dated_folder):
                os.makedirs(dated_folder)
                logging.info(f"Created dated folder: {dated_folder}")
            # Move reports to dated folder, avoid overwrite
            for pdf in [(auto_finance_pdf, 'Auto Finance'), (three_wheeler_pdf, 'Three Wheeler')]:
                src, label = pdf
                if src and os.path.exists(src):
                    dest = os.path.join(dated_folder, os.path.basename(src))
                    if os.path.exists(dest):
                        timestamp = datetime.now().strftime("%H%M%S")
                        dest = os.path.join(dated_folder, f"{os.path.splitext(os.path.basename(src))[0]}_{timestamp}.pdf")
                        logging.warning(f"{label} report already exists in destination. Renaming to avoid overwrite: {dest}")
                    try:
                        shutil.move(src, dest)
                        logging.info(f"Moved {label} report to: {dest}")
                    except PermissionError:
                        logging.error(f"Permission denied when moving {label} report. Is the file open?")
                        return None
            return dated_folder
        except Exception as e:
            logging.error(f"Error creating dated folder and moving reports: {e}")
            return None
    
    def move_dataset_to_bin(self, excel_file_path):
        """
        Move the dataset from data folder to data_bin folder. Avoids overwriting files.
        """
        try:
            if excel_file_path and os.path.exists(excel_file_path):
                filename = os.path.basename(excel_file_path)
                destination = os.path.join(self.data_bin_dir, filename)
                if os.path.exists(destination):
                    timestamp = datetime.now().strftime("%H%M%S")
                    destination = os.path.join(self.data_bin_dir, f"{os.path.splitext(filename)[0]}_{timestamp}.xlsx")
                    logging.warning(f"Dataset already exists in bin. Renaming to avoid overwrite: {destination}")
                try:
                    shutil.move(excel_file_path, destination)
                    logging.info(f"Moved dataset to data_bin: {destination}")
                    return True
                except PermissionError:
                    logging.error("Permission denied when moving dataset. Is the file open?")
                    return False
            else:
                logging.warning("No dataset file found to move")
                return False
        except Exception as e:
            logging.error(f"Error moving dataset to bin: {e}")
            return False
    
    def send_reports_via_email(self, auto_finance_pdf, three_wheeler_pdf, dated_folder):
        """Send reports via email to all configured groups."""
        if not EMAIL_ENABLED:
            logging.warning("Email functionality not enabled - skipping email sending")
            return True
        
        try:
            logging.info("📧 Sending reports via email...")
            email_sender = EmailSender()
            
            # Get the full paths of the reports in the dated folder
            auto_finance_path = os.path.join(dated_folder, os.path.basename(auto_finance_pdf))
            three_wheeler_path = os.path.join(dated_folder, os.path.basename(three_wheeler_pdf))
            
            # Check which reports exist
            auto_finance_exists = os.path.exists(auto_finance_path)
            three_wheeler_exists = os.path.exists(three_wheeler_path)
            
            if auto_finance_exists and three_wheeler_exists:
                # Send both reports
                success = email_sender.send_annual_review_reports(
                    auto_finance_path, 
                    three_wheeler_path, 
                    dated_folder
                )
            elif auto_finance_exists and not three_wheeler_exists:
                # Send only Auto Finance report
                logging.info("📧 Sending Auto Finance report only (Three Wheeler report not generated)")
                success = email_sender.send_partial_reports(
                    auto_finance_path, 
                    None, 
                    dated_folder,
                    "Auto Finance"
                )
            elif not auto_finance_exists and three_wheeler_exists:
                # Send only Three Wheeler report
                logging.info("📧 Sending Three Wheeler report only (Auto Finance report not generated)")
                success = email_sender.send_partial_reports(
                    None, 
                    three_wheeler_path, 
                    dated_folder,
                    "Three Wheeler"
                )
            else:
                # No reports generated
                logging.error("❌ No reports were generated")
                return False
            
            if success:
                logging.info("✅ Email sent successfully to all configured groups")
                return True
            else:
                logging.error("❌ Failed to send email")
                return False
                
        except Exception as e:
            logging.error(f"Error sending email: {e}")
            return False

    def run_complete_workflow(self):
        """Run the complete workflow with retry logic and logging, using .py scripts for analysis."""
        logging.info("🚀 Starting complete workflow...")
        try:
            # Step 1: Download with retries
            excel_file = self.download_with_retries()
            if not excel_file:
                logging.error("❌ Data download failed after all attempts. Stopping workflow.")
                return False
            
            # Step 2: Generate reports by running .py scripts
            auto_finance_pdf = 'Auto_finance_annual_review_report.pdf'
            three_wheeler_pdf = 'ThreeWheeler_annual_review_report.pdf'
            auto_finance_pdf_path = os.path.join(self.base_dir, auto_finance_pdf)
            three_wheeler_pdf_path = os.path.join(self.base_dir, three_wheeler_pdf)
            
            # Remove old PDFs if they exist
            for pdf in [auto_finance_pdf_path, three_wheeler_pdf_path]:
                if os.path.exists(pdf):
                    os.remove(pdf)
            
            # Run Auto Finance script
            auto_finance_result = self.run_report_script(
                os.path.join(self.base_dir, 'AutoFinance_report.py'),
                excel_file,
                auto_finance_pdf
            )
            
            # Run Three Wheeler script
            three_wheeler_result = self.run_report_script(
                os.path.join(self.base_dir, 'ThreeWheel_report.py'),
                excel_file,
                three_wheeler_pdf
            )
            
            # Check which reports were generated
            auto_finance_exists = auto_finance_result and os.path.exists(auto_finance_pdf_path)
            three_wheeler_exists = three_wheeler_result and os.path.exists(three_wheeler_pdf_path)
            
            # Handle case where both reports fail
            if not auto_finance_exists and not three_wheeler_exists:
                logging.warning("⚠️ Both reports failed to generate. Sending notification email...")
                
                try:
                    if EMAIL_ENABLED:
                        email_sender = EmailSender()
                        
                        # Define filter criteria for both reports
                        auto_finance_criteria = """- Blank REPORT_REVIEW_DATE
- Non-blank PRE_APPROVED_DATE
- Target products: VEHICLE LOAN-REGISTERED, TRACTOR LEASE, PLEDGE LOAN, OTHER LEASE, Murabaha, MINI TRUCK LEASE, IJARAH LEASE, HIRE PURCHASE-UN-REGISTERED, HIRE PURCHASE-REGISTERED, VEHICLE LOAN-UN-REGISTERED"""
                        
                        three_wheeler_criteria = """- Blank REPORT_REVIEW_DATE
- Non-blank PRE_APPROVED_DATE
- Target products: CASH IN HAND, Three Wheeler-Lease-Registered, Three Wheeler-Lease-Brand New, IJARAH SMALL LEASE"""
                        
                        # Send both reports failed notification
                        success = email_sender.send_both_reports_failed_notification(
                            auto_finance_criteria, 
                            three_wheeler_criteria
                        )
                        
                        if success:
                            logging.info("✅ Both reports failed notification sent successfully")
                        else:
                            logging.error("❌ Failed to send both reports failed notification")
                    else:
                        logging.warning("Email functionality not enabled - skipping notification")
                        
                except Exception as e:
                    logging.error(f"Error sending both reports failed notification: {e}")
                
                # Move dataset to bin even if no reports were generated
                self.move_dataset_to_bin(excel_file)
                
                logging.info("✅ Workflow completed with notification (no reports generated)")
                return True
            
            if not auto_finance_exists:
                logging.warning("⚠️ Auto Finance report generation failed")
            
            if not three_wheeler_exists:
                logging.warning("⚠️ Three Wheeler report generation failed (likely no data available)")
            
            # Step 3: Create dated folder and move available reports
            dated_folder = self.create_dated_folder_and_move_reports(auto_finance_pdf, three_wheeler_pdf)
            if not dated_folder:
                logging.error("❌ Failed to create dated folder and move reports.")
                return False
            
            # Step 4: Move dataset to bin
            self.move_dataset_to_bin(excel_file)
            
            # Step 5: Send reports via email
            email_success = self.send_reports_via_email(auto_finance_pdf, three_wheeler_pdf, dated_folder)
            
            # Log final status
            if auto_finance_exists and three_wheeler_exists:
                logging.info("✅ Complete workflow finished successfully!")
                logging.info("📊 Both reports have been generated and saved.")
            elif auto_finance_exists:
                logging.info("✅ Partial workflow completed successfully!")
                logging.info("📊 Auto Finance report generated and saved.")
                logging.info("📊 Three Wheeler report not generated (no data available).")
            elif three_wheeler_exists:
                logging.info("✅ Partial workflow completed successfully!")
                logging.info("📊 Three Wheeler report generated and saved.")
                logging.info("📊 Auto Finance report not generated.")
            
            logging.info(f"📁 Reports saved in: {dated_folder}")
            if email_success:
                logging.info("📧 Reports sent via email")
            else:
                logging.warning("⚠️ Email sending failed, but reports were generated successfully")
            return True
            
        except Exception as e:
            logging.error(f"❌ Workflow failed with error: {e}")
            return False

def main():
    orchestrator = ReportOrchestrator()
    success = orchestrator.run_complete_workflow()
    if success:
        print("\n🎉 Workflow completed successfully!")
        print("📊 Reports have been generated and saved.")
        print("📁 Dataset has been moved to data_bin folder.")
        print("📧 Reports have been sent via email.")
    else:
        print("\n❌ Workflow failed. Please check the logs above and in annual_review.log.")
        sys.exit(1)

if __name__ == "__main__":
    main() 