import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import logging
from dotenv import load_dotenv

# Load credentials
load_dotenv("credentials.env")

class EmailSender:
    def __init__(self):
        # Email configuration from credentials.env
        self.smtp_server = os.environ.get('SMTP_SERVER', 'smtp.gmail.com')
        self.smtp_port = int(os.environ.get('SMTP_PORT', '587'))
        self.sender_email = os.environ.get('SENDER_EMAIL')
        self.sender_password = os.environ.get('SENDER_PASSWORD')
        
        # Email groups configuration
        self.email_groups = {
            'general': self._parse_email_list(os.environ.get('GENERAL_EMAILS', '')),
            'group1': self._parse_email_list(os.environ.get('GROUP1_EMAILS', '')),
            'group2': self._parse_email_list(os.environ.get('GROUP2_EMAILS', '')),
            'group3': self._parse_email_list(os.environ.get('GROUP3_EMAILS', '')),
            'group4': self._parse_email_list(os.environ.get('GROUP4_EMAILS', '')),
            'group5': self._parse_email_list(os.environ.get('GROUP5_EMAILS', ''))
        }
        
        # Validate configuration
        if not self.sender_email or not self.sender_password:
            logging.error("Email credentials not configured in credentials.env")
            raise ValueError("Email credentials not configured")
    
    def _parse_email_list(self, email_string):
        """Parse comma-separated email list"""
        if not email_string:
            return []
        return [email.strip() for email in email_string.split(',') if email.strip()]
    
    def send_annual_review_reports(self, auto_finance_pdf, three_wheeler_pdf, dated_folder):
        """Send annual review reports to all email groups"""
        try:
            # Email content
            subject = f"Annual Credit Review Reports - {datetime.now().strftime('%Y-%m-%d')}"
            message = """
Dear Team,

Please find the Annual Credit Review Reports for Auto Finance and 3W facilities.

Reports included:
1. Auto Finance Annual Review Report
2. Three Wheeler Annual Review Report

These reports contain comprehensive analysis of credit evaluations including:
- Deviation analysis
- Extreme value analysis
- PRE_APPROVED_USER summaries
- Distribution charts

Best regards,
Credit Evaluation Team
            """.strip()
            
            # Get all unique recipients
            all_recipients = set()
            for group_name, emails in self.email_groups.items():
                all_recipients.update(emails)
                logging.info(f"Email group '{group_name}': {len(emails)} recipients")
            
            if not all_recipients:
                logging.warning("No email recipients configured")
                return False
            
            # Send email
            success = self._send_email(
                recipients=list(all_recipients),
                subject=subject,
                message=message,
                attachments=[auto_finance_pdf, three_wheeler_pdf]
            )
            
            if success:
                logging.info(f"[SUCCESS] Email sent successfully to {len(all_recipients)} recipients")
                return True
            else:
                logging.error("[ERROR] Failed to send email")
                return False
                
        except Exception as e:
            logging.error(f"Error sending email: {e}")
            return False
    
    def send_partial_reports(self, auto_finance_pdf, three_wheeler_pdf, dated_folder, available_report):
        """Send partial reports when only one report is available"""
        try:
            # Email content
            subject = f"Annual Credit Review Report - {available_report} - {datetime.now().strftime('%Y-%m-%d')}"
            message = f"""
Dear Team,

Please find the Annual Credit Review Report for {available_report} facilities.

Report included:
1. {available_report} Annual Review Report

Note: The other report was not generated due to insufficient data availability.

This report contains comprehensive analysis of credit evaluations including:
- Deviation analysis
- Extreme value analysis
- PRE_APPROVED_USER summaries
- Distribution charts

Best regards,
Credit Evaluation Team
            """.strip()
            
            # Get all unique recipients
            all_recipients = set()
            for group_name, emails in self.email_groups.items():
                all_recipients.update(emails)
                logging.info(f"Email group '{group_name}': {len(emails)} recipients")
            
            if not all_recipients:
                logging.warning("No email recipients configured")
                return False
            
            # Prepare attachments list
            attachments = []
            if auto_finance_pdf and os.path.exists(auto_finance_pdf):
                attachments.append(auto_finance_pdf)
            if three_wheeler_pdf and os.path.exists(three_wheeler_pdf):
                attachments.append(three_wheeler_pdf)
            
            # Send email
            success = self._send_email(
                recipients=list(all_recipients),
                subject=subject,
                message=message,
                attachments=attachments
            )
            
            if success:
                logging.info(f"[SUCCESS] Partial report email sent successfully to {len(all_recipients)} recipients")
                return True
            else:
                logging.error("[ERROR] Failed to send partial report email")
                return False
                
        except Exception as e:
            logging.error(f"Error sending partial report email: {e}")
            return False
    
    def send_no_data_notification(self, report_type, filter_criteria):
        """Send notification when no data is available for report generation"""
        try:
            # Email content
            subject = f"{report_type} Annual Credit Review - No Data Available - {datetime.now().strftime('%Y-%m-%d')}"
            message = f"""
Dear Team,

This is to inform you that no records were found for the {report_type} Annual Credit Review report with the current filtered criteria.

Filter criteria applied:
{filter_criteria}

The {report_type.lower()}_annual_review_report.pdf file will not be generated due to insufficient data.

Please review the data source and filtering criteria if this is unexpected.

Best regards,
Credit Evaluation Team
            """.strip()
            
            # Get all unique recipients
            all_recipients = set()
            for group_name, emails in self.email_groups.items():
                all_recipients.update(emails)
                logging.info(f"Email group '{group_name}': {len(emails)} recipients")
            
            if not all_recipients:
                logging.warning("[WARNING] No email recipients configured")
                return False
            
            # Send email without attachments
            success = self._send_email(
                recipients=list(all_recipients),
                subject=subject,
                message=message,
                attachments=None
            )
            
            if success:
                logging.info(f"[SUCCESS] No-data notification sent successfully to {len(all_recipients)} recipients")
                return True
            else:
                logging.error("[ERROR] Failed to send no-data notification")
                return False
                
        except Exception as e:
            logging.error(f"Error sending no-data notification: {e}")
            return False
    
    def send_both_reports_failed_notification(self, auto_finance_criteria, three_wheeler_criteria):
        """Send notification when both reports fail due to insufficient data"""
        try:
            # Email content
            subject = f"Annual Credit Review - No Data Available for Both Reports - {datetime.now().strftime('%Y-%m-%d')}"
            message = f"""
Dear Team,

This is to inform you that no records were found for both Annual Credit Review reports with the current filtered criteria.

Auto Finance Report - Filter criteria applied:
{auto_finance_criteria}

Three Wheeler Report - Filter criteria applied:
{three_wheeler_criteria}

Neither Auto_finance_annual_review_report.pdf nor ThreeWheeler_annual_review_report.pdf files will be generated due to insufficient data.

Please review the data source and filtering criteria if this is unexpected.

Best regards,
Credit Evaluation Team
            """.strip()
            
            # Get all unique recipients
            all_recipients = set()
            for group_name, emails in self.email_groups.items():
                all_recipients.update(emails)
                logging.info(f"Email group '{group_name}': {len(emails)} recipients")
            
            if not all_recipients:
                logging.warning("[WARNING] No email recipients configured")
                return False
            
            # Send email without attachments
            success = self._send_email(
                recipients=list(all_recipients),
                subject=subject,
                message=message,
                attachments=None
            )
            
            if success:
                logging.info(f"[SUCCESS] Both reports failed notification sent successfully to {len(all_recipients)} recipients")
                return True
            else:
                logging.error("[ERROR] Failed to send both reports failed notification")
                return False
                
        except Exception as e:
            logging.error(f"Error sending both reports failed notification: {e}")
            return False
    
    def _send_email(self, recipients, subject, message, attachments=None):
        """Send email with attachments"""
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = ', '.join(recipients)
            msg['Subject'] = subject
            
            # Add message body
            msg.attach(MIMEText(message, 'plain'))
            
            # Add attachments
            if attachments:
                for file_path in attachments:
                    if os.path.exists(file_path):
                        with open(file_path, "rb") as attachment:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(attachment.read())
                        
                        encoders.encode_base64(part)
                        part.add_header(
                            'Content-Disposition',
                            f'attachment; filename= {os.path.basename(file_path)}'
                        )
                        msg.attach(part)
                        logging.info(f"Attached: {os.path.basename(file_path)}")
                    else:
                        logging.warning(f"Attachment not found: {file_path}")
            
            # Send email
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                text = msg.as_string()
                server.sendmail(self.sender_email, recipients, text)
            
            return True
            
        except Exception as e:
            logging.error(f"Error in _send_email: {e}")
            return False
    
    def test_email_configuration(self):
        """Test email configuration"""
        logging.info("Testing email configuration...")
        
        # Check credentials
        if not self.sender_email:
            logging.error("[ERROR] SENDER_EMAIL not configured")
            return False
        if not self.sender_password:
            logging.error("[ERROR] SENDER_PASSWORD not configured")
            return False
        
        # Check email groups
        total_recipients = 0
        for group_name, emails in self.email_groups.items():
            logging.info(f"Group '{group_name}': {len(emails)} emails")
            total_recipients += len(emails)
        
        if total_recipients == 0:
            logging.warning("[WARNING] No email recipients configured")
        else:
            logging.info(f"[SUCCESS] Total recipients: {total_recipients}")
        
        return True

def create_email_template():
    """Create email configuration template"""
    template = """
# Email Configuration Template
# Add these lines to your credentials.env file

# SMTP Configuration
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
SENDER_EMAIL=your_email@gmail.com
SENDER_PASSWORD=your_app_password

# Email Groups (comma-separated)
GENERAL_EMAILS=email1@company.com,email2@company.com,email3@company.com
GROUP1_EMAILS=manager1@company.com,supervisor1@company.com
GROUP2_EMAILS=analyst1@company.com,analyst2@company.com
GROUP3_EMAILS=finance1@company.com,finance2@company.com
GROUP4_EMAILS=risk1@company.com,risk2@company.com
GROUP5_EMAILS=admin1@company.com,admin2@company.com
"""
    
    with open('email_config_template.txt', 'w') as f:
        f.write(template)
    
    logging.info("Email configuration template created: email_config_template.txt")

if __name__ == "__main__":
    # Test email configuration
    try:
        sender = EmailSender()
        sender.test_email_configuration()
    except Exception as e:
        logging.error(f"Email configuration test failed: {e}")
        create_email_template() 