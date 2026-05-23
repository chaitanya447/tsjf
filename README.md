#!/bin/bash

################################################################################
# Database to Excel Email Script
# Extracts data from SQL Server using freebcp, converts to Excel, and emails
################################################################################

set -euo pipefail

# Color codes for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

################################################################################
# USAGE FUNCTION
################################################################################
usage() {
 cat << EOF
Usage: $0 -s <sql_query> -d <database_server> -u <username> -p <password> \
 -w <work_directory> -f <freebcp_path> -e <recipient_email> [-c <cop_date>]

Required Arguments:
 -s SQL Query to execute (enclosed in quotes)
 -d Database server name (e.g., "SERVER_NAME")
 -u Database username
 -p Database password
 -w Working directory for output files (must exist)
 -f Path to freebcp executable (e.g., "/opt/sybutils/bin/freebcp")
 -e Recipient email DL (e.g., "krishna.k.kesavarapu@pwc.com")

Optional Arguments:
 -c COP_DATE for queries using \$COP_DATE (default: $(date +%Y-%m-%d))
 -I Path to interface file (freebcp -I parameter)
 -h Display this help message

Examples:
 # Basic usage with date parameter
 $0 -s "select * from employees where date = \$COP_DATE" \
 -d "MY_SERVER" \
 -u "myuser" \
 -p "mypass" \
 -w "/home/user/data" \
 -f "/opt/sybutils/bin/freebcp" \
 -e "team@pwc.com" \
 -c "2026-05-20"

 # Without date parameter
 $0 -s "select * from departments" \
 -d "MY_SERVER" \
 -u "myuser" \
 -p "mypass" \
 -w "/home/user/data" \
 -f "/opt/sybutils/bin/freebcp" \
 -e "individual@pwc.com"

EOF
 exit 1
}

################################################################################
# LOGGING FUNCTIONS
################################################################################
log_info() {
 echo -e "${GREEN}[INFO]${NC} $(date '+%Y-%m-%d %H:%M:%S') - $1"
}

log_warn() {
 echo -e "${YELLOW}[WARN]${NC} $(date '+%Y-%m-%d %H:%M:%S') - $1"
}

log_error() {
 echo -e "${RED}[ERROR]${NC} $(date '+%Y-%m-%d %H:%M:%S') - $1"
}

################################################################################
# VALIDATION FUNCTIONS
################################################################################
validate_freebcp() {
 local freebcp_path="$1"
 if [[ ! -x "$freebcp_path" ]]; then
 log_error "freebcp not found or not executable at: $freebcp_path"
 exit 1
 fi
 log_info "✓ freebcp found at: $freebcp_path"
}

validate_work_directory() {
 local work_dir="$1"
 if [[ ! -d "$work_dir" ]]; then
 log_error "Work directory does not exist: $work_dir"
 exit 1
 fi
 if [[ ! -w "$work_dir" ]]; then
 log_error "Work directory is not writable: $work_dir"
 exit 1
 fi
 log_info "✓ Work directory valid: $work_dir"
}

validate_email() {
 local email="$1"
 if [[ ! "$email" =~ ^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$ ]]; then
 log_error "Invalid email format: $email"
 exit 1
 fi
 log_info "✓ Email address valid: $email"
}

validate_command() {
 if ! command -v "$1" &> /dev/null; then
 log_error "Required command not found: $1"
 exit 1
 fi
 log_info "✓ Command available: $1"
}

################################################################################
# DATABASE EXTRACTION FUNCTION
################################################################################
extract_from_database() {
 local sql_query="$1"
 local database_server="$2"
 local database_username="$3"
 local database_password="$4"
 local work_dir="$5"
 local freebcp_path="$6"
 local interface_dir="$7"
 local output_csv="$8"

 log_info "Executing freebcp query..."

# Replace $COP_DATE if present (will be substituted by calling script)
 sql_query="${sql_query//\$COP_DATE/$COP_DATE}"

# Build freebcp command
 local freebcp_cmd="$freebcp_path \"$sql_query\" queryout \"$output_csv\" -c -t \",\""

# Add optional interface parameter
 if [[ -n "$interface_dir" && -f "$interface_dir" ]]; then
 freebcp_cmd="$freebcp_cmd -I \"$interface_dir\""
 fi

# Add authentication and server parameters
 freebcp_cmd="$freebcp_cmd -S \"$database_server\" -U \"$database_username\" -P \"$database_password\""

 log_info "Command: $freebcp_cmd"

# Execute the command
 if eval "$freebcp_cmd" > /dev/null 2>&1; then
 log_info "✓ Data extracted successfully to: $output_csv"

# Validate output file
 if [[ ! -f "$output_csv" ]]; then
 log_error "Output file was not created: $output_csv"
 exit 1
 fi

 local line_count=$(wc -l < "$output_csv")
 log_info "✓ CSV file contains $line_count lines"
 return 0
 else
 log_error "freebcp extraction failed"
 exit 1
 fi
}

################################################################################
# CSV TO EXCEL CONVERSION FUNCTION
################################################################################
convert_csv_to_excel() {
 local csv_file="$1"
 local excel_file="$2"
 local conversion_method=""

 log_info "Attempting to convert CSV to Excel format..."

# Try method 1: libreoffice (most reliable)
 if command -v libreoffice &> /dev/null; then
 log_info "Using LibreOffice for conversion..."
 conversion_method="libreoffice"

 if libreoffice --headless --convert-to xlsx \
 --outdir "$(dirname "$excel_file")" \
 "$csv_file" &> /dev/null; then

# LibreOffice creates file with .xlsx extension
 local temp_excel="${csv_file%.csv}.xlsx"
 if [[ -f "$temp_excel" ]]; then
 mv "$temp_excel" "$excel_file"
 log_info "✓ Converted using LibreOffice: $excel_file"
 return 0
 fi
 fi
 fi

# Try method 2: ssconvert from gnumeric
 if command -v ssconvert &> /dev/null; then
 log_info "Using ssconvert (Gnumeric) for conversion..."
 conversion_method="ssconvert"

 if ssconvert "$csv_file" "$excel_file" &> /dev/null; then
 log_info "✓ Converted using ssconvert: $excel_file"
 return 0
 fi
 fi

# Method 3: If no conversion tool available, use CSV with Excel-compatible encoding
 log_warn "No Excel conversion tool found (libreoffice or ssconvert)"
 log_warn "Proceeding with CSV file (Excel can open CSV files natively)"

# Copy CSV with UTF-8 BOM to make it Excel-friendly
 {
 printf '\xef\xbb\xbf' # UTF-8 BOM
 cat "$csv_file"
 } > "$excel_file"

 log_info "✓ Created Excel-compatible CSV with UTF-8 BOM: $excel_file"
 return 0
}

################################################################################
# EMAIL SENDING FUNCTION
################################################################################
send_email() {
 local recipient_email="$1"
 local attachment_file="$2"
 local subject="$3"
 local body_text="$4"

 log_info "Preparing email with attachment..."

# Validate mailx/mail command
 local mail_cmd=""
 if command -v mailx &> /dev/null; then
 mail_cmd="mailx"
 elif command -v mail &> /dev/null; then
 mail_cmd="mail"
 else
 log_error "Neither 'mailx' nor 'mail' command found"
 exit 1
 fi

 log_info "Using mail command: $mail_cmd"

# Prepare email body with uuencoded attachment
 local temp_email_body
 temp_email_body=$(mktemp)

 cat > "$temp_email_body" << EMAILEOF
$body_text

---
Attachment: $(basename "$attachment_file")
Generated: $(date '+%Y-%m-%d %H:%M:%S')
EMAILEOF

# Send email with uuencoded attachment
# Method 1: Try using uuencode (traditional approach)
 if command -v uuencode &> /dev/null; then
 log_info "Sending email with uuencoded attachment..."

 {
 cat "$temp_email_body"
 echo ""
 echo "---BEGIN ATTACHMENT---"
 uuencode "$attachment_file" "$(basename "$attachment_file")"
 } | $mail_cmd -s "$subject" "$recipient_email"

 if [[ $? -eq 0 ]]; then
 log_info "✓ Email sent successfully to: $recipient_email"
 rm -f "$temp_email_body"
 return 0
 fi
 fi

# Method 2: Modern mailx with -a flag
 if [[ "$mail_cmd" == "mailx" ]]; then
 log_info "Sending email with attachment using -a flag..."

 cat "$temp_email_body" | $mail_cmd -s "$subject" -a "$attachment_file" "$recipient_email"

 if [[ $? -eq 0 ]]; then
 log_info "✓ Email sent successfully to: $recipient_email"
 rm -f "$temp_email_body"
 return 0
 fi
 fi

 log_error "Failed to send email"
 rm -f "$temp_email_body"
 exit 1
}

################################################################################
# CLEANUP FUNCTION
################################################################################
cleanup() {
 local exit_code=$?

 if [[ $exit_code -eq 0 ]]; then
 log_info "Process completed successfully"
 else
 log_error "Process failed with exit code: $exit_code"
 fi

 exit $exit_code
}

trap cleanup EXIT

################################################################################
# MAIN EXECUTION
################################################################################
main() {
# Parse command line arguments
 local sql_query=""
 local database_server=""
 local database_username=""
 local database_password=""
 local work_dir=""
 local freebcp_path=""
 local recipient_email=""
 local cop_date=$(date +%Y-%m-%d)
 local interface_dir=""

 while getopts "s:d:u:p:w:f:e:c:I:h" opt; do
 case $opt in
s) sql_query="$OPTARG" ;;
d) database_server="$OPTARG" ;;
u) database_username="$OPTARG" ;;
p) database_password="$OPTARG" ;;
w) work_dir="$OPTARG" ;;
f) freebcp_path="$OPTARG" ;;
e) recipient_email="$OPTARG" ;;
c) cop_date="$OPTARG" ;;
I) interface_dir="$OPTARG" ;;
h) usage ;;
 *) usage ;;
 esac
 done

# Validate required arguments
 if [[ -z "$sql_query" || -z "$database_server" || -z "$database_username" || \
 -z "$database_password" || -z "$work_dir" || -z "$freebcp_path" || \
 -z "$recipient_email" ]]; then
 log_error "Missing required arguments"
 usage
 fi

# Export COP_DATE for use in freebcp
 export COP_DATE="$cop_date"

 log_info "=== Database to Excel Email Script Started ==="
 log_info "Database Server: $database_server"
 log_info "Work Directory: $work_dir"
 log_info "Recipient: $recipient_email"
 log_info "COP Date: $cop_date"

# Validations
 validate_freebcp "$freebcp_path"
 validate_work_directory "$work_dir"
 validate_email "$recipient_email"
 validate_command "mailx"

# Generate filenames with timestamp
 local timestamp=$(date +%Y%m%d_%H%M%S)
 local output_csv="$work_dir/data_extract_${timestamp}.csv"
 local output_excel="$work_dir/data_extract_${timestamp}.xlsx"

# Extract data
 log_info "--- Step 1: Extracting Data from Database ---"
 extract_from_database "$sql_query" "$database_server" "$database_username" \
 "$database_password" "$work_dir" "$freebcp_path" \
 "$interface_dir" "$output_csv"

# Convert to Excel
 log_info "--- Step 2: Converting CSV to Excel ---"
 convert_csv_to_excel "$output_csv" "$output_excel"

# Determine which file to attach
 local attachment_file="$output_excel"
 if [[ ! -f "$output_excel" ]]; then
 log_warn "Excel file not available, using CSV"
 attachment_file="$output_csv"
 fi

# Send email
 log_info "--- Step 3: Sending Email ---"
 local email_subject="Database Extract Report - $(date +%Y-%m-%d)"
 local email_body="Please find attached the requested database extract.

Extract Details:
- Generated: $(date '+%Y-%m-%d %H:%M:%S')
- COP Date: $cop_date
- File: $(basename "$attachment_file")
- Records: $(wc -l < "$output_csv")"

 send_email "$recipient_email" "$attachment_file" "$email_subject" "$email_body"

 log_info "=== All Steps Completed Successfully ==="
 log_info "Files generated:"
 log_info " CSV: $output_csv"
 log_info " Excel: $output_excel"
}

# Run main function
main "$@"
