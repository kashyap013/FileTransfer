import datetime
import os
import re
import shutil
import sys
import time

# === LOG FILE SETUP ===
def setup_logging():
    """
    Sets up logging to a unique log file in the Logs directory.
    Returns the path to the created log file.
    """
    # Determine base directory (where the exe is located)
    if getattr(sys, 'frozen', False):
        # If running as executable
        base_path = os.path.dirname(sys.executable)
    else:
        # If running as script
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    # Create Logs directory at the base path
    logs_dir = os.path.join(base_path, "Logs")
    if not os.path.exists(logs_dir):
        try:
            os.makedirs(logs_dir)
        except Exception as e:
            print(f"Error creating Logs directory: {e}")
            # Fall back to base directory
            logs_dir = base_path
    
    # Create a unique log filename based on current timestamp
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f"transfer_log_{timestamp}.txt"
    log_path = os.path.join(logs_dir, log_filename)
    
    # Write header to the new log file
    with open(log_path, "w", encoding="utf-8") as log:
        log.write(f"=== File Transfer Log: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n\n")
    
    return log_path

# === GLOBALS ===
LOG_FILE = setup_logging()  # Initialize with a unique log file path
DESTINATION_ROOT = r"\\10.7.8.8\cor\PRD\__Product_Quality"  # Root path where files will be organized and moved

# === DESTINATION OPTIONS ===
# Mapping of user choices to specific subfolders in the destination
DESTINATION_OPTIONS = {
    "0": "0__QI",
    "1": "1__PCA Incoming",
    "2": "2__Pre-Conformal Coating",
    "3": "3__Post-Conformal Coating",
    "4": "4__Assembly",
    "5": "5__Final Outgoing",
    "6": "6__Misc",
    "7": "7__Pre-Vibration & Shock Testing",
    "8": "8__Post-Vibration & Shock Testing"
}

# === CREATE UNIQUE FILENAME WHERE SAME NAME FILE IS ALREADY PRESENT ===
def get_unique_filename(folder, filename):
    """
    Ensures the filename is unique within the folder.
    Adds a numeric suffix if needed to prevent overwriting.
    """
    base, ext = os.path.splitext(filename)
    counter = 1
    unique_path = os.path.join(folder, filename)
    while os.path.exists(unique_path):
        unique_path = os.path.join(folder, f"{base}_{counter}{ext}")
        counter += 1
    return unique_path

# === CHECK IF EXTRACTED SERIAL NUMBER IS 10-CHARACTER ALPHANUMERIC ===
def is_valid_serial(text):
    """
    Validates whether the provided string is a 10-character alphanumeric serial number.
    Returns cleaned version if valid, else None.
    """
    cleaned = text.strip().upper()
    return cleaned if len(cleaned) == 10 and cleaned.isalnum() else None

# === LOAD VALID PREFIXES FROM EXTERNAL FILE ===
def load_valid_prefixes(file_path="valid_prefixes.txt"):
    """
    Loads a list of valid 6-character serial number prefixes from a file.
    Raises errors if file is missing or empty.
    """
    if not os.path.exists(file_path):
        error_msg = f"Warning: Prefix file '{file_path}' not found. Please ensure it exists in the script directory.\n"
        log_message(error_msg, print_console=True)
        raise FileNotFoundError(error_msg)

    with open(file_path, "r", encoding="utf-8") as f:
        prefixes = [line.strip().upper() for line in f if line.strip()]

    if not prefixes:
        error_msg = f"Warning: Prefix file '{file_path}' cannot be empty. Please add valid prefixes list.\n"
        log_message(error_msg, print_console=True)
        raise ValueError(error_msg)

    return prefixes

# === LOGGING ===
def log_message(message, print_console=False):
    """
    Logs a message with a timestamp to a log file.
    Optionally prints to the console.
    """
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{timestamp}] {message}"
    with open(LOG_FILE, "a", encoding="utf-8") as log:
        log.write(entry + "\n")
    if print_console:
        print(message)

# === FILE OPERATIONS ===
def move_file(serial_number, file_path, source_folder, destination_subfolder):
    """
    Moves a file to the correct folder if the destination exists.
    Logs success or failure. Increments global counters accordingly.
    """
    global moved_count, skipped_count
    original_filename = os.path.basename(file_path)

    # Build full target path based on serial prefix and destination choice
    board_prefix = serial_number[:6]
    target_folder = os.path.join(DESTINATION_ROOT, board_prefix, serial_number, destination_subfolder)

    # Create destination folder if it doesn't exist
    if not os.path.exists(target_folder):
        try:
            os.makedirs(target_folder)
            log_message(f"Created destination folder: {target_folder}\n", print_console=True)
        except OSError as e:
            log_message(f"Warning: Error creating destination folder {target_folder}: {e}\n", print_console=True)
            skipped_count += 1
            return

    # Ensure filename doesn't overwrite existing one
    destination_path = get_unique_filename(target_folder, original_filename)

    try:
        shutil.move(file_path, destination_path)
        moved_count += 1
        log_message(f"Moved to: {destination_path}\n", print_console=True)
    except Exception as e:
        skipped_count += 1
        log_message(f"Warning: Error moving {original_filename} to {target_folder}: {e}\n", print_console=True)

def validate_serial_and_prefix(raw_serial, filename, valid_prefixes):
    """
    Validates the provided serial number and its prefix.

    Parameters:
    - raw_serial: The raw serial number string extracted from the filename.
    - filename: The filename being processed (used in log messages).
    - valid_prefixes: List of allowed 6-character prefixes.

    Returns:
    - Tuple (is_valid, message, cleaned_serial):
        - is_valid: True if serial number and prefix are valid, False otherwise.
        - message: Error message if validation fails, empty string if valid.
        - cleaned_serial: Cleaned and validated serial number if valid, None otherwise.
    """

    # Clean and validate serial number format (should be 10-character alphanumeric)
    cleaned_serial = is_valid_serial(raw_serial)
    if not cleaned_serial:
        # Serial number is missing, invalid length, or contains invalid characters
        return False, f"Warning: {filename}: Invalid or missing serial number. Skipping move.\n", None

    # Extract the prefix (first 6 characters) from the cleaned serial number
    prefix = cleaned_serial[:6]

    # Check if the prefix exists in the list of valid prefixes
    if prefix not in valid_prefixes:
        # Prefix is not allowed
        return False, f"Warning: {filename}: The detected prefix '{prefix}' is not in the valid prefixes list. Skipping move.\n", None

    # Serial number and prefix are valid
    return True, "", cleaned_serial

# === MAIN EXECUTION ===
def main():
    """
    Entry point of the script. Handles:
    - User prompt for destination folder
    - Prefix loading
    - Scans 'Files' folder
    - Validates filenames and moves files
    - Logging summary at the end including timing stats
    """
    global moved_count, skipped_count
    moved_count = 0
    skipped_count = 0

    # Load valid prefixes
    try:
        valid_prefixes = load_valid_prefixes()
    except (FileNotFoundError, ValueError):
        sys.exit(1)

    # Destination folder selection
    print("Please select the destination folder for Files:")
    for key, name in DESTINATION_OPTIONS.items():
        print(f"{key}. {name}")
    selected_key = None
    while selected_key not in DESTINATION_OPTIONS:
        user_input = input("Enter the number corresponding to your choice (0-8), or 'q' to quit: ").strip()
        if user_input.lower() == 'q':
            print("👋 Exiting script. Goodbye!")
            sys.exit(0)
        if user_input in DESTINATION_OPTIONS:
            selected_key = user_input
        else:
            print("❌ Invalid input.")

    destination_subfolder = DESTINATION_OPTIONS[selected_key]

    input(f"✅ You selected: {destination_subfolder}\nPlease copy all files into 'Files' folder and press Enter to begin...\n")

    # Determine script's working directory
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    source_folder = os.path.join(base_path, "Files")

    # Create 'Files' directory if it doesn't exist
    if not os.path.exists(source_folder):
        try:
            os.makedirs(source_folder)
            log_message(f"'Files' folder created at: {source_folder}\n", print_console=True)
        except OSError as e:
            log_message(f"Error creating 'Files' folder: {e}", print_console=True)
            sys.exit(1)
    elif not os.path.isdir(source_folder):
        log_message(f"Warning: '{source_folder}' exists but is not a directory!", print_console=True)
        sys.exit(1)

    # Validate that the source folder exists
    if not os.path.isdir(source_folder):
        log_message(f"Warning: 'Files' folder not found at expected path: {source_folder}\n", print_console=True)
        sys.exit(1)

    # Start log
    log_message("*" * 80, print_console=True)
    log_message("File Transfer Script execution started.\n", print_console=True)

    # Start timing
    start_time = time.time()
    total_files_processed = 0
    prefix_counts = {}       # Track files by prefix

    # === Process all files in the 'Files' folder ===
    for filename in os.listdir(source_folder):
        file_path = os.path.join(source_folder, filename)
        
        if os.path.isfile(file_path):
            total_files_processed += 1
            try:
                extracted_serial = re.split(r"[_\-]", filename)[0]
            except IndexError:
                extracted_serial = ""  # or some default value
                log_message(f"Warning: Could not extract serial number from {filename}. Invalid filename format.", print_console=True)
            
            # === Serial number validation and move file operation ===
            is_valid, message, serial = validate_serial_and_prefix(extracted_serial, filename, valid_prefixes)
            
            if not is_valid:
                skipped_count += 1
                log_message(message, print_console=True)
            else:
                # Extract and track prefix
                prefix = serial[:6]
                if prefix not in prefix_counts:
                    prefix_counts[prefix] = 0
                prefix_counts[prefix] += 1
                
                # Move the file
                move_file(serial, file_path, source_folder, destination_subfolder)
           
    # === TIME CALCULATION ===
    end_time = time.time()
    minutes, seconds = divmod(int(end_time - start_time), 60)
    average_time = (end_time - start_time) / total_files_processed if total_files_processed > 0 else 0

    # === FINAL SUMMARY ===
    log_message("=" * 40, print_console=True)
    log_message(f"Total files processed: {total_files_processed}", print_console=True)
    log_message(f"Total files moved: {moved_count}", print_console=True)
    log_message(f"Total files skipped: {skipped_count}", print_console=True)
    log_message(f"Total time taken: {minutes} minute(s) {seconds} second(s)", print_console=True)
    log_message(f"Average time per file: {average_time:.2f} second(s)", print_console=True)
    log_message("=" * 40, print_console=True)
    
    # Display destination summary
    log_message(f"All files were moved to: {destination_subfolder}", print_console=True)
    log_message("=" * 40, print_console=True)
    
    # Display prefix breakdown
    if prefix_counts:
        log_message("Files by prefix:", print_console=True)
        for prefix, count in prefix_counts.items():
            log_message(f"  {prefix}: {count} file(s)", print_console=True)
    log_message("=" * 40, print_console=True)
    
    log_message("File Transfer Script execution completed.", print_console=True)
    log_message("*" * 80, print_console=True)
    
    # Log the path to the log file itself
    print(f"\nLog file saved to: {LOG_FILE}")
    
    # Keep console window open when running as exe
    if getattr(sys, 'frozen', False):
        print("\nPress Enter to exit...")
        input()  # Wait for user to press Enter

# === SCRIPT ENTRY POINT ===
if __name__ == "__main__":
    main()
