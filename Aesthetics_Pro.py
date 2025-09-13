from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import os
import shutil
import pyautogui
import glob


# Configure PyAutoGUI
pyautogui.FAILSAFE = True  # Move mouse to top left to abort
pyautogui.PAUSE = 0.5

# Create main folder at the start
MAIN_FOLDER = "Aesthetics Pro"
if not os.path.exists(MAIN_FOLDER):
    os.makedirs(MAIN_FOLDER)
    print(f"✓ Created main folder: {MAIN_FOLDER}")

def wait_for_spinner_to_disappear(timeout=30):
    """Wait for any loading spinners to disappear"""
    try:
        wait = WebDriverWait(driver, timeout)
        wait.until(EC.invisibility_of_element_located((By.ID, "spinnerapplayout")))
        print("Spinner disappeared")
    except TimeoutException:
        print("Spinner timeout - continuing anyway")

def wait_for_loading_to_complete(timeout=30):
    """Wait for the 'Loading...' message to disappear and actual content to load"""
    try:
        wait = WebDriverWait(driver, timeout)
        
        # Wait for the loading message to disappear
        wait.until(EC.invisibility_of_element_located(
            (By.XPATH, "//td[contains(@class, 'dataTables_empty') and contains(text(), 'Loading...')]")
        ))
        print("Loading completed!")
        
        # Give a small buffer for any final rendering
        time.sleep(1)
        
    except TimeoutException:
        print("Loading timeout - continuing anyway")

def close_modal_if_present():
    """Close the modal popup if it appears"""
    try:
        print("    Waiting for modal to appear...")
        
        # Wait for the modal to appear after page load
        modal = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "clientnotepop"))
        )
        
        # Wait for modal to be visible
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "clientnotepop"))
        )
        
        print("     Modal appeared and is visible")
        
        # Find all close buttons and get the visible one
        close_buttons = driver.find_elements(By.ID, "cboxClose")
        visible_close_button = None
        
        for button in close_buttons:
            if button.is_displayed():
                visible_close_button = button
                break
        
        if visible_close_button:
            # Try to click the visible close button
            try:
                visible_close_button.click()
                print("     Modal closed with regular click")
            except:
                # Fallback to JavaScript click
                driver.execute_script("arguments[0].click();", visible_close_button)
                print("     Modal closed with JavaScript click")
        else:
            # Force close with JavaScript if no visible button found
            driver.execute_script("document.getElementById('clientnotepop').style.display = 'none';")
            print("     Modal force closed")
        
        # Wait for modal to disappear
        time.sleep(2)
        return True
            
    except TimeoutException:
        print("     Timeout waiting for modal")
        return False
    except Exception as e:
        print(f"     Error with modal: {e}")
        return False

def extract_client_info():
    """Extract client name and ID from the client profile section"""
    try:
        print("    Extracting client information...")
        
        # Find the client profile div
        client_profile_div = driver.find_element(By.ID, "clientProfileInfoDiv")
        
        # Extract client name from h5 tag
        client_name_element = client_profile_div.find_element(By.TAG_NAME, "h5")
        client_name = client_name_element.text.strip()
        
        # Extract client ID from the paragraph containing "ID:"
        id_paragraph = client_profile_div.find_element(By.XPATH, ".//p[contains(text(), 'ID:')]")
        client_id = id_paragraph.text.replace("ID:", "").strip()
        
        print(f"     Client Name: {client_name}")
        print(f"     Client ID: {client_id}")
        
        # Format folder name: Client Name - ID
        folder_name = f"{client_name} - {client_id}"
        print(f"     Folder Name: {folder_name}")
        
        return {
            'name': client_name,
            'id': client_id,
            'folder_name': folder_name
        }
        
    except Exception as e:
        print(f"     Error extracting client info: {e}")
        return {
            'name': 'Unknown',
            'id': 'Unknown',
            'folder_name': 'Unknown'
        }


def go_to_client_page(letter, page_number):
    """
    Click the letter and navigate to the correct pagination page.
    Page 1 = no action. Page 2+ = click page number or next button (fallback).
    """
    print(f"    Navigating to page {page_number} for letter {letter}...")

    if not click_letter(letter):
        print(f"     Failed to click letter {letter}")
        return False

    if page_number == 1:
        print(f"    ✓ Already on page 1")
        return True

    try:
        # First try to click the page number directly
        page_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//div[@id='clientlistTableBody_paginate']//a[text()='{page_number}']")
            )
        )
        page_button.click()
        print(f"     Clicked on page number {page_number}")
        wait_for_spinner_to_disappear()
        time.sleep(2)
        return True

    except Exception as e:
        print(f"     Page number {page_number} button not found, using 'Next' fallback...")

        # Fallback: click "Next" multiple times
        for i in range(1, page_number):
            try:
                next_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "clientlistTableBody_next"))
                )
                next_button.click()
                print(f"     Clicked 'Next' ({i}/{page_number - 1})")
                wait_for_spinner_to_disappear()
                time.sleep(2)
            except Exception as ex:
                print(f"     Failed to click 'Next' on step {i}: {ex}")
                return False

        return True


def click_electronic_records():
    """Click on Electronic Records tab"""
    try:
        print("    Clicking Electronic Records tab...")
        
        # Use direct JavaScript execution since it works reliably
        driver.execute_script("changeClientTab(2);")
        print("     Electronic Records tab clicked")
        
        # Wait for page to load
        time.sleep(3)
        return True
        
    except Exception as e:
        print(f"     Error clicking Electronic Records: {e}")
        return False

def create_client_folder(client_info):
    """Create client folder inside Scripting Data"""
    try:
        client_folder_path = os.path.join(MAIN_FOLDER, client_info['folder_name'])
        
        if not os.path.exists(client_folder_path):
            os.makedirs(client_folder_path)
            print(f"     Created client folder: {client_folder_path}")
        else:
            print(f"     Client folder already exists: {client_folder_path}")
        
        return client_folder_path
        
    except Exception as e:
        print(f"     Error creating client folder: {e}")
        return None


def handle_print_dialog():
    """Handle Chrome's print dialog using PyAutoGUI"""
    try:
        print("      Handling print dialog with PyAutoGUI...")
        
        # Wait for print dialog to appear
        time.sleep(2)
        
        # The print dialog should already have "Save as PDF" selected
        # Just press Enter to proceed to save dialog
        pyautogui.press('enter')
        print("       Pressed Enter on print dialog")
        
        # Wait for save dialog
        time.sleep(1)
        return True
        
    except Exception as e:
        print(f"       Error handling print dialog: {e}")
        return False

def handle_save_dialog(file_name, client_folder_path, record_title, client_name):
    """Handle the OS save dialog using PyAutoGUI with client subfolder structure"""
    try:
        print("      Handling save dialog with PyAutoGUI...")
        
        # Wait for save dialog to appear
        time.sleep(1)
        
        # Parse the date from record_title (e.g., "06/08/2023 HRT LABS" or "11/23/2022 HRT")
        try:
            # Split the record title to get date and rest
            parts = record_title.split(' ', 1)
            date_part = parts[0]  # "06/08/2023"
            record_name = parts[1] if len(parts) > 1 else ""  # "HRT LABS"
            
            # Convert date from MM/DD/YYYY to YYYY-MM-DD
            date_components = date_part.split('/')
            if len(date_components) == 3:
                month = date_components[0].zfill(2)
                day = date_components[1].zfill(2)
                year = date_components[2]
                formatted_date = f"{year}-{month}-{day}"
            else:
                # Fallback if date parsing fails
                formatted_date = date_part.replace('/', '-')
                
            print(f"      Parsed date: {date_part} -> {formatted_date}")
            print(f"      Record name: {record_name}")
            
        except Exception as e:
            print(f"      Error parsing date: {e}")
            formatted_date = "unknown-date"
            record_name = record_title
        
        # Clean the file name and other components
        safe_file_name = file_name.replace("/", "-").replace("\\", "-").replace(":", "").replace("*", "").replace("?", "").replace('"', "").replace("<", "").replace(">", "").replace("|", "")
        safe_record_name = record_name.replace("/", "-").replace("\\", "-").replace(":", "").replace("*", "").replace("?", "").replace('"', "").replace("<", "").replace(">", "").replace("|", "")
        safe_client_name = client_name.replace("/", "-").replace("\\", "-").replace(":", "").replace("*", "").replace("?", "").replace('"', "").replace("<", "").replace(">", "").replace("|", "")
        
        # Create the new filename format: YYYY-MM-DD_RecordName FileName_Treatment_Records_ClientName.pdf
        new_filename = f"{formatted_date}_{safe_record_name} {safe_file_name}_Treatment_Records_{safe_client_name}.pdf"
        
        # Full path includes client subfolder
        full_path = os.path.join(os.path.abspath(client_folder_path), new_filename)
        
        print(f"      New filename format: {new_filename}")
        print(f"      Typing full path: {full_path}")
        
        # Clear any existing filename and path
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.2)
        
        # Type the full path slowly to ensure it's entered correctly
        pyautogui.typewrite(full_path, interval=0.02)
        
        # Wait a moment before pressing enter
        time.sleep(0.5)
        
        # Press Enter to save
        pyautogui.press('enter')
        print("       Pressed Enter to save file")
        
        # Wait for file to be saved
        time.sleep(3)
        
        # Check if file was saved with the new name
        if os.path.exists(full_path):
            print(f"       File saved successfully: {new_filename}")
            return True
        else:
            # Sometimes the file might be saved with a different name or in wrong location
            try:
                # Check in main folder
                pdfs_main = glob.glob(os.path.join(os.path.abspath(MAIN_FOLDER), '*.pdf'))
                # Check in client folder
                pdfs_client = glob.glob(os.path.join(os.path.abspath(client_folder_path), '*.pdf'))
                
                all_pdfs = pdfs_main + pdfs_client
                
                if all_pdfs:
                    newest_pdf = max(all_pdfs, key=os.path.getmtime)
                    # If it's not our expected file, move/rename it
                    if newest_pdf != full_path:
                        print(f"      File saved as: {os.path.basename(newest_pdf)}")
                        print(f"      Moving/renaming to: {new_filename}")
                        # Ensure client folder exists
                        if not os.path.exists(client_folder_path):
                            os.makedirs(client_folder_path)
                        shutil.move(newest_pdf, full_path)
                        print(f"       File moved/renamed successfully")
                    return True
                else:
                    print(f"       No PDFs found after save!")
                    return False
            except Exception as e:
                print(f"       Error checking/moving file: {e}")
                return False
    
    except Exception as e:
        print(f"       Error handling save dialog: {e}")
        return False


def click_print_button():
    """Click the print button inside the modal/slide"""
    try:
        print("      Clicking print button in modal...")
        
        # Initial wait for modal to appear
        time.sleep(7)
        
        # Remove any toast notifications that might interfere
        try:
            toasts = driver.find_elements(By.CSS_SELECTOR, "div.toast-body, div.toast, div.d-flex:has(div.toast-body)")
            removed_count = 0
            for toast in toasts:
                if toast.is_displayed():
                    try:
                        # Try to remove it from DOM
                        driver.execute_script("arguments[0].remove();", toast)
                        removed_count += 1
                    except:
                        try:
                            # If removal fails, try to hide it
                            driver.execute_script("arguments[0].style.display = 'none';", toast)
                            removed_count += 1
                        except:
                            pass
            
            if removed_count > 0:
                print(f"       Removed {removed_count} toast notification(s)")
                time.sleep(1)
        except Exception as e:
            print(f"      Toast removal attempted")
        
        # Check if the form is fully loaded by looking for the action buttons
        print("      Checking if form is fully loaded...")
        form_loaded = False
        max_wait = 40  # Maximum 20 seconds to wait for form
        wait_time = 0
        
        while wait_time < max_wait and not form_loaded:
            try:
                # Look for the specific form action buttons that indicate the form is loaded
                er_icons_div = driver.find_elements(By.CSS_SELECTOR, "div.er_icons")
                email_download_buttons = driver.find_elements(By.XPATH, "//div[@class='er_icons']//div[contains(@onclick, 'emailDownloadERC')]")
                refresh_form_buttons = driver.find_elements(By.XPATH, "//div[@class='er_icons']//div[contains(@onclick, 'refreshFormERC')]")
                
                # Also check for the print button itself
                print_slide = driver.find_elements(By.ID, "printslide")
                
                if er_icons_div and email_download_buttons and refresh_form_buttons and print_slide:
                    form_loaded = True
                    print(f"       Form is fully loaded (Email/Download and Refresh Form buttons are present)")
                else:
                    components_found = []
                    if er_icons_div:
                        components_found.append("er_icons div")
                    if email_download_buttons:
                        components_found.append("Email/Download button")
                    if refresh_form_buttons:
                        components_found.append("Refresh Form button")
                    if print_slide:
                        components_found.append("print slide")
                    
                    print(f"      Waiting for form to load... ({wait_time}s) - Found: {', '.join(components_found) if components_found else 'none'}")
                    time.sleep(3)
                    wait_time += 2
                    
            except Exception as e:
                print(f"      Error checking form load status: {e}")
                time.sleep(2)
                wait_time += 2
        
        if not form_loaded:
            print("        Form action buttons not found, but proceeding anyway...")
        
        # The print button is inside the printslide div
        try:
            # First try to find the printslide div
            print_slide = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "printslide"))
            )
            
            # Find the print button within the printslide
            print_button = print_slide.find_element(By.ID, "printKey")
            
            # Ensure print button is visible and enabled
            if not print_button.is_displayed():
                driver.execute_script("arguments[0].scrollIntoView(true);", print_button)
                time.sleep(1)
            
            # Click the button
            try:
                print_button.click()
                print("       Print button clicked")
            except:
                driver.execute_script("arguments[0].click();", print_button)
                print("       Print button clicked (JavaScript)")
                
        except Exception as e:
            print(f"      DEBUG: Error with standard approach: {e}")
            # Fallback: Try to execute the onclick function directly
            try:
                driver.execute_script("runPrint();")
                print("       Print function called")
            except Exception as e2:
                print(f"      DEBUG: runPrint() failed: {e2}")
                return False
        
        return True
        
    except Exception as e:
        print(f"       Error clicking print button: {e}")
        return False

def close_pdf_dialog():
    """Close the PDF dialog/modal after saving"""
    try:
        print("      Closing modal after save...")
        
        # Look for the specific close button div
        close_button = None
        
        # Method 1: Look for the closeCustomSlide div
        try:
            close_button = driver.find_element(By.CSS_SELECTOR, "div.closeCustomSlide.nav-customclose")
        except:
            pass
        
        # Method 2: Try clicking the image inside the close div
        if not close_button:
            try:
                close_button = driver.find_element(By.CSS_SELECTOR, "div.closeCustomSlide img[src*='menu-close']")
            except:
                pass
        
        # Method 3: Look for any element with closeCustomSlide class
        if not close_button:
            try:
                close_button = driver.find_element(By.CLASS_NAME, "closeCustomSlide")
            except:
                pass
        
        if close_button and close_button.is_displayed():
            try:
                # Scroll into view first
                driver.execute_script("arguments[0].scrollIntoView(true);", close_button)
                time.sleep(0.5)
                close_button.click()
                print("       Modal closed")
            except:
                driver.execute_script("arguments[0].click();", close_button)
                print("       Modal closed (JavaScript)")
        else:
            # If no close button found, try pressing ESC
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                print("       Modal closed with ESC key")
            except:
                print("       Failed to close modal")
                return False
        
        # Wait for modal to close
        time.sleep(2)
        return True
        
    except Exception as e:
        print(f"       Error closing modal: {e}")
        return False

def process_treatment_record_files(treatment_records_li, record_title, client_folder_path, client_name):
    """Process all files in a Treatment Record folder"""
    try:
        print(f"      Processing files in: {record_title}")
        
        # Wait a moment for the Treatment Records to fully expand
        time.sleep(2)
        
        # Find only the files within the Treatment Records li element
        treatment_files = treatment_records_li.find_elements(By.CSS_SELECTOR, "ul.nested.sub-nested > li[id*='_doc'] > div.slide > a[onclick*='launchERForm']")
        
        total_files = len(treatment_files)
        print(f"      Found {total_files} files in Treatment Records")
        
        if total_files == 0:
            print(f"      No files found in Treatment Records for {record_title}")
            return True
        
        processed_files = []
        processed_count = 0
        
        # Process each file - but keep track to avoid double-processing
        for file_index in range(total_files):
            try:
                # Re-find the files to avoid stale elements
                treatment_files = treatment_records_li.find_elements(By.CSS_SELECTOR, "ul.nested.sub-nested > li[id*='_doc'] > div.slide > a[onclick*='launchERForm']")
                
                # Check if we've already processed all files we intended to
                if processed_count >= total_files:
                    print(f"        All {total_files} files have been processed")
                    break
                
                if file_index >= len(treatment_files):
                    print(f"        File index {file_index} out of range")
                    break
                    
                file_element = treatment_files[file_index]
                
                # Get file name from the text content
                file_name = None
                
                # First try to get text from .treetextitem element
                try:
                    file_name_element = file_element.find_element(By.CSS_SELECTOR, ".treetextitem")
                    file_name = file_name_element.text.strip()
                except:
                    # If that fails, try to get text directly from the anchor
                    try:
                        file_name = file_element.text.strip()
                    except:
                        file_name = f"File_{file_index + 1}"
                
                if not file_name or file_name == "":
                    file_name = f"File_{file_index + 1}"
                
                print(f"\n        ========== FILE {file_index + 1}/{total_files} ==========")
                print(f"        File name: {file_name}")
                print(f"        Starting to process this file...")
                
                # Scroll the element into view if needed
                if not file_element.is_displayed():
                    driver.execute_script("arguments[0].scrollIntoView(true);", file_element)
                    time.sleep(1)
                    print(f"        Scrolled file into view")
                
                # Click the file to open it
                print(f"        Attempting to click on file...")
                try:
                    file_element.click()
                    print(f"        ✓ Successfully clicked on file: {file_name}")
                except Exception as e:
                    print(f"        Regular click failed: {e}")
                    # If regular click fails, try JavaScript click
                    try:
                        driver.execute_script("arguments[0].click();", file_element)
                        print(f"        ✓ Successfully clicked on file using JavaScript: {file_name}")
                    except Exception as e2:
                        print(f"         JavaScript click also failed: {e2}")
                        continue
                
                # Wait for the file modal/form to load
                print(f"        Waiting for file to load...")
                time.sleep(5)
                
                # Click print button
                print(f"        Looking for print button...")
                if click_print_button():
                    print(f"         Print button clicked successfully")
                    
                    # Handle print dialog with PyAutoGUI
                    if handle_print_dialog():
                        print(f"         Print dialog handled successfully")
                        
                        # Handle save dialog with PyAutoGUI - now passing client_name
                        if handle_save_dialog(file_name, client_folder_path, record_title, client_name):
                            print(f"         File saved successfully")
                            processed_files.append(file_name)
                            processed_count += 1
                            
                            # Wait for save to complete
                            time.sleep(3)
                            
                            # Close the modal after saving
                            print(f"        Attempting to close the file modal...")
                            if close_pdf_dialog():
                                print(f"         Modal closed successfully")
                            else:
                                print(f"          Could not close modal, but file was saved")
                            
                            # IMPORTANT: Wait longer for DOM to stabilize after closing modal
                            print(f"        Waiting for page to stabilize...")
                            time.sleep(4)
                            
                        else:
                            print(f"         Failed to save file {file_name}")
                    else:
                        print(f"         Failed to handle print dialog for {file_name}")
                else:
                    print(f"         Failed to click print button for {file_name}")
                
                print(f"        ========== END OF FILE {file_index + 1}/{total_files} ==========\n")
                
                # Check if we've processed all files
                if processed_count >= total_files:
                    print(f"         All files processed, exiting loop")
                    break
                
                # Small delay before processing next file
                time.sleep(2)
                
            except Exception as e:
                print(f"         Error processing file at index {file_index}: {e}")
                import traceback
                traceback.print_exc()
                continue
        
        print(f"       Completed processing. Successfully processed {len(processed_files)} out of {total_files} files in {record_title}")
        return True
        
    except Exception as e:
        print(f"       Error processing treatment record files: {e}")
        import traceback
        traceback.print_exc()
        return False

def process_electronic_records(client_info):
    """Process electronic records - click on each record and look for Treatment Records"""
    try:
        print("    Processing Electronic Records...")
        
        # Create client folder
        client_folder_path = create_client_folder(client_info)
        if not client_folder_path:
            print("     Failed to create client folder")
            return False, []
        
        treatment_record_folders = []  # List to store folders with Treatment Records
        
        # Find all parent records
        parent_records = driver.find_elements(By.CSS_SELECTOR, "li.parentrec")
        total_records = len(parent_records)
        print(f"    Found {total_records} electronic records")
        
        if total_records == 0:
            print("    No electronic records found")
            return True, treatment_record_folders
        
        # Process each record
        for record_index in range(total_records):
            try:
                # Re-find parent records to avoid stale elements
                parent_records = driver.find_elements(By.CSS_SELECTOR, "li.parentrec")
                if record_index >= len(parent_records):
                    print(f"    Record index {record_index} out of range")
                    break
                
                record = parent_records[record_index]
                
                # Get the record title
                record_title_element = record.find_element(By.CSS_SELECTOR, "span.caret.parentcaret a")
                record_title = record_title_element.text.strip()
                print(f"\n    [{record_index + 1}/{total_records}] Processing record: {record_title}")
                
                # Check if parent record needs to be expanded
                parent_caret_span = record.find_element(By.CSS_SELECTOR, "span.caret.parentcaret")
                parent_classes = parent_caret_span.get_attribute('class')
                
                if 'caret-down' not in parent_classes:
                    print(f"      Expanding parent record...")
                    record_title_element.click()
                    time.sleep(2)  # Wait for expansion
                    print(f"       Parent record expanded")
                else:
                    print(f"      Parent record already expanded")
                
                # Look for "Treatment Records" folder within this record
                try:
                    # Re-find the record after expansion
                    parent_records = driver.find_elements(By.CSS_SELECTOR, "li.parentrec")
                    record = parent_records[record_index]
                    
                    # Find the Treatment Records folder
                    treatment_records_li = record.find_element(By.XPATH, ".//li[span[contains(@class, 'caret')]//a[contains(text(), 'Treatment Records')]]")
                    treatment_records_link = treatment_records_li.find_element(By.XPATH, ".//span[contains(@class, 'caret')]//a[contains(text(), 'Treatment Records')]")
                    
                    print(f"       Found 'Treatment Records' folder")
                    treatment_record_folders.append(record_title)
                    
                    # Check if Treatment Records folder needs to be expanded
                    tr_parent_span = treatment_records_link.find_element(By.XPATH, "./parent::span")
                    tr_span_classes = tr_parent_span.get_attribute('class')
                    
                    if 'caret-down' not in tr_span_classes:
                        print(f"      Expanding 'Treatment Records' folder...")
                        treatment_records_link.click()
                        time.sleep(2)  # Wait for expansion
                        print(f"       'Treatment Records' folder expanded")
                        
                        # Verify it's actually expanded
                        tr_parent_span = treatment_records_link.find_element(By.XPATH, "./parent::span")
                        tr_span_classes = tr_parent_span.get_attribute('class')
                        if 'caret-down' in tr_span_classes:
                            print(f"       Confirmed: Treatment Records is now expanded")
                        else:
                            print(f"        Warning: Treatment Records may not have expanded properly")
                    else:
                        print(f"      'Treatment Records' folder already expanded")
                    
                    # Check if the nested ul has the 'active' class
                    nested_ul = treatment_records_li.find_element(By.CSS_SELECTOR, "ul.nested.sub-nested")
                    ul_classes = nested_ul.get_attribute('class')
                    print(f"      Treatment Records UL classes: {ul_classes}")
                    
                    if 'active' not in ul_classes:
                        print(f"        Warning: Nested UL doesn't have 'active' class, files might not be visible")
                    
                    # Process files in the Treatment Records folder
                    process_treatment_record_files(treatment_records_li, record_title, client_folder_path, client_info['name'])
                    
                except Exception as e:
                    print(f"       No 'Treatment Records' folder found in this record: {e}")
                
            except Exception as e:
                print(f"       Error processing record {record_index + 1}: {e}")
                import traceback
                traceback.print_exc()
                continue
        
        print(f"\n     Finished processing all electronic records")
        print(f"     Found {len(treatment_record_folders)} folders with Treatment Records:")
        for folder in treatment_record_folders:
            print(f"       {folder}")
        
        return True, treatment_record_folders
        
    except Exception as e:
        print(f"     Error processing electronic records: {e}")
        import traceback
        traceback.print_exc()
        return False, []


def cleanup_open_modals():
    """Close any open modals/slides before proceeding"""
    try:
        print("    Running modal cleanup...")
        modals_closed = 0
        
        # Look for closeCustomSlide buttons
        try:
            close_buttons = driver.find_elements(By.CSS_SELECTOR, "div.closeCustomSlide.nav-customclose")
            for close_btn in close_buttons:
                if close_btn.is_displayed():
                    try:
                        close_btn.click()
                        modals_closed += 1
                        time.sleep(0.5)
                    except:
                        try:
                            driver.execute_script("arguments[0].click();", close_btn)
                            modals_closed += 1
                            time.sleep(0.5)
                        except:
                            pass
        except:
            pass
        
        # Look for any close button with X if above fails
        try:
            x_buttons = driver.find_elements(By.XPATH, "//a[@class='close' and text()='×']")
            for x_btn in x_buttons:
                if x_btn.is_displayed():
                    try:
                        x_btn.click()
                        modals_closed += 1
                        time.sleep(0.5)
                    except:
                        try:
                            driver.execute_script("arguments[0].click();", x_btn)
                            modals_closed += 1
                            time.sleep(0.5)
                        except:
                            pass
        except:
            pass
        
        # Press ESC key if all above fails
        try:
            driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            time.sleep(0.5)
        except:
            pass
        
        if modals_closed > 0:
            print(f"     Closed {modals_closed} modal(s)")
            time.sleep(1)  # Give time for animations to complete
        else:
            print("     No modals to close")
            
        return True
        
    except Exception as e:
        print(f"     Error during modal cleanup: {e}")
        return False


def navigate_to_client_list():
    """Navigate to client list through the menu: Clients > Client List"""
    try:
        print("    Navigating back to Client List through menu...")
        
        # CLEANUP Try to close any open modals before navigation
        print("    Checking for any open modals to close...")
        try:
            # Look for the closeCustomSlide element
            close_buttons = driver.find_elements(By.CSS_SELECTOR, "div.closeCustomSlide.nav-customclose")
            closed_count = 0
            
            for close_btn in close_buttons:
                if close_btn.is_displayed():
                    try:
                        # Try regular click
                        close_btn.click()
                        closed_count += 1
                        print(f"     Closed modal {closed_count} with regular click")
                        time.sleep(1)
                    except:
                        try:
                            # Try JavaScript click
                            driver.execute_script("arguments[0].click();", close_btn)
                            closed_count += 1
                            print(f"     Closed modal {closed_count} with JavaScript click")
                            time.sleep(1)
                        except:
                            pass
            
            if closed_count > 0:
                print(f"    ✓ Closed {closed_count} modal(s) during cleanup")
                time.sleep(2)  # Wait for modals to fully close
            else:
                print("    No visible modals found during cleanup")
                
        except Exception as e:
            print(f"    Cleanup check completed (no modals or error: {e})")
        
        # Also try pressing ESC key as additional cleanup
        try:
            driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            time.sleep(0.5)
            print("     Pressed ESC key for additional cleanup")
        except:
            pass
        
        # Now check for blocking modals (like clientnotepop)
        try:
            # Look for any visible modal
            modals = driver.find_elements(By.CSS_SELECTOR, ".modal.show, .modal.fade.show, #clientnotepop")
            for modal in modals:
                if modal.is_displayed():
                    print("    DEBUG Found blocking modal, attempting to close...")
                    # Try to find and click close button
                    try:
                        close_btn = modal.find_element(By.CSS_SELECTOR, ".close, .btn-close, [data-dismiss='modal'], [data-bs-dismiss='modal']")
                        close_btn.click()
                        print("     Closed blocking modal")
                        time.sleep(1)
                    except:
                        # Try ESC key
                        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                        print("     Closed modal with ESC")
                        time.sleep(1)
        except:
            pass
        
        # Click "Clients" in the main menu
        clients_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Clients') and contains(@class, 'dc-mega')]"))
        )
        clients_button.click()
        print("     Clicked 'Clients' menu")
        time.sleep(1)
        
        # Click "Client List" submenu
        client_list_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Client List') and contains(@class, 'menulink')]"))
        )
        client_list_link.click()
        print("     Clicked 'Client List' submenu")
        
        # Wait for client list page to load
        wait_for_spinner_to_disappear()
        wait_for_loading_to_complete()
        
        print("     Successfully navigated to Client List")
        return True
        
    except Exception as e:
        print(f"     Error navigating to client list: {e}")
        
        # Try recovery: navigate directly to client list URL
        try:
            print("    Attempting direct navigation to client list...")
            driver.get("https://www.myaestheticspro.com/6722np22veg5/clients/index.cfm")
            time.sleep(3)
            wait_for_spinner_to_disappear()
            wait_for_loading_to_complete()
            print("     Direct navigation successful")
            return True
        except:
            pass
            
        return False

def get_client_links_for_letter():
    """Get all client links for the current letter"""
    client_links = []
    try:
        # Find all client name links with onclick="clientdetails"
        client_elements = driver.find_elements(By.CSS_SELECTOR, "#tblClientLeadListBody tr:not(.dataTables_empty) td.clientName a[onclick*='clientdetails']")
        
        for element in client_elements:
            try:
                client_name = element.text.strip()
                onclick_attr = element.get_attribute('onclick')
                client_links.append({
                    'name': client_name,
                    'onclick': onclick_attr
                })
            except Exception as e:
                print(f"Error getting client link: {e}")
                continue
                
    except Exception as e:
        print(f"Error getting client links: {e}")
    
    return client_links

def click_letter(letter):
    """Click on a specific letter to filter clients"""
    try:
        letter_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//div[contains(@class,'sortingRow')]//a[text()='{letter}']"))
        )
        letter_link.click()
        print(f"     Clicked letter: {letter}")
        
        # Wait for the list to load
        wait_for_spinner_to_disappear()
        
        # Add extra wait time to ensure the letter filter is applied
        time.sleep(3)
        
        wait_for_loading_to_complete()
        return True
        
    except Exception as e:
        print(f"     Error clicking letter {letter}: {e}")
        return False

def robust_navigate_to_client_list():
    """Robustly navigates to the Client > Client List page, handling UI glitches and scroll bugs."""
    print("\n Navigating to Client List (robust fallback enabled)...")

    # Try regular menu click
    try:
        client_menu = driver.find_element(By.LINK_TEXT, "Client")
        client_menu.click()
        time.sleep(0.5)
        client_list = driver.find_element(By.LINK_TEXT, "Client List")
        client_list.click()
        wait_for_spinner_to_disappear()
        print("Reached Client List via normal sidebar.")
        return True
    except Exception as e:
        print(f"Sidebar click failed: {e}")

    # Try re-expanding collapsed sidebar
    try:
        toggle = driver.find_element(By.CSS_SELECTOR, ".sidebar-collapse-toggle")
        if toggle.is_displayed():
            toggle.click()
            print("Sidebar re-expanded.")
            time.sleep(1)
    except:
        pass  # Not all views have a collapsible sidebar

    # Scroll to top and try JS click on Client List menu
    try:
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)
        client_list_js = driver.find_element(By.CSS_SELECTOR, "a.menulink[onclick*='client_list']")
        driver.execute_script("arguments[0].click();", client_list_js)
        wait_for_spinner_to_disappear()
        print("Client List reached via scroll + JS click.")
        return True
    except Exception as e:
        print(f"JS click after scroll failed: {e}")

    # Try directly triggering the JS page loader
    try:
        driver.execute_script("pagenav('clients/client_list/index.cfm', 0);")
        wait_for_spinner_to_disappear()
        print("Triggered Client List via JavaScript pagenav().")
        return True
    except Exception as e:
        print(f"JS function pagenav() failed: {e}")

    # FINAL fallback load the Client List page directly
    try:
        print("Loading Client List page directly...")
        driver.get("https://secure.aestheticspro.com/clients/client_list/index.cfm")
        wait_for_spinner_to_disappear()
        print("Loaded Client List via direct URL.")
        return True
    except Exception as e:
        print(f"All Client List navigation attempts failed: {e}")
        return False


def process_clients_for_letter(letter, start_client_number=1):
    """Process all clients for a specific letter starting from a given client number"""
    print(f"\n{'='*60}")
    print(f"Processing clients for letter: {letter}")
    print(f"Starting from client number: {start_client_number}")
    print(f"{'='*60}")
    
    processed_clients = []
    client_index = start_client_number - 1  # Convert to 0-based index
    
    while True:
        try:
            # Calculate which page this client should be on (100 clients per page)
            page_number = (client_index // 100) + 1
            page_local_index = client_index % 100
            
            print(f"\n  Client #{client_index + 1} → Page {page_number}, Position {page_local_index + 1}")
            
            # Navigate to the correct page
            if not go_to_client_page(letter, page_number):
                print(f"     Failed to navigate to page {page_number}")
                break
            
            # Get all client links on the current page
            client_links_current = driver.find_elements(By.CSS_SELECTOR, 
                "#tblClientLeadListBody tr:not(.dataTables_empty) td.clientName a[onclick*='clientdetails']")
            
            total_clients_on_page = len(client_links_current)
            print(f"    Found {total_clients_on_page} clients on page {page_number}")
            
            # Check if we've processed all clients
            if total_clients_on_page == 0:
                print(f"    No more clients found on page {page_number}")
                break
            
            # Check if our target client exists on this page
            if page_local_index >= total_clients_on_page:
                print(f"    No client at position {page_local_index + 1} on page {page_number}")
                # Check if there's a next page
                try:
                    next_button = driver.find_element(By.ID, "clientlistTableBody_next")
                    if "disabled" in next_button.get_attribute("class"):
                        print("    No more pages available")
                        break
                    else:
                        print("    Moving to next page...")
                        client_index = page_number * 100  # Jump to first client of next page
                        continue
                except:
                    print("    No pagination found")
                    break
            
            # Get the specific client to process
            client_link = client_links_current[page_local_index]
            actual_client_name = client_link.text.strip()
            
            print(f"\n  [Client #{client_index + 1}] Processing: {actual_client_name}")
            print(f"    (Page {page_number}, Position {page_local_index + 1} on page)")
            
            # Check if we already processed this client (avoid duplicates)
            already_processed = any(p['name'] == actual_client_name for p in processed_clients)
            if already_processed:
                print(f"     Client {actual_client_name} already processed, moving to next")
                client_index += 1
                continue
            
            # Scroll to and click the client
            driver.execute_script("arguments[0].scrollIntoView(true);", client_link)
            time.sleep(1)
            
            # Try clicking with JavaScript if regular click fails
            try:
                client_link.click()
                print(f"     Clicked on client: {actual_client_name}")
            except Exception as click_error:
                print(f"    Regular click failed, trying JavaScript click: {click_error}")
                driver.execute_script("arguments[0].click();", client_link)
                print(f"     JavaScript clicked on client: {actual_client_name}")
            
            # Wait for client page to load completely
            print(f"    Waiting for client page to load...")
            time.sleep(5)
            
            # Close modal if it appears
            modal_closed = close_modal_if_present()
            
            # Extract client information before clicking Electronic Records
            client_info = extract_client_info()
            
            # Click on Electronic Records tab after closing modal
            electronic_records_clicked = click_electronic_records()
            
            # Process electronic records if tab was clicked successfully
            if electronic_records_clicked:
                electronic_records_processed, treatment_record_folders = process_electronic_records(client_info)
                
                # CLEANUP: Close any open modals after processing all records
                print("    Performing cleanup after processing electronic records...")
                cleanup_open_modals()
                
            else:
                electronic_records_processed = False
                treatment_record_folders = []
            
            # Record the processing result
            processed_clients.append({
                'name': actual_client_name,
                'letter': letter,
                'processed': True,
                'modal_appeared': modal_closed,
                'client_name': client_info['name'],
                'client_id': client_info['id'],
                'folder_name': client_info['folder_name'],
                'electronic_records_visited': electronic_records_clicked,
                'electronic_records_processed': electronic_records_processed,
                'treatment_record_folders': treatment_record_folders,
                'client_number': client_index + 1,
                'page': page_number,
                'position_on_page': page_local_index + 1
            })
            
            # Navigate back to client list through menu
            if not robust_navigate_to_client_list():
                print(f"     Could not return to Client List. Skipping {actual_client_name}")
                client_index += 1
                continue
            
            print(f"     Client #{client_index + 1} ({actual_client_name}) processed successfully")
            
            # Move to next client
            client_index += 1
            
            # Small delay between clients
            time.sleep(2)
            
        except Exception as e:
            print(f"     Error processing client at index {client_index}: {e}")
            import traceback
            traceback.print_exc()
            
            # Try to recover by navigating back to client list
            try:
                print(f"    Attempting recovery...")
                navigate_to_client_list()
                print(f"     Recovery successful, moving to next client")
                client_index += 1
            except:
                print(f"     Recovery failed, moving to next client anyway")
                client_index += 1
            
            time.sleep(3)
            continue
    
    print(f"\nFinal count for letter {letter}: {len(processed_clients)} clients processed")
    print(f"   Started from client #{start_client_number}")
    print(f"   Ended at client #{client_index}")
    return processed_clients


def get_client_count():
    """Get the number of client rows currently displayed"""
    try:
        client_rows = driver.find_elements(By.CSS_SELECTOR, "#tblClientLeadListBody tr:not(.dataTables_empty)")
        return len(client_rows)
    except:
        return -1

# --- Auto-detect Chrome binary ---
default_chrome_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
chrome_path = default_chrome_path if os.path.exists(default_chrome_path) else shutil.which("google-chrome")

if not chrome_path:
    raise Exception("Chrome binary not found. Please install Chrome or specify the path manually.")

# --- Auto-detect ChromeDriver binary ---
chromedriver_path = shutil.which("chromedriver")
if not chromedriver_path:
    raise Exception("ChromeDriver not found in PATH. Please install it or add to PATH.")

# Setup Chrome options
options = webdriver.ChromeOptions()
options.binary_location = chrome_path
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])

# Configure Chrome to auto-download PDFs
prefs = {
    "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":""}],"selectedDestinationId":"Save as PDF","version":2}',
    "savefile.default_directory": os.path.abspath(MAIN_FOLDER),  # Set default download directory
    "download.default_directory": os.path.abspath(MAIN_FOLDER),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "plugins.always_open_pdf_externally": True,  # Download PDFs instead of opening
    "printing.default_destination_selection_rules": {
        "kind": "local",
        "namePattern": "Save as PDF",
    }
}
options.add_experimental_option("prefs", prefs)

# Add argument to handle print dialog
options.add_argument("--kiosk-printing")  # This enables automatic printing without dialog

# Initialize WebDriver
driver = webdriver.Chrome(service=Service(chromedriver_path), options=options)
wait = WebDriverWait(driver, 10)

try:
    # Login process
    print("Starting AestheticsPro Web Scraper...")
    print("PyAutoGUI is configured for handling save dialogs")
    print("Note: Move mouse to top-left corner to abort if needed")
    print("-" * 60)
    
    print("Logging in...")
    driver.get("https://www.aestheticspro.com/Login/")
    time.sleep(2)

    username = os.getenv("USERNAME")
    password = os.getenv("PASSWORD")

    # Enter credentials
    driver.find_element(By.ID, "username").send_keys(username)
    driver.find_element(By.ID, "password").send_keys(password)
    driver.find_element(By.ID, "btnSignin").click()

    # Wait for login to complete
    time.sleep(5)
    print("Current URL after login:", driver.current_url)

    # Navigate to Client List initially
    print("Navigating to Client List...")
    navigate_to_client_list()

    # Process user-selected letter and starting client
    all_processed_clients = []
    
    # letter input from user
    selected_letter = input("Enter the letter of the client names to process (A–Z): ").upper().strip()
    
    if len(selected_letter) != 1 or not selected_letter.isalpha():
        print("Invalid input. Please enter a single letter from A to Z.")
    else:
        # Get starting client number from user
        start_client_input = input(f"Enter the starting client number for letter {selected_letter} (default is 1): ").strip()
        
        # Parse the starting client number
        try:
            start_client_number = int(start_client_input) if start_client_input else 1
            if start_client_number < 1:
                print("Invalid client number. Starting from client 1.")
                start_client_number = 1
        except ValueError:
            print("Invalid input. Starting from client 1.")
            start_client_number = 1
        
        print(f"\nWill process letter {selected_letter} starting from client #{start_client_number}")
        
        # Calculate which page this will start on
        start_page = ((start_client_number - 1) // 100) + 1
        position_on_page = ((start_client_number - 1) % 100) + 1
        print(f"   This is page {start_page}, position {position_on_page} on that page")
        
        # Process clients
        processed_clients = process_clients_for_letter(selected_letter, start_client_number)
        all_processed_clients.extend(processed_clients)
        
        print(f"\n{'='*40}")
        print(f"Letter {selected_letter} Summary:")
        print(f"Started from: Client #{start_client_number}")
        print(f"Processed: {len(processed_clients)} clients")
        print(f"{'='*40}")


    # Step 4: Final Summary
    print(f"\n{'='*60}")
    print(f"SCRAPING COMPLETED!")
    print(f"{'='*60}")
    print(f"Total clients processed: {len(all_processed_clients)}")
    
    # Group by letter for summary
    letter_counts = {}
    for client in all_processed_clients:
        letter = client['letter']
        letter_counts[letter] = letter_counts.get(letter, 0) + 1
    
    print("\nClients processed per letter:")
    for letter in sorted(letter_counts.keys()):
        print(f"  {letter}: {letter_counts[letter]} clients")

    # Count clients with modals
    modal_count = sum(1 for client in all_processed_clients if client.get('modal_appeared', False))
    print(f"\nClients with modals: {modal_count}")

    # Optional: Save to file
    save_to_file = input("\nDo you want to save the processing log to a CSV file? (y/n): ").lower().strip()
    if save_to_file == 'y':
        import csv
        filename = f"aestheticspro_processed_clients_{int(time.time())}.csv"
        
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['letter', 'name', 'processed', 'modal_appeared', 'client_name', 'client_id', 'folder_name', 'electronic_records_visited', 'electronic_records_processed', 'treatment_record_folders', 'index']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            # Convert treatment_record_folders list to string for CSV
            for client in all_processed_clients:
                if 'treatment_record_folders' in client and isinstance(client['treatment_record_folders'], list):
                    client['treatment_record_folders'] = '; '.join(client['treatment_record_folders'])
            
            writer.writerows(all_processed_clients)
        
        print(f"Processing log saved to: {filename}")

    print("\nBrowser will remain open. Press CTRL+C or close the window to quit.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("Exiting...")

finally:
    # Uncomment the next line if you want to automatically close the browser
    # driver.quit()
    pass
