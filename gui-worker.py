import customtkinter as ctk
import json
import os
import threading
import time
import sys
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import tkinter as tk
import subprocess  # Import subprocess for opening files

# Global variable to track progress
progress = {'current': 0, 'total': 0}

def get_resource_path(filename):
    """Return the path to a resource file."""
    if getattr(sys, 'frozen', False):  # If the application is bundled
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)

def is_file_open(file_path):
    """Check if the file is currently open by attempting to open it exclusively."""
    try:
        with open(file_path, 'a'):
            pass
        return False
    except IOError:
        return True

def read_search_policies_from_file(filename):
    try:
        with open(filename, 'r', encoding='utf8') as file:
            data = json.load(file)
            return data.get('policies', [])
    except Exception as e:
        print(f'Error reading JSON file: {e}')
        return []

def export_to_excel(results):
    """Export the results to an Excel file, saving it in the same directory as the executable."""
    exe_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(exe_dir, 'result/result.xlsx')
    print(f"Excel file saved in {excel_path}")

    # Check if the file already exists and is open
    if os.path.exists(excel_path):
        if is_file_open(excel_path):
            # Show a message box if the file is open
            root = tk.Tk()
            root.withdraw()  # Hide the main window
            messagebox.showwarning("File in Use", f"The file '{excel_path}' is currently open. Please close it before proceeding.")
            return

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = 'Assigned Producers'

    # Add headers
    worksheet.append(['Policy', 'Assigned Producer'])

    # Add data rows
    for result in results:
        worksheet.append([result['searchPolicy'], result['assignedProducer']])

    # Save workbook
    workbook.save(excel_path)
    print(f'Excel file saved successfully at {excel_path}.')

def update_progress(current, total):
    """Update global progress variable."""
    global progress
    progress = {'current': current, 'total': total}
    root.after(100, update_progress_display)  # Schedule update for the main thread

def wait_for_login(driver, wait):
    login_button_selector = By.ID, 'btnLogin'
    target_url = 'https://app.ezlynx.com/applicantportal/Commissions/Statements'

    while True:
        current_url = driver.current_url
        if current_url == target_url:
            print('Already logged in or redirected to the correct page.')
            break

        try:
            # Check if login button is present
            wait.until(EC.presence_of_element_located(login_button_selector))
            print('Login button found. Please log in.')
            # Wait a bit before checking again
            time.sleep(2)
        except Exception:
            print('Login button not found or another issue.')
            # Continue checking the URL
            time.sleep(2)

def process_policies():
    # Create user data directory
    user_data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'User Data')
    if not os.path.exists(user_data_dir):
        os.makedirs(user_data_dir)

    # Chrome options and setup
    options = Options()
    options.add_argument(f'user-data-dir={user_data_dir}')
    options.add_argument('profile-directory=Default')

    # Set up WebDriver
    service = Service(get_resource_path('chromedriver-win64/chromedriver.exe'))
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 10)

    try:
        driver.get('https://app.ezlynx.com/applicantportal/Commissions/Statements')

        # Wait for the user to be logged in
        wait_for_login(driver, wait)
        
        driver.get('https://app.ezlynx.com/applicantportal/Commissions/Statements')

        # Read search policies from JSON file
        search_policies = read_search_policies_from_file(get_resource_path('search_policies.json'))
        total_policies = len(search_policies)

        # List to store results
        results = []

        for idx, policy in enumerate(search_policies):
            search_input = wait.until(EC.presence_of_element_located((By.ID, 'quickSearchInput')))
            search_input.clear()
            search_input.send_keys(policy)

            # Wait for search results to load
            time.sleep(2)  # Adjust as necessary

            assigned_producer_name = 'Missing Assigned Producer'

            try:
                assigned_producer_element = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '*[id^="mat-option-"] > span > a > span:nth-child(4)'))
                )
                assigned_producer_name = assigned_producer_element.text.replace('Assigned Producer: ', '')
            except Exception:
                assigned_producer_name = 'Missing Assigned Producer'

            # Custom transformation logic
            if assigned_producer_name == 'Accounting Unidentified':
                assigned_producer_name = 'PIA Select Staff'
            elif assigned_producer_name == 'Anthony R':
                assigned_producer_name = 'PIA Select Staff'
            elif assigned_producer_name == 'Nagamani GarikaCSR':
                assigned_producer_name = 'Kavita Sood'

            print(f'Policy: {policy}, Assigned Producer: {assigned_producer_name}')

            # Store the result in the array
            results.append({'searchPolicy': policy, 'assignedProducer': assigned_producer_name})

            # Update progress
            update_progress(idx + 1, total_policies)

        # Export all results to Excel
        export_to_excel(results)
        # Final progress update
        update_progress(total_policies, total_policies)

    finally:
        driver.quit()

def on_button_click():
    # Get the text from the text area
    text_content = text_area.get("1.0", "end-1c").strip()
    if not text_content:
        messagebox.showwarning("Input Error", "No policies found in the text area.")
        return

    # Create the JSON file
    policies = text_content.splitlines()
    json_data = {"policies": policies}
    with open(get_resource_path('search_policies.json'), 'w') as file:
        json.dump(json_data, file)

    # Check if there are policies to process
    if len(policies) == 0:
        hide_progress_components()  # Hide progress components if no policies
        button.configure(state='normal')
        open_result_button.configure(state='normal')
        return

    # Show and initialize the progress bar
    show_progress_components()
    progress_bar.set(0)
    
    # Disable the buttons and start progress
    button.configure(state='disabled')
    open_result_button.configure(state='disabled')

    # Start worker script in a separate thread
    threading.Thread(target=process_policies).start()

def update_progress_display():
    """Update the progress bar and progress count with the current progress."""
    current = progress.get('current', 0)
    total = progress.get('total', 1)  # Avoid division by zero
    if total > 0:
        progress_bar.set(current / total)
    progress_label.configure(text=f'{current}/{total}')  # Use configure instead of config
    if current >= total:
        # Hide the progress bar and re-enable the buttons
        hide_progress_components()
        button.configure(state='normal')
        open_result_button.configure(state='normal')
        return
    # Schedule the next update check
    root.after(1000, update_progress_display)

def show_progress_components():
    """Show the progress bar and progress count label."""
    progress_frame.grid(row=3, column=0, columnspan=2, padx=20, pady=(10, 5), sticky='ew')

def hide_progress_components():
    """Hide the progress bar and progress count label and reset progress values."""
    progress_frame.grid_forget()
    # Reset progress values
    global progress
    progress = {'current': 0, 'total': 0}

def open_result_file():
    """Open the result Excel file with the default system application."""
    excel_path = get_resource_path('result/result.xlsx')
    if os.path.exists(excel_path):
        if sys.platform == "win32":  # Windows
            os.startfile(excel_path)
        elif sys.platform == "darwin":  # macOS
            subprocess.call(('open', excel_path))
        else:  # Linux
            subprocess.call(('xdg-open', excel_path))
    else:
        messagebox.showerror("File Not Found", "The result file does not exist.")

# Create the main window
root = ctk.CTk()  # Use CTk for the main window
root.title("Producer Finder")
root.resizable(False, False)

# Set the dark mode theme
ctk.set_appearance_mode("dark")  # 'light' for light mode

# Create a Text Area with larger dimensions
text_area = ctk.CTkTextbox(root, height=200, width=350, wrap="word")
text_area.grid(row=0, column=0, columnspan=2, padx=20, pady=20)

# Create a Find Button
button = ctk.CTkButton(root, text="Find", command=on_button_click)
button.grid(row=1, column=0, pady=10, padx=(20, 5), sticky='ew')

# Create an Open Result Button
open_result_button = ctk.CTkButton(root, text="Open Result", command=open_result_file)
open_result_button.grid(row=1, column=1, pady=10, padx=(5, 20), sticky='ew')

# Center the buttons in the grid
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

# Create a frame to contain the progress bar
progress_frame = ctk.CTkFrame(root)

# Create a Progress Bar with a specified width
progress_bar = ctk.CTkProgressBar(progress_frame, orientation="horizontal", width=300)

# Create a Progress Label
progress_label = ctk.CTkLabel(progress_frame, text="0/0")

# Layout the progress bar and label inside the frame
progress_bar.grid(row=0, column=0, padx=10, pady=5)
progress_label.grid(row=0, column=1, padx=10, pady=5)

# Start the GUI event loop
root.after(1000, update_progress_display)  # Ensure periodic updates for progress
root.mainloop()
