from pystray import MenuItem as item
from PIL import Image, ImageDraw
import urllib.request
import customtkinter
import configparser
import subprocess
import threading
import datetime
import tzlocal
import pystray
import socket
import ntplib
import ctypes
import winreg
import time
import pytz
import os

# Check if app has registry key to run on startup
def check_registry_key():
    # Registry key information
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    file_name = os.path.splitext(os.path.basename(__file__))[0]

    # Check if the registry key exists
    key_exists = False
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ)
        try:
            winreg.QueryValueEx(key, file_name)
            key_exists = True
        except FileNotFoundError:
            pass
        winreg.CloseKey(key)
    except FileNotFoundError:
        pass

    return key_exists

# Call the check_registry_key() function to check if the key exists
if os.path.isfile("addregkey.exe"):
    if not check_registry_key():
        print("The registry key does not exist, trying to add registry key...")
        command = f"addregkey.exe {os.path.basename(__file__)}"
        os.popen(command)
    else:
        print("The registry key already exists.")
else:
    print("'addregkey.exe' not found, trying to add registry key failed.")

# Format the internet time
def format_internet_time(internet_time):
    format_str = '%Y-%m-%d %I:%M:%S %p' if internet_time.strftime('%p') else '%Y-%m-%d %H:%M:%S'
    return internet_time.strftime(format_str)

# Show popup if current time does not match time from internet
def show_time_mismatch_popup(internet_time):
    popup = customtkinter.CTkToplevel()
    popup.title("Time Mismatch")
    popup.geometry("450x200")
    popup.attributes('-topmost', True)

    label1 = customtkinter.CTkLabel(master=popup, text="The current time on your device doesn't match the time from the internet",
                                   font=customtkinter.CTkFont(size=14))
    label1.pack(pady=(10, 2))

    label2_text = "Internet Time: {}".format(format_internet_time(internet_time))
    label2 = customtkinter.CTkLabel(master=popup, text=label2_text, font=customtkinter.CTkFont(size=14), wraplength=400)
    label2.pack(pady=2)

    label3_text = "Please note that the displayed internet time might also be incorrect for your region."
    label3 = customtkinter.CTkLabel(master=popup, text=label3_text, font=customtkinter.CTkFont(size=13), wraplength=400)
    label3.pack(pady=(10, 15))

    checkbox_var = customtkinter.IntVar()
    checkbox = customtkinter.CTkCheckBox(master=popup, text="Don't show this message again", variable=checkbox_var,
                                         font=customtkinter.CTkFont(size=12), width=4, height=3,
                                         checkbox_width=16, checkbox_height=16, border_width=2)
    checkbox.pack(pady=(5, 0))

    def save_settings_and_close():
        config = configparser.ConfigParser()
        config.read(config_file)
        if 'POPUP_SETTINGS' not in config:
            config['POPUP_SETTINGS'] = {}
        if checkbox_var.get() == 1:
            config['POPUP_SETTINGS']['ShowPopup'] = 'False'

        # Save the modified settings to the config file
        with open(config_file, 'w') as configfile:
            config.write(configfile)
        if popup.winfo_exists():
            popup.destroy()

    ok_button = customtkinter.CTkButton(master=popup, text="OK", command=save_settings_and_close)
    ok_button.pack(pady=(5, 10))

# Show popup to notify user that the app is in the tray
def show_tray_icon_popup():
    notify = customtkinter.CTkToplevel()
    notify.title("App is in the tray")
    notify.geometry("300x140")
    notify.attributes('-topmost', True)

    label1 = customtkinter.CTkLabel(master=notify, text="Clock Out is now running in the system tray. Click the Clock Out icon to access it.",
        font=customtkinter.CTkFont(size=14), wraplength=270)
    label1.pack(pady=(30, 2))

    ok_button = customtkinter.CTkButton(master=notify, text="OK", command=notify.destroy)
    ok_button.pack(pady=(10, 10))

# Show popup to notify user the computer will hibernate
def show_hibernate_popup():
    notify = customtkinter.CTkToplevel()
    notify.title("Notify")
    notify.geometry("300x140")
    notify.attributes('-topmost', True)

    label1 = customtkinter.CTkLabel(master=notify, text="Your allocated computer usage time has ended. The system will automatically hibernate in few seconds.",
        font=customtkinter.CTkFont(size=14), wraplength=270)
    label1.pack(pady=(30, 2))

    ok_button = customtkinter.CTkButton(master=notify, text="OK", command=notify.destroy)
    ok_button.pack(pady=(10, 10))

# Setup config and config_file file
config = configparser.ConfigParser()
config_file = os.path.splitext(__file__)[0] + ".ini"

# Validate the start minute entry box
def validate_start_minute(*args):
    minute = start_minute_var.get()
    if not minute.isdigit():
        start_minute_var.set("00")
    else:
        minute = int(minute)
        if minute < 0:
            minute = abs(minute)
            if minute > 59:
                minute = 59
            start_minute_var.set(str(minute).zfill(2))
        elif minute > 59:
            start_minute_var.set("59")
        else:
            start_minute_var.set(str(minute).zfill(2))
    update_start_time(None)

# Validate the end minute entry box
def validate_end_minute(*args):
    minute = end_minute_var.get()
    if not minute.isdigit():
        end_minute_var.set("00")
    else:
        minute = int(minute)
        if minute < 0:
            minute = abs(minute)
            if minute > 59:
                minute = 59
            end_minute_var.set(str(minute).zfill(2))
        elif minute > 59:
            end_minute_var.set("59")
        else:
            end_minute_var.set(str(minute).zfill(2))
    update_end_time(None)

# Shutdown device in 5 minutes
def shutdown_device():
    command = 'shutdown /s /t 300 /c "Your allocated computer usage time has ended. The system will automatically shut down in 5 minutes. Please save your work."'
    subprocess.Popen(command, shell=True, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP)

# Hibernate device in few seconds
def hibernate_device():
    show_hibernate_popup()
    
    # Define the delay function
    def delay():
        # Wait for 30 seconds
        time.sleep(30)
        
        # Execute the "shutdown /h" command to hibernate the device
        subprocess.Popen(['shutdown', '/h'])
    
    # Start the delay function in a separate thread
    delay_thread = threading.Thread(target=delay)
    delay_thread.start()

# Check if device can be hibernated
def check_hibernate_state():
    hibernate_state = ctypes.windll.powrprof.IsPwrHibernateAllowed()
    return bool(hibernate_state)

# Check if the device has to shutdown or hibernate
def hibernate_or_shutdown():
    has_hibernate = check_hibernate_state()
    print(f"Computer has hibernate option: {has_hibernate}")
    # hibernate if possible else shutdown
    if has_hibernate:
        hibernate_device()
    else:
        shutdown_device()

# Check if the device should be running or not
def is_within_time_range(start_hour, start_minute, start_am_pm, end_hour, end_minute, end_am_pm):
    # Get the current time
    current= datetime.datetime.now().time()
    current_time = datetime.time(current.hour, current.minute)

    # Extract the values from StringVar objects and convert them to integers
    start_hour_val = int(start_hour.get())
    start_minute_val = int(start_minute.get())
    start_am_pm_val = start_am_pm.get()
    end_hour_val = int(end_hour.get())
    end_minute_val = int(end_minute.get())
    end_am_pm_val = end_am_pm.get()

    # Convert start and end time to datetime objects
    if start_am_pm_val == "PM" and start_hour_val != 12:
        start_hour_val += 12
    elif start_am_pm_val == "AM" and start_hour_val == 12:
        start_hour_val = 0

    if end_am_pm_val == "PM" and end_hour_val != 12:
        end_hour_val += 12
    elif end_am_pm_val == "AM" and end_hour_val == 12:
        end_hour_val = 0

    start_time = datetime.time(start_hour_val, start_minute_val)
    end_time = datetime.time(end_hour_val, end_minute_val)
    print(start_time, current_time, end_time)

    # Check if the current time is within the specified range
    if start_time <= current_time <= end_time:
        return True
    else:
        return False

# Loop to check if the device is within time range
def run_loop():
    while True:
        if is_within_time_range(start_hour_var, start_minute_var, start_am_pm_var, end_hour_var, end_minute_var, end_am_pm_var):
            print("Computer is within the specified time range.")
        else:
            print("Computer is outside the specified time range.")
            hibernate_or_shutdown()

        # Delay for a few seconds before checking again
        time.sleep(60)

# Create a separate thread for running the loop
loop_thread = threading.Thread(target=run_loop)

# Function to check if the loop thread is running
def is_loop_thread_running():
    return loop_thread.is_alive()

# Create a separate thread for running the loop
loop_thread = threading.Thread(target=run_loop)

# Validate and adjust minutes if necessary
def validate_minutes():
    start_minute = int(start_minute_var.get())
    end_minute = int(end_minute_var.get())

    if start_minute < 0:
        start_minute_var.set(str(-start_minute))
    elif start_minute > 59:
        start_minute_var.set("59")

    if end_minute < 0:
        end_minute_var.set(str(-end_minute))
    elif end_minute > 59:
        end_minute_var.set("59")

# Handle the "Save" button click event
def save_settings():
    validate_minutes()
    config = configparser.ConfigParser()
    config.read(config_file)

    # Check if the 'GENERAL_SETTINGS' section exists
    if 'GENERAL_SETTINGS' not in config:
        config['GENERAL_SETTINGS'] = {}

    # Save the other settings
    config['GENERAL_SETTINGS']['StartHour'] = start_hour_var.get()
    config['GENERAL_SETTINGS']['StartMinute'] = start_minute_var.get()
    config['GENERAL_SETTINGS']['StartAMPM'] = start_am_pm_var.get()
    config['GENERAL_SETTINGS']['EndHour'] = end_hour_var.get()
    config['GENERAL_SETTINGS']['EndMinute'] = end_minute_var.get()
    config['GENERAL_SETTINGS']['EndAMPM'] = end_am_pm_var.get()
    config['GENERAL_SETTINGS']['Theme'] = theme_combobox.get()

    # Save the modified settings to the config file
    with open(config_file, 'w') as configfile:
        config.write(configfile)

    print("Settings saved.")

    if not is_loop_thread_running():
        loop_thread.start()

    root.withdraw() # Hide the root window
    show_tray_icon_popup() # Show popup to notify app is in system tray

# Check internet connectivity
def check_internet_connection():
    try:
        # Check if a connection to Google's DNS server can be established
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        return True
    except OSError:
        pass
    return False

# Get the current internet time using NTP
def get_internet_time():
    try:
        ntp_client = ntplib.NTPClient()
        response = ntp_client.request('time.google.com', version=3, timeout=3)
        internet_time = datetime.datetime.fromtimestamp(response.tx_time, pytz.utc)
        return internet_time
    except ntplib.NTPException:
        return None

# Compare local time with internet time
def compare_time():
    if check_internet_connection():
        internet_time = get_internet_time()
        if internet_time:
            local_tz = tzlocal.get_localzone()
            local_time = datetime.datetime.now(pytz.utc)
            local_time = local_time.replace(second=0, microsecond=0)  # Ignore seconds and microseconds
            internet_time = internet_time.replace(second=0, microsecond=0)  # Ignore seconds and microseconds
            if local_time.astimezone(internet_time.tzinfo) != internet_time:
                show_time_mismatch_popup(internet_time)
    else:
        print("No internet connection. Skipping time comparison.")

# Load the settings from the config file, if it exists
def load_settings():
    if os.path.isfile(config_file):
        config.read(config_file)
        if 'GENERAL_SETTINGS' in config:
            settings = config['GENERAL_SETTINGS']
            start_hour_var.set(settings.get('StartHour', '06'))
            start_minute_var.set(settings.get('StartMinute', '00'))
            start_am_pm_var.set(settings.get('StartAMPM', 'AM'))
            end_hour_var.set(settings.get('EndHour', '12'))
            end_minute_var.set(settings.get('EndMinute', '00'))
            theme_combobox.set(settings.get('Theme', 'System'))
            print("Settings loaded")
        else:
            print("No 'GENERAL_SETTINGS' section found in the config file.")

        if 'POPUP_SETTINGS' in config:
            popup_settings = config['POPUP_SETTINGS']
            if 'ShowPopup' in popup_settings and popup_settings.get('ShowPopup') == 'False':
                # Don't show the popup if the setting is set to False
                return
        else:
            print("No 'POPUP_SETTINGS' section found in the config file.")

        # Call the compare_time function to trigger the popup
        compare_time()
    else:
        print("Config file not found. Using default settings.")
        compare_time()

# Check if 'appicon.ico' file exists
if not os.path.exists("appicon.ico"):
    print("Trying to downloading 'appicon.ico'.")
    try:
        # Download the file from the provided URL
        url = "https://raw.githubusercontent.com/TechWhizKid/ClockOut/main/appicon.ico"
        urllib.request.urlretrieve(url, "appicon.ico")
        print("'appicon.ico' downloaded successfully.")
    except:
        def create_icon(icon_path, icon_size=(256, 256), background_color=(220, 220, 220), cross_color=(178, 34, 34)):
            # Create a new blank image with an alpha channel
            icon = Image.new("RGBA", icon_size, (0, 0, 0, 0))

            # Draw the background color
            icon.paste(background_color, [0, 0, icon_size[0], icon_size[1]])

            # Calculate the position and size of the cross
            cross_size = min(icon_size) // 2
            cross_position = ((icon_size[0] - cross_size) // 2, (icon_size[1] - cross_size) // 2)

            # Draw the cross
            draw = ImageDraw.Draw(icon)
            line_width = 8  # Adjust the line width as desired
            draw.line((cross_position[0], cross_position[1], cross_position[0] + cross_size - 1, cross_position[1] + cross_size - 1), fill=cross_color, width=line_width)
            draw.line((cross_position[0], cross_position[1] + cross_size - 1, cross_position[0] + cross_size - 1, cross_position[1]), fill=cross_color, width=line_width)

            # Save the icon as an ICO file
            icon.save(icon_path, format="ICO")
        
        icon_path = "appicon.ico"
        create_icon(icon_path)

# Set CustomTkinter appearance and theme
customtkinter.set_appearance_mode("system")
customtkinter.set_default_color_theme("dark-blue")

# Create CustomTkinter window
root = customtkinter.CTk()
root.geometry("360x385")
root.title("Clock Out")

# Create a frame in the main window
frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

# Create app title label
label_clock_out = customtkinter.CTkLabel(master=frame, text="Clock Out",
    font=customtkinter.CTkFont(size=20, weight="bold"))
label_clock_out.pack(pady=12, padx=10)

# Create frame for the clock
clock_frame = customtkinter.CTkFrame(master=frame)
clock_frame.pack(pady=12)

# Create variables to store the selected start time
start_hour_var = customtkinter.StringVar(value="06")
start_minute_var = customtkinter.StringVar(value="00")
start_am_pm_var = customtkinter.StringVar(value="AM")

# Update the start time based on the selected values
def update_start_time():
    hour = start_hour_var.get()
    minute = start_minute_var.get()
    am_pm = start_am_pm_var.get()
    time_str = f"{hour}:{minute} {am_pm}"
    print("Start Time:", time_str)

# Create "Day starts at:" label
label_start = customtkinter.CTkLabel(master=clock_frame, text="Day starts at:", font=customtkinter.CTkFont(size=15, weight="bold"))
label_start.pack()
start_frame = customtkinter.CTkFrame(master=clock_frame)
start_frame.pack()

# Create start hour combobox
start_hour_combobox = customtkinter.CTkComboBox(master=start_frame, values=[str(i).zfill(2) for i in range(1, 13)],
    variable=start_hour_var, font=("Arial", 14), width=65, command=update_start_time)
start_hour_combobox.pack(side="left", padx=(4, 0))

# Create colon label
label_colon_start = customtkinter.CTkLabel(master=start_frame, text=":", font=customtkinter.CTkFont(size=18, weight="bold"))
label_colon_start.pack(side="left")

# Create start minute entry box
start_minute_entry = customtkinter.CTkEntry(master=start_frame, textvariable=start_minute_var, font=("Arial", 14), width=65)
start_minute_entry.pack(side="left")

# Create start AM/PM combobox
start_am_pm_combobox = customtkinter.CTkComboBox(master=start_frame, values=["AM", "PM"], variable=start_am_pm_var,
    font=("Arial", 14), width=65, command=update_start_time)
start_am_pm_combobox.pack(side="left", padx=(4, 0))

# Bind the event to validate the start minute entry box
start_minute_entry.bind("<FocusOut>", validate_start_minute)

# Create variables to store the selected end time
end_hour_var = customtkinter.StringVar(value="11")
end_minute_var = customtkinter.StringVar(value="00")
end_am_pm_var = customtkinter.StringVar(value="PM")

# Update the end time based on the selected values
def update_end_time():
    hour = end_hour_var.get()
    minute = end_minute_var.get()
    am_pm = end_am_pm_var.get()
    time_str = f"{hour}:{minute} {am_pm}"
    print("End Time:", time_str)

# Create "Day ends at:" label
label_end = customtkinter.CTkLabel(master=clock_frame, text="Day ends at:", font=customtkinter.CTkFont(size=15, weight="bold"))
label_end.pack()
end_frame = customtkinter.CTkFrame(master=clock_frame)
end_frame.pack()

# Create end hour combobox
end_hour_combobox = customtkinter.CTkComboBox(master=end_frame, values=[str(i).zfill(2) for i in range(1, 13)],
    variable=end_hour_var, font=("Arial", 14), width=65, command=update_end_time)
end_hour_combobox.pack(side="left", padx=(4, 0))

# Create colon label
label_colon_end = customtkinter.CTkLabel(master=end_frame, text=":", font=customtkinter.CTkFont(size=18, weight="bold"))
label_colon_end.pack(side="left")

# Create end minute entry box
end_minute_entry = customtkinter.CTkEntry(master=end_frame, textvariable=end_minute_var, font=("Arial", 14), width=65)
end_minute_entry.pack(side="left")

# Create end AM/PM combobox
end_am_pm_combobox = customtkinter.CTkComboBox(master=end_frame, values=["AM", "PM"], variable=end_am_pm_var,
    font=("Arial", 14), width=65, command=update_end_time)
end_am_pm_combobox.pack(side="left", padx=(4, 0))

# Bind the event to validate the end minute entry box
end_minute_entry.bind("<FocusOut>", validate_end_minute)

# Create the "Save" button
save_button = customtkinter.CTkButton(master=clock_frame, text="Save", command=save_settings)
save_button.pack(pady=12)

# Change theme
def change_theme():
    theme = theme_combobox.get()
    if theme == "Light":
        customtkinter.set_appearance_mode("light")
    elif theme == "Dark":
        customtkinter.set_appearance_mode("dark")
    else:
        customtkinter.set_appearance_mode("system")

# Create theme selection combobox
theme_combobox = customtkinter.CTkComboBox(frame, values=["System", "Light", "Dark"], width=82)
theme_combobox.pack(pady=10)

# Create theme change button
theme_button = customtkinter.CTkButton(frame, text="Change Theme", command=change_theme)
theme_button.pack(pady=10)

load_settings() # Call the load_settings() function to load the settings on startup
change_theme() # Call the change_theme() function to load the theme on startup

# Create-System-Tray-Icon------------------------

# Show the main window
def on_tray_icon_clicked():
    root.deiconify()  # Restore the window when tray icon is clicked

# Exit app
def exit_app():
    tray_icon.stop() # Stop the tray icon
    os._exit(0) # Exit the mainloop

# Create the menu items for the system tray icon
menu = (
    item('Open', on_tray_icon_clicked),
    item('Exit', exit_app)
)

# Load the icon image
icon_image = Image.open('appicon.ico')

# Create the system tray icon
tray_icon = pystray.Icon("Clock Out", icon_image, "Clock Out", menu)

# Minimize the window to the system tray on closing
def on_closing():
    root.withdraw()
    show_tray_icon_popup()

# Bind the closing event to the minimize function
root.protocol("WM_DELETE_WINDOW", on_closing)

# Run the tray icon on a different thread
tray_icon.run_detached()

if os.path.isfile(config_file):
    config.read(config_file)
    if 'GENERAL_SETTINGS' in config:
        if not is_loop_thread_running():
            loop_thread.start()
        root.withdraw()
else:
    pass # Config file not found. Using default settings.

# Run the main loop
root.mainloop()
