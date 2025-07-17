import os
import re
import shutil
import subprocess
from datetime import datetime

# === AppleScript Prompt Functions ===

def prompt_with_list(title, prompt, options):
    list_items = "{" + ", ".join([f'\"{item}\"' for item in options]) + "}"
    script = f'set theList to {list_items}\nchoose from list theList with prompt "{prompt}" with title "{title}"'
    result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
    choice = result.stdout.strip().strip('"')
    return choice if choice.lower() != "false" else None

def prompt_with_list_or_custom(title, prompt, options, custom_label="Other..."):
    options_with_custom = options + [custom_label]
    choice = prompt_with_list(title, prompt, options_with_custom)
    if choice == custom_label:
        # AppleScript text input dialog
        script = f'display dialog "Enter custom value:" with title "{title}" default answer ""'
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
        # Parse AppleScript output, look for 'text returned:'
        m = re.search(r'text returned:(.*)', result.stdout)
        return m.group(1).strip() if m else ""
    return choice

def prompt_yes_no(title, prompt):
    script = f'display dialog "{prompt}" with title "{title}" buttons {{"No", "Yes"}} default button "No"'
    result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
    return "Yes" in result.stdout

# === Helper: Extract First Name from "Last, First"
def get_first_name(full_name):
    if ',' in full_name:
        parts = full_name.split(',')
        return parts[1].strip()
    else:
        return full_name.split()[0]

# === Load or Prompt Technician List ===
technicians_file = os.path.expanduser("~/Downloads/PharosDMG/technicians.txt")
if os.path.exists(technicians_file):
    with open(technicians_file) as f:
        technician_display_list = [line.strip() for line in f if line.strip()]
    # Try to map to "Last, First" if using internal style
    technician_map = {}
    for t in technician_display_list:
        if ',' in t:
            # Already in "Last, First"
            name = t
            display = f"{name.split(',')[1].strip()} {name.split(',')[0].strip()}"
            technician_map[display] = name
        else:
            # Just use as display & internal
            technician_map[t] = t
    technician_display_list = list(technician_map.keys())
else:
    # Default entries
    technicians_internal = ["Croucher, Mike", "Tian, Zhiyong"]
    technician_map = {
        f"{name.split(',')[1].strip()} {name.split(',')[0].strip()}": name
        for name in technicians_internal
    }
    technician_display_list = list(technician_map.keys())

technician_display = prompt_with_list_or_custom("Technician", "Select your name:", technician_display_list)
technician_full = technician_map.get(technician_display, technician_display)

# === Personalized Warning ===
if technician_full:
    first_name = get_first_name(technician_full)
    warning_message = (
        f"Hello {first_name}, please ensure the correct manufacturer and model printer drivers "
        f"are installed on this Mac before proceeding.\n\n"
        f"Also, ensure you have downloaded and saved the latest Pharos Popup and Notify Client package."
        f"to your ~/Downloads/PharosDMG folder. This file should be renamed to Installer.pkg."
    )
    warning_script = f'display dialog "{warning_message}" with title "Driver Installation Required" buttons {{"OK"}} default button "OK"'
    subprocess.run(['osascript', '-e', warning_script])

# === Read Queues from File or Allow Custom ===
queues_file = os.path.expanduser("~/Downloads/PharosDMG/queues.txt")
if os.path.exists(queues_file):
    with open(queues_file, "r") as f:
        queues = [line.strip() for line in f if line.strip()]
else:
    queues = []

queue_name = prompt_with_list_or_custom("Queue", "Select the Pharos queue name:", queues)
if not queue_name:
    print("❌ No queue name selected or entered.")
    exit()

# === Generate Popup Name from Queue ===
popup_name = queue_name.replace("YesAuth_", "") + "_Popup"

# === Read Servers from File or Allow Custom ===
servers_file = os.path.expanduser("~/Downloads/PharosDMG/servers.txt")
if os.path.exists(servers_file):
    with open(servers_file, "r") as f:
        servers = [line.strip() for line in f if line.strip()]
else:
    servers = ["ps1.ohio.edu", "ps2.ohio.edu", "psb.ohio.edu"]

server = prompt_with_list_or_custom("Server", "Select the print server:", servers)
if not server:
    print("❌ No server selected or entered.")
    exit()

# === Get Available Drivers ===
resources_dir = "/Library/Printers/PPDs/Contents/Resources"
ppd_files = sorted([f for f in os.listdir(resources_dir) if f.endswith(".gz")])

def extract_manufacturer(filename):
    name = filename.replace(".PPD.gz", "").replace(".ppd.gz", "").replace(".gz", "")
    match = re.match(r"^([A-Z]{2,}(?=[^a-z]|$)|[A-Z][a-z]+|[A-Z][a-zA-Z]+(?=[A-Z]))", name)
    return match.group(1) if match else name.split()[0]

manufacturers = sorted(set(extract_manufacturer(f) for f in ppd_files))
manufacturer = prompt_with_list_or_custom("Manufacturer", "Select the printer manufacturer:", manufacturers)
if not manufacturer:
    print("❌ No manufacturer selected or entered.")
    exit()

filtered_drivers = [f[:-3] for f in ppd_files if extract_manufacturer(f) == manufacturer]
selected_driver = prompt_with_list_or_custom("Driver", "Select the printer model:", filtered_drivers)
if not selected_driver:
    print("❌ No driver selected or entered.")
    exit()

# === Validate Required Inputs ===
if not all([technician_full, queue_name, popup_name, selected_driver, server]):
    print("❌ Operation cancelled or incomplete selection.")
    exit()

# === PPD Setup with Sanitization ===
driver_filename = selected_driver if selected_driver.lower().endswith(".ppd") else selected_driver + ".ppd"
sanitized_driver_filename = driver_filename.replace(" ", "")

# === Directory Setup ===
current_date = datetime.now().strftime("%m/%d/%Y")
volume_name = queue_name
dmg_name = f"{volume_name}.dmg"
temp_dir = f"/tmp/{volume_name}"

if os.path.exists(temp_dir):
    shutil.rmtree(temp_dir)
custom_dir = os.path.join(temp_dir, "Custom")
os.makedirs(custom_dir, exist_ok=True)

# === Copy and Decompress PPD, using sanitized name ===
source_ppd_gz = os.path.join(resources_dir, selected_driver + ".gz")
destination_ppd = os.path.join(custom_dir, sanitized_driver_filename)
with open(destination_ppd, "wb") as out_f:
    subprocess.run(["gunzip", "-c", source_ppd_gz], stdout=out_f)

# === Copy Installer.pkg ===
installer_src = os.path.expanduser("~/Downloads/PharosDMG/Installer.pkg")
installer_dst = os.path.join(temp_dir, "Installer.pkg")
if os.path.exists(installer_src):
    subprocess.run(["cp", installer_src, installer_dst])
else:
    print("❌ Installer.pkg not found in ~/Downloads/PharosDMG")
    exit()

# === Create InstallFiles.txt, using sanitized name ===
install_txt_path = os.path.join(custom_dir, "InstallFiles.txt")
with open(install_txt_path, "w") as f:
    f.write("# PPD file need to be copied\n")
    f.write("#\n")
    f.write(f"# {technician_full} -- {current_date}\n")
    f.write("#\n")
    f.write(f"/etc/cups/ppd/{sanitized_driver_filename}\n")

# === Create PostInstall.sh, using sanitized name and the chosen server ===
postinstall_path = os.path.join(custom_dir, "PostInstall.sh")
# Remove planning unit prefix (up to and including the first '-')
if '-' in popup_name:
    popup_name_no_unit = popup_name.split('-', 1)[1]
else:
    popup_name_no_unit = popup_name

with open(postinstall_path, "w") as f:
    f.write("# Install print queue in CUPS\n")
    f.write("#\n")
    f.write(f"# {technician_full} -- {current_date}\n")
    f.write("#\n")
    f.write(f"lpadmin -p {popup_name_no_unit} -v popup://PS1.ohio.edu/{queue_name} -E -P {driver_filename}\n")

# === Create DMG ===
output_dir = os.path.expanduser("~/Downloads/PharosDMG")
os.makedirs(output_dir, exist_ok=True)
dmg_path = os.path.join(output_dir, dmg_name)

if os.path.exists(dmg_path):
    overwrite = prompt_yes_no("Overwrite DMG", f"The file {dmg_name} already exists. Overwrite?")
    if not overwrite:
        print("❌ Operation cancelled by user.")
        exit()

subprocess.run([
    "hdiutil", "create", "-volname", volume_name,
    "-srcfolder", temp_dir,
    "-ov", "-format", "UDZO", dmg_path
])
print(f"✅ DMG created at: {dmg_path}")
