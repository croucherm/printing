import os
import re
import shutil
import subprocess
from datetime import datetime

# === AppleScript Prompt Function ===
def prompt_with_list(title, prompt, options):
    list_items = "{" + ", ".join([f'\"{item}\"' for item in options]) + "}"
    script = f'set theList to {list_items}\nchoose from list theList with prompt "{prompt}" with title "{title}"'
    result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
    choice = result.stdout.strip().strip('"')  # Removes surrounding quotes
    return choice if choice.lower() != "false" else None

# === AppleScript Yes/No Prompt ===
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

# === Technician Setup ===
# Internal format: "Last, First"
technicians_internal = ["Croucher, Mike", "Tian, Zhiyong"]

# Display format: "First Last"
technician_map = {
    f"{name.split(',')[1].strip()} {name.split(',')[0].strip()}": name
    for name in technicians_internal
}
technician_display_list = list(technician_map.keys())

# === Prompt for Technician ===
technician_display = prompt_with_list("Technician", "Select your name:", technician_display_list)
technician_full = technician_map.get(technician_display)

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

# === Other Predefined Options ===
queues = ["YesAuth_OIT-BAKER-105", "YesAuth_OIT-GROVER-201", "YesAuth_OIT-CLIPPINGER-301"]
popups = ["BAKER-105_Popup", "GROVER-201_Popup", "CLIPPINGER-301_Popup"]

# === Prompt for Other Info ===
queue_name = prompt_with_list("Queue", "Select the Pharos queue name:", queues)
popup_name = prompt_with_list("Popup", "Select the popup name:", popups)

# === Get Available Drivers ===
resources_dir = "/Library/Printers/PPDs/Contents/Resources"
ppd_files = sorted([f for f in os.listdir(resources_dir) if f.endswith(".gz")])

def extract_manufacturer(filename):
    name = filename.replace(".PPD.gz", "").replace(".ppd.gz", "").replace(".gz", "")
    match = re.match(r"^([A-Z]{2,}(?=[^a-z]|$)|[A-Z][a-z]+|[A-Z][a-zA-Z]+(?=[A-Z]))", name)
    return match.group(1) if match else name.split()[0]

manufacturers = sorted(set(extract_manufacturer(f) for f in ppd_files))
manufacturer = prompt_with_list("Manufacturer", "Select the printer manufacturer:", manufacturers)

filtered_drivers = [f[:-3] for f in ppd_files if extract_manufacturer(f) == manufacturer]
selected_driver = prompt_with_list("Driver", "Select the printer model:", filtered_drivers)

# === Validate Required Inputs ===
if not all([technician_full, queue_name, popup_name, selected_driver]):
    print("❌ Operation cancelled or incomplete selection.")
    exit()

# === PPD Setup ===
driver_filename = selected_driver if selected_driver.lower().endswith(".ppd") else selected_driver + ".ppd"

# === Directory Setup ===
current_date = datetime.now().strftime("%m/%d/%Y")
volume_name = queue_name.replace("YesAuth_", "")
dmg_name = f"{volume_name}.dmg"
temp_dir = f"/tmp/{volume_name}"

if os.path.exists(temp_dir):
    shutil.rmtree(temp_dir)
custom_dir = os.path.join(temp_dir, "Custom")
os.makedirs(custom_dir, exist_ok=True)

# === Copy and Decompress PPD ===
source_ppd_gz = os.path.join(resources_dir, selected_driver + ".gz")
destination_ppd = os.path.join(custom_dir, driver_filename)
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

# === Create InstallFiles.txt ===
install_txt_path = os.path.join(custom_dir, "InstallFiles.txt")
with open(install_txt_path, "w") as f:
    f.write("# PPD file need to be copied\n")
    f.write("#\n")
    f.write(f"# {technician_full} -- {current_date}\n")
    f.write("#\n")
    f.write(f"/etc/cups/ppd/{driver_filename}\n")

# === Create PostInstall.sh ===
postinstall_path = os.path.join(custom_dir, "PostInstall.sh")
with open(postinstall_path, "w") as f:
    f.write("# Install print queue in CUPS\n")
    f.write("#\n")
    f.write(f"# {technician_full} -- {current_date}\n")
    f.write("#\n")
    f.write(f"lpadmin -p {popup_name} -v popup://PS1.ohio.edu/{queue_name} -E -P {driver_filename}\n")

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
