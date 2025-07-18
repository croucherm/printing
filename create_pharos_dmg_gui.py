import os
import re
import shutil
import subprocess
from datetime import datetime


# === AppleScript Prompt Functions ===

def prompt_with_list(title, prompt, options):
    """Prompt user with a selectable list via AppleScript."""
    list_items = "{" + ", ".join([f'\"{item}\"' for item in options]) + "}"
    script = f'set theList to {list_items}\nchoose from list theList with prompt "{prompt}" with title "{title}"'
    result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
    choice = result.stdout.strip().strip('"')
    return choice if choice.lower() != "false" else None

def prompt_with_list_or_custom(title, prompt, options, custom_label="Other..."):
    """Prompt user with a list, or allow custom text entry."""
    options_with_custom = options + [custom_label]
    choice = prompt_with_list(title, prompt, options_with_custom)
    if choice == custom_label:
        script = f'display dialog "Enter custom value:" with title "{title}" default answer ""'
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
        m = re.search(r'text returned:(.*)', result.stdout)
        return m.group(1).strip() if m else ""
    return choice

def prompt_yes_no(title, prompt):
    """Prompt user with a Yes/No dialog via AppleScript."""
    script = f'display dialog "{prompt}" with title "{title}" buttons {{"No", "Yes"}} default button "No"'
    result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
    return "Yes" in result.stdout


# === Helper Functions ===

def get_first_name(full_name):
    """Extract first name from 'Last, First' or 'First Last'."""
    if ',' in full_name:
        parts = full_name.split(',')
        return parts[1].strip()
    else:
        return full_name.split()[0]

def extract_bare_queue(file_name):
    """Remove planning unit and .dmg/.DMG from queue name."""
    name = re.sub(r'\.dmg$', '', file_name, flags=re.IGNORECASE)
    m = re.match(r"^YesAuth[-_]([^-_]+)[-_](.+)$", name)
    if m:
        return m.group(2)
    else:
        return name

def load_list_from_file(filepath, default=None):
    """Load a list from a text file, or return default."""
    if os.path.exists(filepath):
        with open(filepath) as f:
            return [line.strip() for line in f if line.strip()]
    return default or []

def get_technician_map(technicians):
    """Return display-to-filename mapping for technicians."""
    tech_map = {}
    for t in technicians:
        if ',' in t:
            last, first = map(str.strip, t.split(',', 1))
            display = f"{first} {last}"
            tech_map[display] = t
        else:
            tech_map[t] = t
    return tech_map

def extract_manufacturer(filename):
    """Guess manufacturer from PPD filename."""
    name = filename.replace(".PPD.gz", "").replace(".ppd.gz", "").replace(".gz", "")
    match = re.match(r"^([A-Z]{2,}(?=[^a-z]|$)|[A-Z][a-z]+|[A-Z][a-zA-Z]+(?=[A-Z]))", name)
    return match.group(1) if match else name.split()[0]

# === Main Workflow ===

def main():
    # Technician selection
    technicians_file = os.path.expanduser("~/Downloads/PharosDMG/technicians.txt")
    technicians = load_list_from_file(technicians_file, ["Croucher, Mike", "Tian, Zhiyong"])
    technician_map = get_technician_map(technicians)
    technician_display_list = list(technician_map.keys())
    technician_display = prompt_with_list_or_custom("Technician", "Select your name:", technician_display_list)
    technician_full = technician_map.get(technician_display, technician_display)

    if technician_full:
        first_name = get_first_name(technician_full)
        warning_message = (
            f"Hello {first_name}, please ensure the correct manufacturer and model printer drivers "
            f"are installed on this Mac before proceeding.\n\n"
            f"Also, ensure you have downloaded and saved the latest Pharos Popup and Notify Client package."
            f" to your ~/Downloads/PharosDMG folder. This file should be renamed to Installer.pkg."
        )
        warning_script = f'display dialog "{warning_message}" with title "Driver Installation Required" buttons {{"OK"}} default button "OK"'
        subprocess.run(['osascript', '-e', warning_script])

    # Package selection
    packages_file = os.path.expanduser("~/Downloads/PharosDMG/packages.txt")
    packages = load_list_from_file(packages_file, [])
    queue_name = prompt_with_list_or_custom("Queue", "Select the Pharos queue/package:", packages)
    if not queue_name:
        print("❌ No queue/package selected or entered.")
        return

    base_queue = re.sub(r'\.dmg$', '', queue_name, flags=re.IGNORECASE)
    bare_queue = extract_bare_queue(queue_name)
    popup_name = f"{bare_queue}_Popup"

    # Server selection
    servers_file = os.path.expanduser("~/Downloads/PharosDMG/servers.txt")
    servers = load_list_from_file(servers_file, ["PS1.ohio.edu", "PS2.ohio.edu", "PSB.ohio.edu"])
    server = prompt_with_list_or_custom("Server", "Select the print server:", servers)
    if not server:
        print("❌ No server selected or entered.")
        return

    # Manufacturer and driver selection
    resources_dir = "/Library/Printers/PPDs/Contents/Resources"
    try:
        ppd_files = sorted([f for f in os.listdir(resources_dir) if f.endswith(".gz")])
    except Exception as e:
        print(f"❌ Could not list printer drivers: {e}")
        return

    manufacturers = sorted(set(extract_manufacturer(f) for f in ppd_files))
    manufacturer = prompt_with_list_or_custom("Manufacturer", "Select the printer manufacturer:", manufacturers)
    if not manufacturer:
        print("❌ No manufacturer selected or entered.")
        return

    filtered_drivers = [f[:-3] for f in ppd_files if extract_manufacturer(f) == manufacturer]
    selected_driver = prompt_with_list_or_custom("Driver", "Select the printer model:", filtered_drivers)
    if not selected_driver:
        print("❌ No driver selected or entered.")
        return

    # Validate required inputs
    if not all([technician_full, queue_name, popup_name, selected_driver, server]):
        print("❌ Operation cancelled or incomplete selection.")
        return

    # Prepare filenames
    driver_filename = selected_driver if selected_driver.lower().endswith(".ppd") else selected_driver + ".ppd"
    # Remove spaces only (not underscores etc.) for sanitized driver name
    sanitized_driver_filename = driver_filename.replace(" ", "")

    # Create temp and output directories
    current_date = datetime.now().strftime("%m/%d/%Y")
    volume_name = base_queue
    temp_dir = f"/tmp/{volume_name}"

    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    custom_dir = os.path.join(temp_dir, "Custom")
    os.makedirs(custom_dir, exist_ok=True)

    # Copy and decompress PPD
    source_ppd_gz = os.path.join(resources_dir, selected_driver + ".gz")
    destination_ppd = os.path.join(custom_dir, sanitized_driver_filename)
    try:
        with open(destination_ppd, "wb") as out_f:
            subprocess.run(["gunzip", "-c", source_ppd_gz], stdout=out_f, check=True)
    except Exception as e:
        print(f"❌ Could not decompress PPD: {e}")
        return

    # Copy Installer.pkg
    installer_src = os.path.expanduser("~/Downloads/PharosDMG/Installer.pkg")
    installer_dst = os.path.join(temp_dir, "Installer.pkg")
    try:
        shutil.copy2(installer_src, installer_dst)
    except Exception as e:
        print(f"❌ Installer.pkg not found in ~/Downloads/PharosDMG: {e}")
        return

    # Create InstallFiles.txt
    install_txt_path = os.path.join(custom_dir, "InstallFiles.txt")
    with open(install_txt_path, "w") as f:
        f.write("# PPD file need to be copied\n#\n")
        f.write(f"# {technician_full} -- {current_date}\n#\n")
        f.write(f"/etc/cups/ppd/{sanitized_driver_filename}\n")

    # Create PostInstall.sh (use sanitized_driver_filename!)
    postinstall_path = os.path.join(custom_dir, "PostInstall.sh")
    with open(postinstall_path, "w") as f:
        f.write("# Install print queue in CUPS\n#\n")
        f.write(f"# {technician_full} -- {current_date}\n#\n")
        f.write(f"lpadmin -p {popup_name} -v popup://{server}/{base_queue} -E -P {sanitized_driver_filename}\n")

    # Create DMG
    output_dir = os.path.expanduser("~/Downloads/PharosDMG")
    os.makedirs(output_dir, exist_ok=True)
    dmg_path = os.path.join(output_dir, queue_name if queue_name.lower().endswith('.dmg') else queue_name + '.dmg')

    if os.path.exists(dmg_path):
        overwrite = prompt_yes_no("Overwrite DMG", f"The file {os.path.basename(dmg_path)} already exists. Overwrite?")
        if not overwrite:
            print("❌ Operation cancelled by user.")
            return

    try:
        subprocess.run([
            "hdiutil", "create", "-volname", volume_name,
            "-srcfolder", temp_dir,
            "-ov", "-format", "UDZO", dmg_path
        ], check=True)
        print(f"✅ DMG created at: {dmg_path}")
    except Exception as e:
        print(f"❌ Failed to create DMG: {e}")


if __name__ == "__main__":
    main()
