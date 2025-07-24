import json
import subprocess
import sys
import os
import platform
import shutil
import urllib.request
import zipfile
import tarfile

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

def resolve_adjective(adj_node, depth=0, max_depth=20):
    if depth > max_depth:
        return "[RECURSION_LIMIT]"

    if not isinstance(adj_node, dict):
        return ""

    key_list = adj_node.get("key", [])
    if not key_list or not isinstance(key_list, list):
        return ""

    key = key_list[0]

    if "variables" not in adj_node:
        return key

    variables = adj_node["variables"]
    if not variables or not isinstance(variables[0], list):
        return key

    for var in variables[0]:
        var_value = var.get("value", [{}])[0]
        resolved = resolve_adjective(var_value, depth + 1)
        if resolved:
            return resolved

    return key

def get_power(entry, key):
    val = entry.get(key)
    if isinstance(val, list) and len(val) > 0:
        return val[0]
    return 0.0

def download_and_extract(url, dest_dir, is_zip=True):
    archive_name = os.path.join(dest_dir, url.split("/")[-1])
    print(f"Downloading {url} ...")
    urllib.request.urlretrieve(url, archive_name)
    print(f"Downloaded archive to {archive_name}")

    if is_zip:
        with zipfile.ZipFile(archive_name, 'r') as zip_ref:
            zip_ref.extractall(dest_dir)
    else:
        with tarfile.open(archive_name, 'r:gz') as tar_ref:
            tar_ref.extractall(dest_dir)

    print(f"Extracted archive to {dest_dir}")
    os.remove(archive_name)
    print(f"Deleted archive {archive_name}")

def ensure_sav2json():
    # Directory where this script lives
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Determine platform and corresponding download info
    system = platform.system()
    if system == "Windows":
        exe_name = "sav2json.exe"
        url = "https://github.com/ErikKalkoken/stellaris-tool/releases/download/v0.1.1/sav2json-v0.1.1-windows-amd64.zip"
        is_zip = True
    elif system == "Linux":
        exe_name = "sav2json"
        url = "https://github.com/ErikKalkoken/stellaris-tool/releases/download/v0.1.1/sav2json-v0.1.1-linux-amd64.tar.gz"
        is_zip = False
    elif system == "Darwin":
        exe_name = "sav2json"
        url = "https://github.com/ErikKalkoken/stellaris-tool/releases/download/v0.1.1/sav2json-v0.1.1-darwin-amd64.zip"
        is_zip = True
    else:
        print(f"Unsupported OS: {system}")
        sys.exit(1)

    exe_path = os.path.join(script_dir, exe_name)

    if os.path.isfile(exe_path):
        print(f"Found existing '{exe_name}' in script directory.")
        return exe_path

    print(f"'{exe_name}' not found. Downloading from {url} ...")
    download_and_extract(url, script_dir, is_zip)

    if not os.path.isfile(exe_path):
        print(f"Error: '{exe_name}' not found after extraction.")
        sys.exit(1)

    # On Unix, make sure executable bit is set
    if system != "Windows":
        os.chmod(exe_path, 0o755)

    print(f"'{exe_name}' is ready to use.")
    return exe_path

def main():
    if len(sys.argv) != 2:
        print("Usage: python script.py <savefile.sav>")
        sys.exit(1)

    save_file = sys.argv[1]

    # Get the filename without path, replace spaces with underscores, change extension to .xlsx
    base_name = os.path.basename(save_file)
    base, ext = os.path.splitext(base_name)
    safe_name = base.replace(" ", "_") + ".xlsx"

    # Save in current working directory
    excel_filename = os.path.join(os.getcwd(), safe_name)

    sav2json_path = ensure_sav2json()

    print(f"Running '{sav2json_path} {save_file}' ...")
    try:
        subprocess.run([sav2json_path, save_file], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running sav2json: {e}")
        sys.exit(1)

    json_file = "gamestate.json"
    if not os.path.exists(json_file):
        print(f"Expected {json_file} not found after sav2json")
        sys.exit(1)

    with open(json_file, 'r') as f:
        data = json.load(f)

    country_data = data.get('country', [{}])[0]
    results = []

    for country_id, entry_list in country_data.items():
        if (
            isinstance(entry_list, list)
            and entry_list
            and isinstance(entry_list[0], dict)
            and entry_list[0].get("type") == ["default"]
        ):
            entry = entry_list[0]

            adjective_str = "UNKNOWN"
            adjective_data = entry.get("adjective", [])
            if adjective_data and isinstance(adjective_data[0], dict):
                adjective_str = resolve_adjective(adjective_data[0])

            military_power = get_power(entry, "military_power")
            tech_power = get_power(entry, "tech_power")
            economy_power = get_power(entry, "economy_power")

            results.append({
                "id": country_id,
                "adjective": adjective_str,
                "military_power": military_power,
                "tech_power": tech_power,
                "economy_power": economy_power,
            })

    wb = Workbook()
    ws = wb.active
    ws.title = "Empires"

    headers = ["ID", "Empire", "Military Power", "Tech Power", "Economy Power"]
    ws.append(headers)

    for r in results:
        # Try to convert id to int for proper numeric sorting
        try:
            id_val = int(r["id"])
        except ValueError:
            id_val = r["id"]

        ws.append([id_val, r["adjective"], r["military_power"], r["tech_power"], r["economy_power"]])

    num_rows = len(results) + 1  # +1 for header
    table_ref = f"A1:E{num_rows}"

    table = Table(displayName="EmpiresTable", ref=table_ref)

    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    table.tableStyleInfo = style

    ws.add_table(table)

    wb.save(excel_filename)
    print(f"Excel file '{excel_filename}' with sortable table created successfully.")

    # Cleanup
    for filename in ["gamestate.json", "meta.json"]:
        if os.path.exists(filename):
            os.remove(filename)
            print(f"Deleted {filename}")

if __name__ == "__main__":
    main()
