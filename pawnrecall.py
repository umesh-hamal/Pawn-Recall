import os
import shutil
import sqlite3
from datetime import datetime, timedelta
import getpass
import argparse
import subprocess

VERSION = "0.1"

BLUE = "\033[94m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
RED = "\033[91m"
NO_COLOR = "\033[0m"

def display_banner():
    banner = (
    r"""        ooooooooo.         .o.       oooooo   oooooo     oooo ooooo      ooo      ooooooooo.   oooooooooooo   .oooooo.         .o.       ooooo        ooooo        
       `888   `Y88.      .888.       `888.    `888.     .8'  `888b.     `8'      `888   `Y88. `888'     `8  d8P'  `Y8b       .888.      `888'        `888'        
        888   .d88'     .8"888.       `888.   .8888.   .8'    8 `88b.    8        888   .d88'  888         888              .8"888.      888          888         
        888ooo88P'     .8' `888.       `888  .8'`888. .8'     8   `88b.  8        888ooo88P'   888oooo8    888             .8' `888.     888          888         
        888           .88ooo8888.       `888.8'  `888.8'      8     `88b.8        888`88b.     888    "    888            .88ooo8888.    888          888         
        888          .8'     `888.       `888'    `888'       8       `888        888  `88b.   888       o `88b    ooo   .8'     `888.   888       o  888       o 
        o888o        o88o     o8888o       `8'      `8'       o8o        `8       o888o  o888o o888ooooood8  `Y8bood8P'  o88o     o8888o o888ooooood8 o888ooooood8

v"""
        + VERSION
        + """ / Red Haired!/Github:umesh-hamal 
"""
    )
    print(RED + banner + NO_COLOR)


def modify_permissions(path):
    try:
        subprocess.run(
            ["icacls", path, "/grant", f"{getpass.getuser()}:(OI)(CI)F", "/T", "/C", "/Q"],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        print(f"{GREEN} Permissions modified for {path} and all its subdirectories and files{ENDC}")
    except subprocess.CalledProcessError as e:
        print(f"{RED} Failed to modify permissions for {path}: {e}{ENDC}")


def main(from_date=None, to_date=None, search_term=None):
    display_banner()
    username = getpass.getuser()
    default_path = f"C:\\Users\\{username}\\AppData\\Local\\CoreAIPlatform.00\\UKP"

    if not os.path.exists(default_path):
        print("Base path does not exist.")
        return

    modify_permissions(default_path)
    guid_folder = next((os.path.join(default_path, folder_name) for folder_name in os.listdir(default_path) if os.path.isdir(os.path.join(default_path, folder_name))), None)

    if not guid_folder:
        print("Could not find the GUID folder.")
        return

    print(f"Recall folder found: {guid_folder}")

    db_path = os.path.join(guid_folder, "ukg.db")
    image_store_path = os.path.join(guid_folder, "ImageStore")

    if not (os.path.exists(db_path) and os.path.exists(image_store_path)):
        print("🚫 Windows Recall feature not found. Nothing to extract.")
        return

    proceed = input("Windows Recall feature found. Do you want to proceed with the extraction? (y/n): ").strip().lower()
    if proceed != "yes":
        print("Extraction aborted.")
        return

    timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M")
    extraction_folder = os.path.join(os.getcwd(), f"{timestamp}_Recall_Extraction")
    os.makedirs(extraction_folder, exist_ok=True)
    print(f"Creating extraction folder: {extraction_folder}\n")

    shutil.copy(db_path, extraction_folder)
    shutil.copytree(image_store_path, os.path.join(extraction_folder, "ImageStore"), dirs_exist_ok=True)

    for image_file in os.listdir(os.path.join(extraction_folder, "ImageStore")):
        image_path = os.path.join(extraction_folder, "ImageStore", image_file)
        new_image_path = f"{image_path}.jpg"
        if not image_path.endswith(".jpg"):
            os.rename(image_path, new_image_path)

    db_extraction_path = os.path.join(extraction_folder, "ukg.db")
    conn = sqlite3.connect(db_extraction_path)
    cursor = conn.cursor()

    from_date_timestamp = int(datetime.strptime(from_date, "%Y-%m-%d").timestamp()) * 1000 if from_date else None
    to_date_timestamp = int((datetime.strptime(to_date, "%Y-%m-%d") + timedelta(days=1)).timestamp()) * 1000 if to_date else None

    query = "SELECT WindowTitle, TimeStamp, ImageToken FROM WindowCapture WHERE (WindowTitle IS NOT NULL OR ImageToken IS NOT NULL)"
    cursor.execute(query)
    rows = cursor.fetchall()

    captured_windows = []
    images_taken = []
    for window_title, timestamp, image_token in rows:
        if (from_date_timestamp is None or from_date_timestamp <= timestamp) and (to_date_timestamp is None or timestamp < to_date_timestamp):
            readable_timestamp = datetime.fromtimestamp(timestamp / 1000).strftime("%Y-%m-%d %H:%M:%S")
            if window_title:
                captured_windows.append(f"[{readable_timestamp}] {window_title}")
            if image_token:
                images_taken.append(f"[{readable_timestamp}] {image_token}")

    captured_windows_count = len(captured_windows)
    images_taken_count = len(images_taken)
    output = [
        f"Captured Windows: {captured_windows_count}",
        f"Images Taken: {images_taken_count}"
    ]

    
    if search_term:
        search_query = f"SELECT c1, c2 FROM WindowCaptureTextIndex_content WHERE c1 LIKE ? OR c2 LIKE ?"
        cursor.execute(search_query, (f"%{search_term}%", f"%{search_term}%"))
        search_results = cursor.fetchall()
        search_results_count = len(search_results)
        output.append(f"Search results for '{search_term}': {search_results_count}")

        search_output = [f"c1: {result[0]}, c2: {result[1]}" for result in search_results]
    else:
        search_output = []

    with open(os.path.join(extraction_folder, "TotalRecall.txt"), "w", encoding="utf-8") as file:
        file.write("Captured Windows:\n")
        file.write("\n".join(captured_windows))
        file.write("\n\nImages Taken:\n")
        file.write("\n".join(images_taken))
        if search_term:
            file.write("\n\nSearch Results:\n")
            file.write("\n".join(search_output))

    conn.close()

    for line in output:
        print(line)

    print(f"\nSummary of the extraction is available in the file:")
    print(f"{YELLOW}{os.path.join(extraction_folder, 'TotalRecall.txt')}{ENDC}")
    print(f"\nFull extraction folder path:")
    print(f"{YELLOW}{extraction_folder}{ENDC}")






if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract and display Windows Recall data.")
    parser.add_argument("--from_date", help="The start date in YYYY-MM-DD format.", type=str, default=None)
    parser.add_argument("--to_date", help="The end date in YYYY-MM-DD format.", type=str, default=None)
    parser.add_argument("--search", help="Search term for text recognition data.", type=str, default=None)
    args = parser.parse_args()

    try:
        if args.from_date:
            datetime.strptime(args.from_date, "%Y-%m-%d")
        if args.to_date:
            datetime.strptime(args.to_date, "%Y-%m-%d")
    except ValueError:
        parser.error("Date format must be YYYY-MM-DD.")

    main(args.from_date, args.to_date, args.search)