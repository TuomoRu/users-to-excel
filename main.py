import json
from pathlib import Path
from logging import log
from datetime import datetime
import requests
from openpyxl import Workbook

DATA_FOLDER = Path(__file__).parent / "data"
USERS_JSON_FILE = DATA_FOLDER / "users.json"
API_URL = "https://jsonplaceholder.typicode.com/users"


def fetch_user_data_to_json() -> None:
    """
    Fetches user data from the API and saves it into a JSON file. If the JSON file already exists,
    the API call is skipped.

    The file is saved into the data folder as users.json. If the file already exists,
    its content is not overwritten and the function does nothing.
    """
    if Path.exists(USERS_JSON_FILE):
        log(level=1, msg="users.json file already exists. Skipping API call.")
        return

    try:
        response = requests.get(API_URL, timeout=60)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        log(level=2, msg=f"Failed to fetch data from API: {e}")
        return

    with open(USERS_JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(response.json(), f)
        log(level=1, msg=f"users.json file created successfully: {USERS_JSON_FILE}.")


def write_user_data_to_excel(save_folder: Path) -> None:
    """
    Reads the users from the JSON file and writes them to an Excel file
    with the following columns: last name, first name, email, street, city,
    zipcode, phone, and website. The Excel file is named employees_YYYYMMDDHHMMSS.xlsx.

    :param save_folder: Path where the Excel file will be saved.
    """
    with open(USERS_JSON_FILE, "r", encoding="utf-8") as f:
        users = json.load(f)

    wb = Workbook()
    ws = wb.active

    ws.append(
        [
            "last name",
            "first name",
            "email",
            "street",
            "city",
            "zipcode",
            "phone",
            "website",
        ]
    )

    sorted_users = sorted(
        users, key=lambda user: (user["name"].split()[-1], user["name"].split()[0])
    )
    for user in sorted_users:
        ws.append(
            [
                user["name"].split()[-1],  # Last name
                user["name"].split()[0],  # First name
                user["email"],
                user["address"]["street"],
                user["address"]["city"],
                user["address"]["zipcode"],
                user["phone"],
                user["website"],
            ]
        )

    now = datetime.now()
    timestamp = now.strftime("%Y%m%d%H%M%S")
    save_path = Path(save_folder, f"employees_{timestamp}.xlsx")
    wb.save(save_path)
    log(level=1, msg=f"Excel file saved successfully to: {save_path}")


if __name__ == "__main__":
    fetch_user_data_to_json()
    write_user_data_to_excel("files")
