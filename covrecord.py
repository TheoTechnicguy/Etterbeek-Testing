# File: covrecord
# Author: Theo Technicguy covrecord-program@licolas.net
# Interpreter: Python 3.8
# Ext: py
# Licenced under GPU GLP v3. See LICENCE file for information.
# Copyright (c) TheoTechnicguy 2020
# -----------------------

# Notes
# Predict vial number.
# -----------------------
import os
import time
import logging
import shutil
import datetime
import json
from xml.etree import ElementTree as ET

import pyperclip
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import WebDriverException


class EmptySearchWarning(Warning):
    """Warning when generating an empty search."""

    def __init__(self):
        """Initialize warning class."""
        super(EmptySearchWarning, self).__init__(
            "Warning! You are generating an empty search url!"
        )


def maximize(driver):
    """Attempt to maximize window."""
    logging.info("Attemting maximization")
    try:
        driver.maximize_window()
    except WebDriverException:
        logging.warning("Could not maximize")


def minimize(driver):
    """Attempt to minimize window."""
    logging.info("Attemting minimization")
    try:
        driver.minimize_window()
    except WebDriverException:
        logging.warning("Could not minimize")


# Setup the log file configutation.
logging.basicConfig(
    filename=__file__ + ".log",
    level=logging.INFO,
    format="At %(asctime)s: %(name)s - %(levelname)s: %(message)s",
    filemode="w",
    datefmt="%d/%m/%Y %I:%M:%S %p",
    encoding="UTF-8",
)

logging.info("Started")
__version__ = "0.2.1"
__author__ = "Theo Technicguy"
logging.info("Version: %s by %s", __version__, __author__)

# Set driver location
GECKO_DRIVER = r"geckodriver.exe"
logging.info("Gecko driver located at: %s", GECKO_DRIVER)
logging.info("Using Firefox driver")

# CovRecord Form fields and buttons
FIELDS = {
    "name": '//*[@id="nom"]',
    "firstname": '//*[@id="prenom"]',
    "nationalnumber": '//*[@id="NISS"]',
    "dateofbirth": '//*[@id="ddn"]',
    "phone": '//*[@id="telephone"]',
    "email": '//*[@id="email"]',
    "test_tube": '//*[@id="numberEcouvillon"]',
    "doctor": '//*[@id="nomMedecin"]',
    "inami": '//*[@id="inamiMedecin"]',
    "gender": '//*[@id="sex"]',
    "zip": '//*[@id="adresse"]',
}
BUTTONS = {
    "print": "button.btn-primary:nth-child(1)",
    "save": "button.btn:nth-child(2)",
}

# INAMI Search data
INAMI_BASE_URL = (
    r"https://ondpanon.riziv.fgov.be/SilverPages/fr/Home/"
    r"SearchByForm?PageOffset=0&PageSize=200"
)
inami_search_data = {
    "lastname": "",
    "firstname": "",
    "nihdinumber": "",
    "where": "",
    "qualification": "",
}

# Set work directory
WORK_DIR = os.path.dirname(__file__)
logging.info("Work directory: %s", WORK_DIR)

# Create Error logs file.
if not os.path.exists(f"{WORK_DIR}\\errors"):
    os.mkdir(f"{WORK_DIR}\\errors")
# Set AutoHotkey script path.
if os.path.exists(os.path.join(WORK_DIR, "eid_viewer_export.exe")):
    AHK_PATH = os.path.join(WORK_DIR, "eid_viewer_export.exe")
elif os.path.exists(os.path.join(WORK_DIR, "eid_viewer_export.ahk")):
    AHK_PATH = os.path.join(WORK_DIR, "eid_viewer_export.ahk")
else:
    raise ImportError("Missing eid_viewer_export AHK script!")

# Set path for eid file
EID_DIR = os.path.join(os.path.expandvars("%TMP%"), "eid")
EID_PATH = os.path.join(EID_DIR, "patient.eid")

# Create temp path if it does not exist.
if not os.path.exists(EID_DIR):
    os.mkdir(EID_DIR)
else:
    # Clean directory by deleting everything.
    contents = os.listdir(EID_DIR)
    if contents:
        for file in contents:
            os.remove(os.path.join(EID_DIR, file))

# Setup Firefox drivers
logging.info("Setting up drivers")
drivers = {
    "covrecord": webdriver.Firefox(executable_path=GECKO_DRIVER),
    "mediris": webdriver.Firefox(executable_path=GECKO_DRIVER),
}

# For each driver, minimize window and implicitly wait 3 seconds.
for name, driver in drivers.items():
    minimize(driver)
    driver.implicitly_wait(3)
logging.info("Drivers setup")

# Open CovRecord page on covrecord driver.
drivers["covrecord"].get("http://croixrougewsl.be/covrecord/index.php")

# Let User login to CovRecord by maximizing the window.
maximize(drivers["covrecord"])
while True:
    try:
        drivers["covrecord"].find_element_by_xpath('//*[@id="username"]')
    except NoSuchElementException:
        logging.info("Logged in to CovRecord")
        break
    else:
        time.sleep(1)
minimize(drivers["covrecord"])

# Open Mediris page
logging.info("Getting Mediris")
drivers["mediris"].get("https://bxltestest.mediris.be/Wachtzaal")

# Login to Mediris page
drivers["mediris"].find_element_by_xpath('//*[@id="username"]').send_keys(
    os.environ["MEDIRIS_USER"]
)
drivers["mediris"].find_element_by_xpath('//*[@id="password"]').send_keys(
    os.environ["MEDIRIS_PASSWORD"], Keys.RETURN
)

# Test tube prediction variable.
test_tube_predict = ""
# ---------- END Setup ----------

# ---------- START Auto-update ----------
GITHUB_URL = (
    "https://api.github.com/repos/TheoTechnicguy/" "Etterbeek-Testing/releases"
)

# Github needs a custom header.
# Authentication is made via a token from github.
# For security, it is stored as an local_user environment variable.
header = {
    "Authenication": "token " + os.environ["GITHUB_CRDEV_TOKEN"],
    "accept": "application/vnd.github.v3+json",
}

logging.info("Getting github repos")

# Get github repos w/ requests "GET" method.
github_page = requests.get(GITHUB_URL, headers=header)

# Check if all is ok (status_code 200)
if github_page.status_code != 200:
    # Notify of fail and write page contents into the log file.
    print(f"Could not autoupdate. You are running version {__version__}.")
    logging.critical("Something went wrong...")
    logging.info(github_page.status_code)
    logging.info(github_page.text)
else:
    # loop (once) throug releases json.
    # The `for` loop is to deal with no releases.
    for release in json.loads(github_page.text):
        if release["tag_name"] > __version__:
            logging.info("Attepting update")
            # Loop throug the assets in the release.
            for asset in release["assets"]:
                logging.info("Getting file: %s", asset["name"])
                # Only get the release if does not exist.
                if not os.path.exists(os.path.join(WORK_DIR, asset["name"])):
                    with open(asset["name"], "wb+") as file:
                        file.write(
                            requests.get(asset["browser_download_url"].content)
                        )
                else:
                    logging.info("File already exists.")
        else:
            logging.info("Already up to date.")

        # Only check 1st release because they are (or should be) incremental.
        break
# ---------- END Auto-update ---------

try:
    while True:
        # ---------- START eID Fetching -----------
        logging.info("---------- Next Patient ----------")
        print("\n\n---------- Next patient ----------")
        # Wait for card to be read.
        card = input("Read card")
        # Exit if asked to quit.
        if card.lower() in ("q", "quit", "e", "exit"):
            raise SystemExit()

        # Export file via executing AHK script
        logging.info("Executing AHK Script at %s", AHK_PATH)
        os.system(AHK_PATH)
        time.sleep(1)

        # Wait for the file to exist.
        while not os.path.exists(EID_PATH):
            logging.info("eID path still not existing")
            time.sleep(1)

        # Parse eID file to XML and get it's root.
        logging.info("eID XML at %s %s", EID_PATH, os.path.exists(EID_PATH))
        xml_file = ET.parse(EID_PATH)
        xml_root = xml_file.getroot()

        # Start full_id dictionnary with identity attributes
        full_id = xml_root.find("identity").attrib

        # Get all Id items (name, nationality)
        for item in xml_root.findall("identity/*"):
            full_id[item.tag] = item.text

        # Delete photo; large and not nessessary.
        # COMBAK: Delete or not?
        del full_id["photo"]

        # Get address
        for item in xml_root.findall("address/*"):
            full_id[item.tag] = item.text

        # Print firs and last name and address.
        print("First name\t", full_id["firstname"])
        print("Last name\t", full_id["name"])
        print(
            "Address\t",
            full_id["streetandnumber"],
            full_id["zip"],
            full_id["municipality"],
        )

        # Cleanup temp eID file.
        os.remove(EID_PATH)
        # ---------- END eID Fetching ----------

        # ---------- START phone and email fetching ----------
        # Let user select patient
        logging.info("Swiching to Mediris")
        maximize(drivers["mediris"])
        get_doctor_info = False

        while True:
            # Wait for user to select the patient
            # Checked by looking for the patent tab.
            logging.info("Waiting for patient select")
            while True:
                try:
                    drivers["mediris"].find_element_by_xpath(
                        '//*[@id="patientCrumb"]'
                    ).click()
                except NoSuchElementException:
                    pass
                else:
                    minimize(drivers["mediris"])
                    break

            # COMBAK: Can fetch w/o user interaction?
            # Make input fields accessible by keyboard (allow editing).
            logging.info("Attempting edit mode.")
            try:
                drivers["mediris"].find_element_by_xpath(
                    "/html/body/div[2]/div[2]/div[3]/div[3]/div[1]/div[1]/div/a"
                ).click()
            except NoSuchElementException:
                # If it fails, try backup button.
                logging.info("Failed edit mode.")
                try:
                    drivers["mediris"].find_element_by_xpath(
                        '//*[@id="inputRijksregisternummer"]'
                    ).send_keys()
                except ElementNotInteractableException:
                    # If backup button fails, ask to enter information manually
                    logging.warning("Switching to manual entry.")
                    check = True
                    while check:
                        # Start by verifying natianl registry number.
                        if str(full_id["nationalnumber"]) != input(
                            "National Number:\t"
                        ):
                            logging.warning("National Numbers do not match!")
                            print("The national numbers do not match!")
                            continue

                        # Ask phone and email.
                        full_id["phone"] = input("Phone Number:\t")
                        full_id["email"] = input("Email Address:\t")

                        # Ask confirmation
                        while check:
                            check_in = input("Is this correct? [yes/no]: ")
                            if check_in.lower().startswith("y"):
                                check = False
                            elif check_in.lower().startswith("n"):
                                break
                            else:
                                print("Input not recognized, use: `Yes`/`No`.")
                else:
                    get_doctor_info = True
            else:
                get_doctor_info = True

            if get_doctor_info:
                # Copy registry number to clipboard.
                drivers["mediris"].find_element_by_xpath(
                    '//*[@id="inputRijksregisternummer"]'
                ).send_keys(Keys.CONTROL, "a", "c")

                # Verify register number fom clipboard.
                if full_id["nationalnumber"] != pyperclip.paste():
                    print(
                        "The national numbers do not match!",
                        "Did you select the correct patient?",
                    )
                    time.sleep(3)
                    maximize(drivers["mediris"])
                    input("Select patient")
                    minimize(drivers["mediris"])
                    continue

                # Copy phone to clipboard.
                drivers["mediris"].find_element_by_xpath(
                    '//*[@id="inputTelefoonnummer"]'
                ).send_keys(Keys.CONTROL, "a", "c")
                # Fetch phone fom clipboard. If it is the registry number
                # no phone is entered, so set to "".
                full_id["phone"] = pyperclip.paste()
                if full_id["phone"] == full_id["nationalnumber"]:
                    logging.warning("No phone selected!")
                    full_id["phone"] = ""

                # Copy email to clipboard.
                drivers["mediris"].find_element_by_xpath(
                    '//*[@id="inputEmail"]'
                ).send_keys(Keys.CONTROL, "a", "c")
                full_id["email"] = pyperclip.paste()
                # Fetch email form clipboard. If it is the phone
                # or the registry number, no email entered, so set to "".
                if full_id["email"] in (
                    full_id["phone"],
                    full_id["nationalnumber"],
                ):
                    full_id["email"] = ""
                    logging.warning("No email address selected!")

                # Get missing info.
                if not full_id["phone"]:
                    full_id["phone"] = input("Phone number: ")
                if not full_id["email"]:
                    full_id["email"] = input("Email address: ")

                # break free of the loop.
                break

        logging.info("After info fetching, full_id: %s", full_id)
        # ---------- END phone and email fetching ----------

        # ---------- START Doctor Fetching ----------
        # Go to Doctor section
        logging.info("Fetching Doctor")
        try:
            drivers["mediris"].find_element_by_xpath(
                '//*[@id="huisartsCrumb"]'
            ).click()
        except ElementClickInterceptedException:
            # If it fails, use backup button.
            logging.warning("Switching to backup button")
            try:
                drivers["mediris"].find_element_by_xpath(
                    "/html/body/div[2]/div[2]/div[3]/div[1]/a[2]/span[2]"
                ).click()
            except ElementClickInterceptedException:
                # If the backup button fails, ask to select it.
                maximize(drivers["mediris"])
                input("Select Doctor tab")
                minimize(drivers["mediris"])

        # Get selected doctor text
        for attempt in range(2):
            try:
                # Try to get the doctor name text.
                full_id["doctor"] = (
                    drivers["mediris"]
                    .find_element_by_xpath(
                        (
                            "/html/body/div[2]/div[2]/div[3]/div[3]/div[5]"
                            "/div[1]/div[1]/span[1]"
                        )
                    )
                    .text
                )
            except NoSuchElementException:
                #  If it fails, try the backup location.
                try:
                    full_id["doctor"] = (
                        drivers["mediris"]
                        .find_element_by_xpath(
                            (
                                "/html/body/div[2]/div[2]/div[3]/div[3]/div[5]"
                                "/div[2]/div[1]/span[1]"
                            )
                        )
                        .text
                    )
                except NoSuchElementException:
                    # Ask to check/confirm that no doctor is selected.
                    # Cannot select an specialized doctor.
                    logging.warning(
                        "No Doctor selected. %s-ing.",
                        ("check", "confirm")[attempt],
                    )

                    input(
                        "No Doctor selected. Please %s."
                        % (("check", "confirm")[attempt])
                    )
                    full_id["doctor"] = ""
            else:
                break
        # ---------- END Doctor Fetching ----------

        # ---------- START Doctor and nihdi number fetching ----------
        if full_id["doctor"]:
            # --- START Decompose doctor's name. ---
            # Conform user input.
            full_name = full_id["doctor"]
            full_name = (
                full_name.lower().replace("dr.", "").replace(" -", "").strip()
            )
            logging.info("After confoming, user input is %s", full_name)

            # Convert to list
            name_list = full_name.split()
            logging.info("Splitting name into %s", name_list)

            # Decompostion for 3 names
            # [0] = last name
            # [1] = first name
            # [2] = middle name

            # Tuple of "train words" in family names.
            unions = (
                "van",
                "den",
                "vanden",
                "vande",
                "de",
                "du",
                "la",
                "le",
                "dela",
                "de la",
            )

            name_list_edited = []
            name_combinig = ""
            jump = False
            # Iterate the names.
            for name in name_list:
                # Ouput (too much) debugging data.
                logging.info("Current name %s", name)
                logging.info("Is union: %s", name in unions)
                logging.info(
                    "Edited names list: %s",
                    name_list_edited,
                )
                logging.info("We are%s juming." % ("" if jump else "n't"))

                # If the current name is a "train word",
                # concatenate it with previous.
                if name in unions:
                    name_combinig += " " + name
                    logging.info("Combined name: %s", name_combinig)
                    jump = True
                elif jump:
                    # Finish combining names and append to names list.
                    name_combinig += " " + name
                    name_combinig = name_combinig.strip()
                    name_list_edited.append(name_combinig)
                    jump = False
                    continue
                else:
                    # Just append the name. #You'reNotSpecial
                    name_list_edited.append(name)

            # Try to access middle name, else create empty middle name.
            try:
                name_list_edited[2]
            except IndexError:
                logging.info("No middle name")
                name_list_edited.append("")

            # Return names dictionnary.
            doc_search = {
                "firstname": name_list_edited[1],
                "lastname": name_list_edited[0],
                "middlename": name_list_edited[2],
            }
            # --- END Decompose Doctor's name ---

            # Search NIHDI number. Pass middlename key.
            for attempt in range(3):
                logging.info("Searching for doctor")
                try:
                    for key, value in doc_search.items():
                        # Conform user input.
                        key = key.strip().lower().replace("_", "")
                        value = value.strip().lower()

                        # Check that int-requiering values are ints.
                        if (
                            key in ("where", "nihdinumber", "qualification")
                            and not value.isdigit()
                        ):
                            raise ValueError(f"{key} expects a number.")

                        # Add the data to the search data.
                        inami_search_data[key] = value
                except KeyError as e:
                    if key == "middlename":
                        pass
                    else:
                        raise KeyError(e)

                # --- START Doctor search ---
                # copy INAMI_BASE_URL - Need the other as template.
                search_url = INAMI_BASE_URL

                # Add the search data into the URL.
                for key, value in inami_search_data.items():
                    if (
                        value and key != "middlename"
                    ):  # Cannot add middle names.
                        search_url += "&" + key + "=" + value

                # Warn if the search is empty (270000+ results :P)
                if search_url == INAMI_BASE_URL:
                    print(EmptySearchWarning())

                # Get the page and parse it to BeautifulSoup.
                page = requests.get(search_url)
                soup = BeautifulSoup(page.text, "html.parser")

                # Iterate all medical staff (devided into div-s class col-sm-4)
                medical_staff_list = []
                for medical_staff in soup.find_all(
                    "div", {"class": "col-sm-4"}
                ):
                    medical_staff_info = {}
                    # Get the full name and conform it
                    full_name = medical_staff.find(
                        "small", {"class": "ng-binding"}
                    ).string
                    full_name = full_name.strip().lower()

                    # Split the name and set the first and last name.
                    names = full_name.split(", ")
                    medical_staff_info["firstname"] = names[1]
                    medical_staff_info["lastname"] = names[0]

                    # Get remaining info (INAMI, Address...)
                    information = medical_staff.find(
                        "div", {"class": "panel-body"}
                    )
                    for row in information.find_all("div", {"class": "row"}):
                        # Skip if label is empty/does not exist
                        if row.label is None:
                            continue

                        # Get label text and try to conform.
                        # If cannot conform, skip.
                        label = row.label.small.string
                        try:
                            label = label.strip().lower()
                        except AttributeError:
                            continue

                        # Get value text and try to conform.
                        # If cannot conform, skip.
                        value = row.div.p.small.string
                        try:
                            value = value.strip().lower()
                        except AttributeError:
                            continue

                        # Switch setting correct attributes
                        if "inami" in label:
                            medical_staff_info["inami"] = value

                        elif "date de qualif" in label:
                            # Try to convert label into a datetime.date
                            try:
                                date_components = value.split("/")
                            except AttributeError:
                                # If it fails, set it to "Unknown".
                                medical_staff_info[label] = "Unknown"
                            else:
                                medical_staff_info[label] = datetime.date(
                                    year=int(date_components[2]),
                                    month=int(date_components[1]),
                                    day=int(date_components[0]),
                                )

                        elif "adresse" in label:
                            # Get address: buiding, street and number.
                            address1 = row.div.p.find_next("small")
                            address2 = address1.find_next("small")

                            # Try to convert and conform address1.
                            medical_staff_info[label] = address1.string
                            try:
                                medical_staff_info[label] = (
                                    medical_staff_info[label].strip().lower()
                                )
                            except AttributeError:
                                # If it fails, not an address or empty.
                                continue

                            # Try to get, convert and conform address2.
                            try:
                                address2 = address2.get_text().strip().lower()
                            except AttributeError:
                                # If it failes, address not available.
                                pass
                            else:
                                # Finish conforming and contcatenate
                                address2 = address2.replace("\xa0", "")
                                address2 = address2.replace("\n\n", "\n")
                                address2 = address2.replace("  ", "")
                                medical_staff_info[label] += " " + address2

                        elif "qualification" in label:
                            # Get qualifications (code + description)
                            qualification_code = row.div.p.find_next("small")
                            qualification_description = (
                                qualification_code.find_next("small")
                            )

                            # Get texts, conform and convert.
                            qualification_code = int(
                                qualification_code.string.strip().lower()
                            )

                            qualification_description = (
                                qualification_description.string
                            )
                            try:
                                qualification_description = (
                                    qualification_description.strip()
                                )
                                qualification_description = (
                                    qualification_description.lower()
                                )
                            except AttributeError:
                                # Don't worry if it doesn't work.
                                pass

                            # Set the attributes.
                            medical_staff_info[
                                "qualification_code"
                            ] = qualification_code
                            medical_staff_info[
                                "qualification_description"
                            ] = qualification_description

                        else:
                            # Collect the rest and set the attribute.
                            # #You'reNotSpecial
                            medical_staff_info[label] = value

                    # Append the data to the ouput list.
                    medical_staff_list.append(medical_staff_info)

                doc_resuts = medical_staff_list
                # --- END Doctor information Search ---

                # Check for qualification
                # NOTE: Should not be a problem.
                # COMBAK: Remove?
                doc_keeper = []
                for doc in doc_resuts:
                    # Only keep doctors with GP status (General Practitioner)
                    if doc["qualification_code"] in {
                        0,
                        1,
                        3,
                        4,
                        5,
                        6,
                        7,
                        8,
                        9,
                    }:
                        logging.info("Doctor added: %s", doc)
                        doc_keeper.append(doc)

                # Check how many doctors we have left.
                # If there is only 1, we autofound our doctor
                if len(doc_keeper) == 1:
                    doc_out = doc_keeper[0]
                    logging.info(
                        "We are happy! Doctor %s %s",
                        doc_out["firstname"],
                        doc_out["lastname"],
                    )

                else:
                    # Let user check name and overwrite.
                    logging.info(
                        "Need user help, have %s items", len(doc_keeper)
                    )
                    doc_search_auto = doc_search.copy()
                    check = True
                    print_out = True
                    while check:
                        # Print names only if needed.
                        if print_out:
                            logging.info("Printing")
                            for key, value in doc_search.items():
                                print(key, value.title(), sep="\t")
                            print_out = False

                        # Let user check names.
                        check_in = input("Is this correct? [yes/no] ")
                        if check_in.lower().strip().startswith("y"):
                            check = False
                        elif check_in.lower().strip().startswith("n"):
                            # Else let user correct.
                            for key in doc_search.keys():
                                doc_search[key] = input(
                                    f"Enter Doctor's {key}: "
                                )
                            print_out = True
                        else:
                            print("Input not recognized, use: `Yes`/`No`.")

                    # Not found, ask for INAMI.
                    if doc_search == doc_search_auto:
                        doc_out = doc_search.copy()
                        doc_out["inami"] = input("INAMI: ")
                    else:
                        # Otherwise search again.
                        continue

                # Beautify doctor and add inami
                logging.info("INAMI and name beautification")
                full_id["doctor"] = (
                    "Dr. " + doc_out["firstname"] + " " + doc_out["lastname"]
                ).title()
                full_id["inami"] = doc_out["inami"]
                break
        else:
            # Worst case, set doctor and INAMI empty.
            full_id["doctor"] = ""
            full_id["inami"] = ""
            logging.info("Skipping Search, no doctor selected")
        # --------- END Doctor nihdi number fetching ----------

        # ---------- START Test Tube ID ----------
        # Get test tube ID
        attempt = 0
        while True:
            logging.info("Predicting test tube ID: %s", test_tube_predict)
            full_id["test_tube"] = input(
                f"Test tube code ({test_tube_predict}) :"
            )
            # If input empty, use predicted test tube.
            if not full_id["test_tube"] and test_tube_predict:
                full_id["test_tube"] = test_tube_predict
            logging.info("Got test tube %s", full_id["test_tube"])

            # Correctly format test tubes.
            if (char := full_id["test_tube"][3]) != "-":
                full_id["test_tube"].replace(char, "-")

            # Assert test tube starts with CD and ends in M.
            if not (
                full_id["test_tube"].startswith("C19")
                and full_id["test_tube"].endswith("M")
            ):
                print("This is not a valid code...")
                attempt += 1
            else:
                # Set next test tube ID prediction into memory.
                test_tube_decompse = full_id["test_tube"].split("-")
                assert test_tube_decompse[1].isdigit()
                test_tube_decompse[1] += 1
                test_tube_predict = "".join(test_tube_decompse)
                break

            if attempt > 1:
                # Let user overwrite Not asserted ID.
                if input("Overwrite? [yes/no]").lower().startswith("y"):
                    logging.warning("User Overwrote program.")
                    break
        # ---------- END Test Tube ID ----------

        # ---------- START Form fillout ----------
        # write all values to CovRecord form.
        logging.info("Wirting out")
        for element, field in FIELDS.items():
            # Find element, clear feld and write value.
            cur_field = drivers["covrecord"].find_element_by_xpath(field)
            cur_field.clear()
            try:
                cur_field.send_keys(full_id[element])
            except KeyError:
                # If no value is set, leave blank.
                cur_field.send_keys("")

        # Keep a copy of test tube ID in the clipboard
        logging.info("Clipping test tube to clipboard")
        pyperclip.copy(full_id["test_tube"])

        # Maximize window for user interaction.
        maximize(drivers["covrecord"])

        # Send to printer.
        logging.info("Sending print")
        drivers["covrecord"].find_element_by_css_selector(
            BUTTONS["print"]
        ).click()

        # Select Corona form on Mediris
        logging.info("Selecting Corona form")
        try:
            drivers["mediris"].find_element_by_xpath(
                '//*[@id="anderebehandelingCrumb"]'
            ).click()
        except ElementClickInterceptedException:
            # Let user finalize Mediris form.
            # NOTE: Not maximizing because user busy with CovRecord from.
            input("Select other treatement tab")

        # Add other treatement.
        drivers["mediris"].find_element_by_xpath(
            (
                "/html/body/div[2]/div[2]/div[3]/div[3]/div[12]/div[2]/table/"
                "tbody/tr/td[4]/a"
            )
        ).click()

        # ---------- Cleanup ----------
        logging.info("Cleaning up")
        try:
            del doc_out
        except NameError:
            pass

except KeyboardInterrupt:
    print("Quitting")
    logging.info("Quitting")
    for key, driver in drivers.items():
        driver.close()

except SystemExit:
    print("Quitting")
    logging.info("Quitting")
    for key, driver in drivers.items():
        driver.close()

except Exception as e:
    logging.critical(e)
    now_string = (
        str(datetime.datetime.now()).replace(" ", "_").replace(":", "-")
    )
    shutil.copyfile(
        f"{__file__}.log",
        f"{WORK_DIR}\\errors\\{now_string}-ERROR.log",
    )
    raise
    input()
