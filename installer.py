# File: installer
# Author: Theo Technicguy installer-program@licolas.net
# Interpreter: Python 3.9
# Ext: py
# Licenced under GPU GLP v3. See LICENCE file for information.
# Copyright (c) 2020 Theo Technicguy All Rights Reserved.
# -----------------------

import os
import requests
import json

FUTURE_WORK_DIR = os.path.join(os.path.expandvars("%APPDATA%"), "covrecord")

if not os.path.exists(FUTURE_WORK_DIR):
    os.mkdir(FUTURE_WORK_DIR)
else:
    for file in os.listdir(FUTURE_WORK_DIR):
        try:
            os.remove(os.path.join(FUTURE_WORK_DIR, file))
        except Exception:
            pass

print(FUTURE_WORK_DIR)
WORK_DIR = FUTURE_WORK_DIR

with open("covrecord.auth", "r") as file:
    AUTH = json.load(file)

# ---------- START Auto-update ----------
GITHUB_URL = (
    "https://api.github.com/repos/TheoTechnicguy/" "Etterbeek-Testing/releases"
)

# Github needs a custom header.
# Authentication is made via a token from github.
# For security, it is stored as an local_user environment variable.
header = {
    "Authenication": "token " + AUTH["github_token"],
    "accept": "application/vnd.github.v3+json",
}

# Get github repos w/ requests "GET" method.
github_page = requests.get(GITHUB_URL, headers=header)

# Check if all is ok (status_code 200)
if github_page.status_code != 200:
    # Notify of fail and write page contents into the log file.
    print("Could not autoupdate.")

else:
    # loop (once) throug releases json.
    # The `for` loop is to deal with no releases.
    for release in json.loads(github_page.text):
        for asset in release["assets"]:
            # Only get the release if does not exist.
            if not os.path.exists(os.path.join(WORK_DIR, asset["name"])):
                with open(asset["name"], "wb+") as file:
                    file.write(
                        requests.get(asset["browser_download_url"].content)
                    )

        # Only check 1st release because they are (or should be) incremental.
        break
# ---------- END Auto-update ---------

try:
    os.symlink(
        os.path.join(WORK_DIR, "covrecord.py"),
        os.path.expandvars("%USERPROFILE%\\desktop\\corecord.lnk"),
    )
except FileExistsError:
    pass
