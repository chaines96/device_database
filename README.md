# The Device Database generator!
When managing IT resources, it is convenient to be able to correlate device information from multiple products and pieces of software. For my purposes, I like to connect Microsoft Graph, SolarWinds, and Zoho Assist and place them in a database for easy viewing! 

# Prerequisites
This program is meant for use with Python 3.14 and uses the **sqlite3** and **json** modules included in the standard library. It also uses **pandas** (to export data an Excel document), **os** (for file management), **dotenv** (for environment variables) and **requests** (for HTTP requests). To install these, run:

```
pip install pandas dotenv requests os
```

# Operation
To run this yourself, run:

```git clone github.com/chaines96/device_database.git```

Navigate to the program directory and create a file named *.env* with the following fields filled in:

```
ZOHO_DEPT_ID=
SW_TOKEN=
AZURE_ID=
GRAPH_ID=
GRAPH_SECRET=
ZOHO_ACC_TO_REFRESH_TOKEN=
ZOHO_CLIENT_ID=
ZOHO_CLIENT_SECRET=
```

With python installed, run: 

```
python device_database.py
```

# Roadmap

Currently, the program can run cleanly on a first attempt and produces an excel file. I intend to change this in the future. Plans include:

* Fix the dotenv calls in main() so that you receive a prompt to enter the variables during runtime if they are not in the .env file.
* Give the user an interactive prompt that gives them a choice between viewing existing data or refreshing it.
* Provide instructions and facilitate using the **datasette** module to view the entries in the SQLite database. 

Other features may include:
* Allow the user to specify which of the three services they will connect to, as the program fails if it cannot connect to any one of SolarWinds, MS Graph, or Zoho Assist.
