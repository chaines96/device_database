import sqlite3, json
import requests as mreq
import pandas as pd
import os
from dotenv import load_dotenv

load_dotenv() # Load variable from any .env file in the directory. These variables will then be called in the main() function.

def zoho_produce_acc_tok(auth_code,cl_sec,cl_id):
    x = mreq.post("https://accounts.zoho.com/oauth/v2/token", data={
        "refresh_token": auth_code,
        "client_id": cl_id,
        "client_secret":cl_sec,
        "grant_type":"authorization_code"})
    return x.json()["refresh_token"]

def get_token(ref_tok,cl_id,cl_sec):
    x = mreq.post("https://accounts.zoho.com/oauth/v2/token", data={
        "refresh_token": ref_tok,
        "client_id": cl_id,
        "client_secret":cl_sec,
        "grant_type":"refresh_token"})
    zoho_resp = x.json()
    token = zoho_resp["access_token"]
    return token

def zoho_devices(zoho_token,zoho_dept_id):
    #Defining a header to use.
    acc_tok = "Bearer " + zoho_token
    my_headers = {"x-com-zoho-assist-department-id":zoho_dept_id,"Authorization":acc_tok}
    device_dict = list() #Blank dict to add JSON docs to.
    #Zoho only returns fifty devices at a time, so a list of dictionaries is created.
    for i in range(0,3):
        device_req = mreq.get("https://assist.zoho.com/api/v2/devices", params= {
            "count":"50",
            "index":str(50*i+1)
        },
        headers = my_headers)
        device_dict.append(device_req.json().copy())
    #Now that the dictionary has been created, we can visit each resource ID and get the serial number.
    device_docs = list()
    for i in range(0,3):
        for j in range(0,len(device_dict[i]["representation"]["computers"])):
            computer_req = mreq.get("https://assist.zoho.com/api/v2/devices/" + device_dict[i]["representation"]["computers"][j]["resource_id"], headers = my_headers)
            device_docs.append(computer_req.json().copy())
    return device_docs

def sw_devices(sw_token):
    my_headers = {"X-Samanage-Authorization":f"Bearer {sw_token}","Accept":"application/json"}
    device_dict = list()
    for i in range(1,5):
        device_req = mreq.get("https://api.samanage.com/hardwares",headers=my_headers, params= {
            "per_page":"100",
            "page" : str(i)
        })
        if device_req.json():
            device_dict.append(device_req.json().copy())
    return device_dict

def graph_token(azure_id,graph_id,graph_secret):
    my_headers = {"Content-Type":"application/x-www-form-urlencoded"}
    my_formdata = {
        "grant_type":"client_credentials",
        "client_id":"{}".format(graph_id), #Using format to strip away irregularities in the dotenv input.
        "client_secret":"".format(graph_secret),
        "resource":"https://graph.microsoft.com"
    }
    print("https://login.microsoftonline.com/{}/oauth2/token?Content-Type=application/x-www-form-urlencoded".format(azure_id))
    x = mreq.post("https://login.microsoftonline.com/{}/oauth2/token?Content-Type=application/x-www-form-urlencoded".format(azure_id),data=my_formdata,headers=my_headers)
    graph_resp = x.json()
    return graph_resp['access_token']

def graph_devices(azure_id,graph_id,graph_secret):
    graph_acc_tok = "Bearer " + graph_token(azure_id,graph_id,graph_secret)
    my_headers = {"Authorization":graph_acc_tok}
    dev_req = mreq.get("https://graph.microsoft.com/v1.0//deviceManagement/managedDevices",headers=my_headers)
    return dev_req.json()

def populateDatabase(dbc):
    dbc.execute('CREATE TABLE combined (serial TEXT PRIMARY KEY)') #List of serial numbers.
    dbc.execute('CREATE TABLE zoho (service_tag TEXT PRIMARY KEY, device_name TEXT, display_name TEXT, status TEXT, last_seen DATETIME)')
    dbc.execute('CREATE TABLE sw (swserial TEXT PRIMARY KEY, name TEXT, state TEXT, occupancy_state TEXT, OS TEXT, site TEXT, tag TEXT, last_seen DATETIME)')
    dbc.execute('CREATE TABLE intune (inserial TEXT PRIMARY KEY, name TEXT, type TEXT, last_seen DATETIME,username TEXT, compliance TEXT)')

def export_to_excel():
    db = sqlite3.connect('devices.db')
    try:
        df = pd.read_sql_query('SELECT * FROM combined LEFT JOIN zoho ON (combined.serial = zoho.service_tag) LEFT JOIN sw ON (combined.serial = sw.swserial) LEFT JOIN intune ON (combined.serial = intune.inserial)', db)
        df.to_excel("Converted.xlsx")
    except Exception as e:
        print(e)
    db.close()

def main():
    #Zoho Connection cariables, with default values, for refresh token, client id, and client secret.
    try:
        ZOHO_DEPT_ID = os.getenv("ZOHO_DEPT_ID")
        if  ZOHO_DEPT_ID is None:
            ZOHO_DEPT_ID = input("No department ID for Zoho found. Please enter one now:")
            with open(".env") as enviro:
                enviro.write(f"ZOHO_DEPT_ID={ZOHO_DEPT_ID}")
                print("Zoho Department ID written to .env. Please check this file on subsequent runs for valid input.")
    except:
        print("Error occured obtaining the department ID.")

    try:
        AZURE_ID = os.getenv("AZURE_ID")
        if AZURE_ID is None:
            AZURE_ID = input("No Azure ID. Please enter one now:")
            with open(".env") as enviro:
                enviro.write(f"AZURE_ID={AZURE_ID}")
                print("Azure ID written to .env. Please check this file on subsequent runs for valid input.")
    except:
        print("Error occured obtaining the Azure ID.")

    try:
        GRAPH_ID= os.getenv("GRAPH_ID")
        if GRAPH_ID is None:
            GRAPH_ID= input("No Client ID for MS Graph found. Please enter one now:")
            with open(".env") as enviro:
                enviro.write(f"GRAPH_ID={GRAPH_ID}")
                print("Graph ID written to .env. Please check this file on subsequent runs for valid input.")
    except:
        print("Error occured obtaining the Graph ID.")

    try:
        GRAPH_SECRET = os.getenv("GRAPH_SECRET")
        if GRAPH_SECRET is None:
            GRAPH_SECRET = input("No Client Secretfor MS Graph found. Please enter one now:")
            with open(".env") as enviro:
                enviro.write(f"GRAPH_SECRET={GRAPH_SECRET}")
                print("Graph secret written to .env. Please check this file on subsequent runs for valid input.")
    except:
        print("Error occured obtaining the Graph secret.")

    try:
        SW_TOKEN = os.getenv("SW_TOKEN")
        if SW_TOKEN is None:
            SW_TOKEN = input("No SolarWinds token found. Please enter one now:")
            with open(".env") as enviro:
                enviro.write(f"SW_TOKEN={SW_TOKEN}")
                print("Solarwinds Token written to .env. Please check this file on subsequent runs for valid input.")
    except:
        print("Error occured obtaining the SolarWinds token.")

    # Access the Zoho variables
    try:
        ZOHO_REFRESH_TOKEN = os.getenv("ZOHO_ACC_TO_REFRESH_TOKEN")
        ZOHO_CLIENT_ID = os.getenv("ZOHO_CLIENT_ID")
        ZOHO_CLIENT_SECRET = os.getenv("ZOHO_CLIENT_SECRET")
        if ZOHO_REFRESH_TOKEN is None or ZOHO_CLIENT_ID is None or ZOHO_CLIENT_SECRET is None:
            print("No access or refresh tokens defined." )
            ZOHO_CLIENT_ID = input("Please enter your Client ID:")
            ZOHO_CLIENT_SECRET = input("Please enter your Client Secret:")
            auth_code = input("Please enter the authentication code:")
            ZOHO_REFRESH_TOKEN = zoho_produce_acc_tok(auth_code,ZOHO_CLIENT_SECRET,ZOHO_CLIENT_ID)
            with open(".env") as enviro:
                enviro.write(f"ZOHO_CLIENT_ID={ZOHO_CLIENT_ID}")
                enviro.write(f"ZOHO_CLIENT_SECRET={ZOHO_CLIENT_SECRET}")
                enviro.write(f"ZOHO_REFRESH_TOKEN={ZOHO_REFRESH_TOKEN}")
                print("Wrote Zoho information to .env. Please check this file on subsequent runs for valid input.")
    except Exception as e:
        print(e)
    REFRESH_TOKEN = ZOHO_REFRESH_TOKEN
    CLIENT_ID = ZOHO_CLIENT_ID
    CLIENT_SECRET = ZOHO_CLIENT_SECRET
    TOKEN = ""
    TOKEN = get_token(REFRESH_TOKEN,CLIENT_ID,CLIENT_SECRET)
    dev_docs = zoho_devices(TOKEN,ZOHO_DEPT_ID)

    #SQLite database initialization.
    db = sqlite3.connect('devices.db')
    dbc = db.cursor()
    populateDatabase(dbc)
    dbc.execute('DELETE FROM combined')
    dbc.execute('DELETE FROM zoho')
    dbc.execute('DELETE FROM sw')
    dbc.execute('DELETE FROM intune')
    #Zoho Loop
    print("In zoho")
    for i in range(0,len(dev_docs)):
        try:
            my_item = dev_docs[i]["representation"]["device_info"] #makes the following statement look cleaner
            dbc.execute("INSERT OR IGNORE INTO COMBINED VALUES (?)",(dev_docs[i]["representation"]["device_info"]["service_tag"],))
            dbc.execute("INSERT OR IGNORE INTO zoho VALUES (?,?,?,?,?)",(my_item["service_tag"],dev_docs[i]["representation"]["display_name"],my_item["name"],my_item["status"],dev_docs[i]["representation"]["agent_updated_time"]))
        except Exception as e:
           print(e)
           with open("log.txt", "a") as g:
               g.write("Error occured for:\n")
               g.write(str(dev_docs[i]["representation"]))
               g.write("\n")
        with open("Total_Zoho_Devices.json", "a") as f:
           f.write(json.dumps(dev_docs[i]["representation"]))
           f.write("\n")
    #SW Loop
    sw_output = sw_devices(SW_TOKEN)
    print("In sw")
    for j in range(0,len(sw_output)):
        for i in range(0,len(sw_output[j])):
            try:
                my_item = sw_output[j][i]
                occ_state = ""
                try:
                    occ_state = my_item["custom_field_values"][0]["value"]
                except:
                    occ_state = "-"
                dbc.execute("INSERT OR IGNORE INTO COMBINED VALUES (?)",(sw_output[j][i]["serial_number"],))
                dbc.execute("INSERT OR IGNORE INTO sw VALUES (?,?,?,?,?,?,?,?)",(my_item["serial_number"],my_item["name"],my_item["status"]["name"],occ_state,my_item["operating_system"],my_item["site"]["name"],my_item["tag"],my_item["updated_at"]))
            except Exception as e:
                print(e)
                with open("log.txt", "a") as g:
                    g.write("Error occured for:\n")
                    g.write(str(sw_output[j][i]))
                    g.write("\n")
            with open("Total_SW_Devices.json", "a") as f:
                f.write(json.dumps(sw_output[j][i]))
                f.write("\n")
    #Intune Loop
    intune_output = graph_devices(AZURE_ID,GRAPH_ID,GRAPH_SECRET)
    print("In graph")
    for i in range(0,len(intune_output["value"])):
        try:
            my_item = intune_output["value"][i]
            if (intune_output["value"][i]["serialNumber"]):
                inserial = intune_output["value"][i]["serialNumber"]
            else:
                inserial = "NOSERIAL" + str(i)
            dbc.execute("INSERT OR IGNORE INTO COMBINED VALUES (?)",(inserial,))
            dbc.execute("INSERT OR IGNORE INTO intune VALUES (?,?,?,?,?,?)",(intune_output["value"][i]["serialNumber"],my_item["deviceName"],my_item["deviceCategoryDisplayName"],my_item["lastSyncDateTime"],my_item["userDisplayName"],my_item["complianceState"]))
        except Exception as e:
           print(e)
           with open("log.txt", "a") as g:
               g.write("Error occured for:\n")
               g.write(str(intune_output["value"][i]))
               g.write("\n")
    with open("Total_Intune_Devices.json", "a") as f:
        f.write(json.dumps(intune_output["value"]))
    db.commit()
    db.close()

main()
export_to_excel()
