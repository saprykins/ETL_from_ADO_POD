import requests
import pandas as pd
import requests
import base64
import openpyxl
import csv


pat = 'y**'
organization = 'go**'
project = 'P**'
authorization = str(base64.b64encode(bytes(':'+pat, 'ascii')), 'ascii')





def extract_application_data(organization, project, authorization):
    ids_url = f"https://dev.azure.com/{organization}/{project}/_apis/wit/wiql/7d714bbe-6bf5-4823-afa7-15bb8aef7da8" # all in ms projects
    
    ids_headers = {
        'Accept': 'application/json',
        'Authorization': 'Basic ' + authorization
    }
    ids_response = requests.get(url=ids_url, headers=ids_headers)
    application_ids = [work_item["id"] for work_item in ids_response.json()["workItems"]]

    # Get detailed information for each application
    application_data = []
    for application_id in application_ids:
        details_url = f"https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/{application_id}?api-version=6.0"
        details_response = requests.get(url=details_url, headers=ids_headers)
        details_json = details_response.json()

        # Check if the required fields exist in the response before accessing them
        app_id = details_json.get("id")
        name = details_json["fields"].get("System.Title", "N/A")
        status = details_json["fields"].get("System.State", "N/A")
        env = details_json["fields"].get("Custom.AppEnvironment", "N/A")
        dc = details_json["fields"].get("Custom.DC", "N/A")
        plnd_wave = details_json["fields"].get("Custom.Waves", "N/A")

        # Create the application_info dictionary with the available data
        application_info = {
            "app_id": app_id,
            "App name": name,
            # "status": status,
            "Environment":env,
            "Data center": dc,
            "Planned Wave": plnd_wave
        }
        application_data.append(application_info)

    return application_data




def extract_server_data(organization, project, authorization):
    # Modify the URL to point to the appropriate WIQL query for servers
    ids_url = f"https://dev.azure.com/{organization}/{project}/_apis/wit/wiql/822cf75e-db2d-4cf0-b2dc-ab74a70450c6"
    # print(ids_url)

    ids_headers = {
        'Accept': 'application/json',
        'Authorization': 'Basic ' + authorization
    }
    ids_response = requests.get(url=ids_url, headers=ids_headers)
    server_ids = [work_item["id"] for work_item in ids_response.json()["workItems"]]


    server_data = []
    for server_id in server_ids:
        details_url = f"https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/{server_id}?$expand=all&api-version=6.0"
        details_response = requests.get(url=details_url, headers=ids_headers)
        details_json = details_response.json()

        # Check if the required fields exist in the response before accessing them
        serv_id = details_json.get("id")
        name = details_json["fields"].get("System.Title", "N/A")
        status = details_json["fields"].get("System.State", "N/A")
        plnd_mig = details_json["fields"].get("Custom.Initialplandate", "N/A")
        # actl_mig_strt_date = details_json["fields"].get("Microsoft.VSTS.Scheduling.StartDate", "N/A")
        # actl_mig_target_date = details_json["fields"].get("Microsoft.VSTS.Scheduling.TargetDate", "N/A")
        act_mig_date = details_json["fields"].get("Custom.MigrationDate", "N/A")
        
        out_biz_hours = details_json["fields"].get("Custom.Nonworkinghours", "N/A")
        app_id = details_json["fields"].get("System.Parent", "N/A")

        # Create the server_info dictionary with the available data
        server_info = {
            "Server id in ADO": serv_id,
            "Server": name,
            "State": status,
            "Actual migration startdate": act_mig_date,
            # "actl_mig_target_date": actl_mig_target_date,
            "Out of business hours": out_biz_hours,
            "Planned migration date": plnd_mig, 
            "app_id": app_id,
            # Add more fields as needed
        }
        server_data.append(server_info)

    return server_data




# id 333055 - app ex
# id 333056 - server ex

# 1. Data Extraction from Azure DevOps via API
def fetch_data_from_azure_devops():
    # Make API requests and fetch data
    # Authentication and API requests code here
    application_data = extract_application_data(organization, project, authorization)
    server_data = extract_server_data(organization, project, authorization)
    return application_data, server_data



# 2. Data Transformation and Manipulation
def transform_and_merge_data(application_data, server_data):
    # Convert data to dataframes using pandas
    application_df = pd.DataFrame(application_data)
    server_df = pd.DataFrame(server_data)

    # 3. Data Merging
    merged_df = pd.merge(application_df, server_df, on='app_id', how='inner')

    # 6. Sort columns as in MPI
    return merged_df



def sort_columns(merged_data):
    # Additional columns
    '''
    additional_columns = ['Server id in ADO', 'Server', 'App name', 'Environment', 'State', 'Entity',
                          'Planned migration date', 'Actual migration startdate', 'Actual migration enddate',
                          'Data center', 'Planned Wave', 'Out of business hours']
    '''
    additional_columns = ['Server id in ADO', 'Server', 'FQDN', 'Sign-off Ops', 'Sign-off DBA_x', 
                      'App name', 'Environment', 'State', 'Entity', 'Planned migration date', 
                      'Actual migration startdate', 'Actual migration enddate', 'Data center', 
                      'Blocker details', 'De-scoping Details', 'Flow opening confirmation', 
                      'Last minute reschedule', 'Migration eligibility', 'Planned Wave', 
                      'Internet access through proxies', 'Outbound Emails', 'Reverse Proxies', 
                      'WAC', 'WAF', 'VPN', 'Load Balancer', 'Service Account in local AD domains', 
                      'Encryption', 'Secret data', 'Fileshare', 
                      'Administration through specific Jump servers', 
                      'Access through specific Citrix Jump servers', 'Out of business hours', 
                      'Zero downtime requirements', 'Risk level', 'Factory', 'Sign-off DBA_y', 
                      'Sign-off Entity', 'Schedule_change_Description', 'Phases', 
                      'Internet access through proxies', 'Service Account in local AD domains']


    # Remove extra columns not present in additional_columns
    merged_data = merged_data[merged_data.columns.intersection(additional_columns)]

    # Add missing columns with NaN values
    for column in additional_columns:
        if column not in merged_data.columns:
            merged_data[column] = None

    # Rearrange columns in the specified order
    merged_data = merged_data[additional_columns]

    return merged_data


# 6. Export
def export_to_excel(data_frame, excel_file_path):
    data_frame.to_excel(excel_file_path, index=False)

def export_to_csv(data, filename):
    data.to_csv(filename, index=False, encoding='utf-8')






def convert_csv_to_excel(csv_file, excel_file):
    # Read the CSV file into a pandas DataFrame
    data = pd.read_csv(csv_file)

    # Write the DataFrame to an Excel file
    data.to_excel(excel_file, index=False)













if __name__ == "__main__":
    application_data, server_data = fetch_data_from_azure_devops()
    merged_data = transform_and_merge_data(application_data, server_data)
    
    sorted_data = sort_columns(merged_data)
    print(sorted_data)
    sorted_data.to_csv('output2.csv', index=False, encoding='utf-8')
    convert_csv_to_excel('output2.csv', 'output2.xlsx')
    # sorted_data.to_excel('output2.xlsx', index=False)
    # export_to_excel(sorted_data, 'output2.xlsx')



