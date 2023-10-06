import requests
import pandas as pd
import requests
import base64

pat = 'y**'
organization = 'go**'
project = 'P**'
authorization = str(base64.b64encode(bytes(':'+pat, 'ascii')), 'ascii')






def extract_application_data(organization, project, authorization):
    """
    
    """
    ids_url = "https://dev.azure.com/" + organization + "/" + project + "/_apis/wit/wiql/7d714bbe-6bf5-4823-afa7-15bb8aef7da8" # all in ms projects
    
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
        application_info = {
            "app_id": details_response.json()["id"],
            "name": details_response.json()["fields"]["System.Title"],
            "status": details_response.json()["fields"]["System.State"],
            # Add more fields as needed
        }
        application_data.append(application_info)

    return application_data


def extract_server_data(organization, project, authorization):
    """
    Extracts data related to servers from Azure DevOps.
    """
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

        # print(details_url)
        details_response = requests.get(url=details_url, headers=ids_headers)
        server_info = {
            "serv_id": details_response.json()["id"],
            "name": details_response.json()["fields"]["System.Title"],
            "status": details_response.json()["fields"]["System.State"],
            "app_id": details_response.json()["fields"]["System.Parent"],
            # Add more fields as needed
        }
        server_data.append(server_info)

    return server_data


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

    # 4. Column Renaming
    # merged_df = merged_df.rename(columns={'old_column_name': 'new_column_name'})

    # 5. Adding Additional Columns
    # merged_df['new_column'] = ...

    return merged_df


applications, servers = fetch_data_from_azure_devops()
merged_data = transform_and_merge_data(applications, servers)
print(merged_data)

'''
# 6. Excel Export
def export_to_excel(data_frame, excel_file_path):
    data_frame.to_excel(excel_file_path, index=False)



if __name__ == "__main__":
    application_data, server_data = fetch_data_from_azure_devops()
    merged_data = transform_and_merge_data(application_data, server_data)
    export_to_excel(merged_data, 'output.xlsx')
'''
