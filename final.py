import requests
import pandas as pd
from base64 import b64encode
from datetime import datetime, timedelta

# -------------- CONFIG --------------
pat = "PAT here"
member_id = 'member id here'
# ftar_field = "Microsoft.VSTS.Scheduling.Effort"
input_file = "input.xlsx"  # Excel file with 'emails' and 'days' columns
# ------------------------------------

# Read emails and number of days from Excel
df_input = pd.read_excel(input_file)
auth = b64encode(f':{pat}'.encode()).decode()
headers = {
    'Authorization': f'Basic {auth}',
    'Content-Type': 'application/json'
}

# Get all organizations
orgs_url = f"https://app.vssps.visualstudio.com/_apis/accounts?memberId={member_id}&api-version=7.0"
orgs_response = requests.get(orgs_url, headers=headers)
organizations = [org['accountName'] for org in orgs_response.json().get('value', [])]

print(f"\n🔎 Found {len(organizations)} organizations.")
output = []

for idx, row in df_input.iterrows():
    email = row['emails']
    num_days = int(row['days'])

    total_done = coding = design = review = other = ftar_sum = 0
    print(f"\n📌 Searching for user: {email} for past {num_days} days")

    today = datetime.utcnow()
    start_date = (today - timedelta(days=num_days)).strftime('%Y-%m-%d')
    end_date = today.strftime('%Y-%m-%d')

    for org in organizations:
        print(f"\n🌐 Organization: {org}")
        projects_url = f"https://dev.azure.com/{org}/_apis/projects?api-version=7.0"
        projects_response = requests.get(projects_url, headers=headers)

        if projects_response.status_code != 200:
            print(f"❌ Failed to fetch projects for org: {org}")
            continue

        projects = projects_response.json().get('value', [])
        print(f"🔧 Found {len(projects)} projects in {org}")

        for project in projects:
            project_name = project['name']
            print(f"   🔍 Searching project: {project_name}")

            wiql = {
                "query": f"""
                SELECT [System.Id]
                FROM WorkItems
                WHERE
                    ([System.AssignedTo] = '{email}' OR [System.CreatedBy] = '{email}')
                    AND [System.ChangedDate] >= '{start_date}'
                    AND [System.ChangedDate] <= '{end_date}'
                    AND [System.State] = 'Done'
                ORDER BY [System.ChangedDate] DESC
                """
            }

            wiql_url = f"https://dev.azure.com/{org}/{project_name}/_apis/wit/wiql?api-version=7.0"
            wiql_response = requests.post(wiql_url, headers=headers, json=wiql)

            if wiql_response.status_code != 200:
                print(f"      ⚠️ WIQL query failed for {project_name}")
                continue

            work_items = wiql_response.json().get('workItems', [])
            if not work_items:
                print(f"      📭 No done work items found.")
                continue

            ids = ','.join(str(wi['id']) for wi in work_items)
            details_url = (
                f"https://dev.azure.com/{org}/_apis/wit/workitems?ids={ids}&fields=" +
                ",".join([
                    "System.WorkItemType",
                    "System.State",
                    "System.Tags",
                    
                ]) + "&api-version=7.0"
            )
            details_response = requests.get(details_url, headers=headers)

            if details_response.status_code != 200:
                print(f"      ❌ Failed to fetch work item details.")
                continue

            for item in details_response.json()['value']:
                task_url = item['url']
                task_details = requests.get(task_url + "?api-version=7.0", headers=headers)

                if task_details.status_code != 200:
                    print(f"      ❌ Failed to fetch full data for task ID {item['id']}")
                    continue

                fields = task_details.json().get("fields", {})
                task_type = fields.get("Microsoft.VSTS.CMMI.TaskType", "").strip()
                ftar_value = fields.get("Custom.FTARValue")

                total_done += 1

                if task_type == "Code Review":
                    review += 1
                elif task_type == "Design":
                    design += 1
                elif task_type == "Coding/Implementation":
                    coding += 1
                else:
                    other += 1

                if ftar_value == 1 and task_type == "Coding/Implementation":
                    ftar_sum += 1

                print(f"Task ID: {item['id']} - Task Type: {task_type or 'N/A'}")

    output.append({
        "Name/Email id": email,
        "Number of days (X)": num_days,
        "Number of tickets in last X days": total_done,
        "Coding": coding,
        "Design": design,
        "Code review": review,
        "Others": other,
        "If Coding, then average of FTAR": ftar_sum / coding if coding >= 1 else 0
    })

# Save to Excel instead of CSV
df_output = pd.DataFrame(output)
df_output.to_excel("done_ticket_summary.xlsx", index=False)
print("\n✅ Done. File saved as 'done_ticket_summary.xlsx'")
