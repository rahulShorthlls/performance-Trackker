
# Azure DevOps Work Item Summary Tool

This Python script retrieves and summarizes work items (tickets) from Azure DevOps for given users over a specified date range. It collects detailed ticket statistics across all organizations and projects the user has access to and exports the result to an Excel file.

## 🔧 Features

- Fetches **all organizations** associated with the given PAT.
- Iterates through all **projects** within each organization.
- Queries **completed work items** (`System.State = 'Done'`) assigned to or created by specified users.
- Classifies work items based on task type:
  - Coding/Implementation
  - Design
  - Code Review
  - Others
- Computes **FTAR** (custom field) statistics for coding tasks.
- Outputs a structured Excel summary per user.

## 📁 Input Format

An Excel file (`input.xlsx`) with the following columns:
| emails               | days |
|----------------------|------|
| user@example.com     | 30   |
| another@domain.com   | 15   |

- `emails`: Azure DevOps account email addresses.
- `days`: Number of past days to search for completed work items.

## 📤 Output

The script creates an Excel file `done_ticket_summary.xlsx` with the following columns:

| Name/Email id       | Number of days (X) | Number of tickets in last X days | Coding | Design | Code review | Others | If Coding, then average of FTAR |
|---------------------|--------------------|----------------------------------|--------|--------|--------------|--------|----------------------------------|
| user@example.com    | 30                 | 12                               | 6      | 2      | 3            | 1      | 0.83                             |

## ⚙️ Configuration

Edit the following values at the top of the script:

```python
pat = "YOUR_PERSONAL_ACCESS_TOKEN"
member_id = "YOUR_AZURE_DEVOPS_MEMBER_ID"
input_file = "input.xlsx"  # Make sure this file exists with proper format
````

> 🔐 **Note**: Always keep your PAT secret. Consider using environment variables or a config file in production.

## ✅ Requirements

Install required packages:

```bash
pip install pandas openpyxl requests
```

Ensure your PAT has at least **read access** to work items across organizations.

## 🏁 Running the Script

```bash
python script_name.py
```

Replace `script_name.py` with the actual filename. The summary will be saved as `done_ticket_summary.xlsx` in the current directory.

---

## 📌 Notes

* Task type detection relies on the `Microsoft.VSTS.CMMI.TaskType` field. Adjust the script if your organization uses a different process template or custom fields.
* The FTAR field is assumed to be `Custom.FTARValue`. Change it as needed.
* Handles errors and missing fields gracefully, but review logs for failed requests.

## ✍️ Author

Rahul Kumar
Intern at Shorthills AI
GitHub: \[your-github-link-here]

```
