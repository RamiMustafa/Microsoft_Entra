# Microsoft Entra Governance PowerShell Scripts

PowerShell scripts for two Microsoft Entra governance tasks:

1. **Group access review automation** — create recurring access reviews for groups listed in a CSV file, with group owners as reviewers.
2. **Privileged-role and PIM reporting** — export role assignments, PIM activation history, and an execution summary.

Maintained by Rami Mustafa.

## Repository contents

```text
.
├── New-GroupAccessReviews.ps1
├── <privileged-role-report-script>.ps1
├── README.md
└── groups.example.csv
```

Replace `<privileged-role-report-script>.ps1` with the actual filename before publishing.

## Prerequisites

- PowerShell and the Microsoft Graph PowerShell modules required by each script.
- Appropriate Microsoft Graph permissions, Microsoft Entra roles, and licensing for the operation being performed.
- Access to the target tenant.

Review each script's required permissions before running it. Do not commit customer-specific group lists, tenant details, credentials, or generated reports to a public repository.

---

## 1. Group access review automation

`New-GroupAccessReviews.ps1` connects to a specified tenant and processes group IDs from a CSV file. For each group, it checks for an owner and an existing individual-group review before attempting to create a recurring membership review.

The script assigns the group's owners as reviewers, displays progress and per-group outcomes, and supports `-WhatIf`. It does **not** configure a fallback reviewer or automatically apply review decisions.

### Input

Create a CSV file with a `GroupId` column:

```csv
GroupId
11111111-1111-1111-1111-111111111111
22222222-2222-2222-2222-222222222222
```

Use example IDs in the repository, not IDs from a customer tenant.

### Preview

```powershell
.\New-GroupAccessReviews.ps1 `
    -TenantId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' `
    -CsvPath '.\groups.csv' `
    -WhatIf
```

`-WhatIf` performs the read checks but does not create reviews. `WhatIf: not created` means the group reached the creation step; it does not guarantee that Microsoft Graph will accept the creation request.

### Create reviews

After checking the preview:

```powershell
.\New-GroupAccessReviews.ps1 `
    -TenantId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' `
    -CsvPath '.\groups.csv'
```

The script reports which groups were created, skipped, or failed, along with a final summary. Read a failed group's full `Error` value before retrying.

### Review settings and group eligibility

The script's defaults are a 14-day review every three months, with the first review starting in seven days. Notifications and reminders are enabled; decisions are not automatically applied.

Check the **version of the script committed to this repository** for its exact eligibility rules. The original version skips dynamic and role-assignable groups. A role-assignable group can be reviewed, but should be assessed before removing that safeguard. Groups governed by PIM for Groups require separate consideration because this workflow does not specify a fallback reviewer.

To keep the group ID out of newly created review names, use this display-name setting in the script:

```powershell
displayName = "Quarterly group membership review - $groupName"
```

The group ID must remain in the review scope and reviewer query. Changing the display-name setting does not rename an existing review.

---

## 2. Privileged-role and PIM reporting

The privileged-role reporting script captures:

- Microsoft Entra privileged roles and their assigned users.
- Assignment status, including **Active** and **Eligible** assignments.
- Assignment details, including the assigner and assignment time.
- PIM activation activity, including approvals, justification, and activation time.
- Execution and tenant-level summary information.

### Generated files

When executed, the reporting script generates three files:

1. **`Entra-PrivilegedRoleAssignments.csv`** — role information and assigned users, including assignment status, assigned by, and assignment time.
2. **`Entra-PIMActivationHistory.csv`** — PIM role activation activity, including activation details, approvals, justification, and activation time.
3. **`Entra-PIMReport-Summary.json`** — execution summary, including tenant information, execution time, Active and Eligible assignment totals, and the location of the generated results.

### Run the reporting script

Add the reporting script's actual command, parameters, permissions, and output-directory behavior here before publishing. Those details should be taken from the script itself rather than assumed from the report filenames.

## Security and publication notes

The reporting outputs can contain sensitive information about privileged access. Store them securely and do not commit generated CSV or JSON reports to the repository. Start the access-review script with a small test CSV and `-WhatIf` before a bulk run.
