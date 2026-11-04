<#
.SYNOPSIS
    Creates recurring Microsoft Entra group access reviews from a CSV file.

.DESCRIPTION
    Connects to a specified Microsoft Entra tenant and processes group IDs
    from a CSV file. For each eligible group, the script checks for an owner
    and an existing individual-group review before creating a recurring
    membership review with the group owners as reviewers.

    The script displays per-group progress and a final summary. It supports
    -WhatIf so that review creation can be previewed.

    No fallback reviewer is configured, and review decisions are not
    automatically applied.

    Review the eligibility checks in the script before a bulk run. In
    particular, the script version originally shared excludes dynamic and
    role-assignable groups. If you have removed the role-assignable exclusion,
    assess groups governed by PIM for Groups separately.

.PARAMETER TenantId
    GUID of the Microsoft Entra tenant to connect to.

.PARAMETER CsvPath
    Path to a CSV file with a GroupId column.

.PARAMETER DurationDays
    Number of days each review instance remains open. Default: 14.

.PARAMETER IntervalMonths
    Number of months between review instances. Default: 3.

.PARAMETER StartInDays
    Number of days from now until the first review starts. Default: 7.

.EXAMPLE
    .\New-GroupAccessReviews.ps1 `
        -TenantId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' `
        -CsvPath '.\groups.csv' `
        -WhatIf

    Previews which groups would have reviews created.

.EXAMPLE
    .\New-GroupAccessReviews.ps1 `
        -TenantId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' `
        -CsvPath '.\groups.csv'

    Creates reviews for eligible groups.

.NOTES
    Requires the Microsoft.Graph.Authentication PowerShell module,
    appropriate Microsoft Graph permissions, an authorized Microsoft Entra
    role, and applicable access-review licensing.

    Test with -WhatIf and a small CSV before a bulk run. -WhatIf prevents
    creation but still performs read requests against Microsoft Graph.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [guid] $TenantId,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string] $CsvPath,

    [ValidateRange(1, 365)]
    [int] $DurationDays = 14,

    [ValidateRange(1, 12)]
    [int] $IntervalMonths = 3,

    [ValidateRange(1, 365)]
    [int] $StartInDays = 7
)

$ErrorActionPreference = 'Stop'

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

Connect-MgGraph `
    -TenantId $TenantId.ToString() `
    -Scopes @('AccessReview.ReadWrite.All', 'Group.Read.All') `
    -NoWelcome `
    -ErrorAction Stop

$context = Get-MgContext

if (-not $context -or $context.TenantId -ne $TenantId.ToString()) {
    throw "Connected tenant does not match requested tenant $TenantId."
}

Write-Host "Connected to tenant: $($context.TenantId)"

function Invoke-GraphGet {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Uri
    )

    Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
}

$rows = @(Import-Csv -LiteralPath $CsvPath)

if ($rows.Count -eq 0) {
    throw 'The CSV is empty.'
}

if (-not ($rows[0].PSObject.Properties.Name -contains 'GroupId')) {
    throw 'The CSV must have a GroupId column.'
}

$groupIds = @(
    $rows |
        ForEach-Object { "$($_.GroupId)".Trim() } |
        Where-Object { $_ -ne '' } |
        Select-Object -Unique
)

if ($groupIds.Count -eq 0) {
    throw 'The CSV contains no group IDs.'
}

$start = (Get-Date).ToUniversalTime().AddDays($StartInDays)

$startDate = $start.ToString(
    'yyyy-MM-ddTHH:mm:ssZ',
    [System.Globalization.CultureInfo]::InvariantCulture
)

$results = @()
$total = $groupIds.Count
$processed = 0

foreach ($inputGroupId in $groupIds) {
    $groupName = ''
    $result = $null

    Write-Progress `
        -Activity 'Creating group access reviews' `
        -Status "Processing $($processed + 1) of ${total}: $inputGroupId" `
        -PercentComplete (100 * $processed / $total)

    try {
        # A failed GUID cast is caught below and reported for this group.
        $groupId = ([guid] $inputGroupId).ToString()

        $group = Invoke-GraphGet -Uri (
            "/v1.0/groups/$groupId" +
            '?$select=id,displayName,groupTypes,membershipRule,isAssignableToRole'
        )

        $groupName = $group.displayName

        if (
            $group.membershipRule -or
            @($group.groupTypes) -contains 'DynamicMembership' -or
            $group.isAssignableToRole
        ) {
            $result = [pscustomobject]@{
                GroupId      = $groupId
                Group        = $groupName
                Status       = 'Skipped: special group; assess separately'
                DefinitionId = ''
                Error        = ''
            }
        }
        else {
            $owners = Invoke-GraphGet -Uri (
                "/v1.0/groups/$groupId/owners" +
                '?$select=id&$top=1'
            )

            if (@($owners.value).Count -eq 0) {
                $result = [pscustomobject]@{
                    GroupId      = $groupId
                    Group        = $groupName
                    Status       = 'Skipped: no group owner'
                    DefinitionId = ''
                    Error        = ''
                }
            }
            else {
                $filter = (
                    "contains(" +
                    "scope/microsoft.graph.accessReviewQueryScope/query, " +
                    "'/groups/$groupId'" +
                    ")"
                )

                $encodedFilter = [System.Uri]::EscapeDataString($filter)

                $existing = Invoke-GraphGet -Uri (
                    '/v1.0/identityGovernance/accessReviews/definitions' +
                    '?$filter=' + $encodedFilter +
                    '&$top=1'
                )

                if (@($existing.value).Count -gt 0) {
                    $result = [pscustomobject]@{
                        GroupId      = $groupId
                        Group        = $groupName
                        Status       = 'Skipped: review already exists'
                        DefinitionId = $existing.value[0].id
                        Error        = ''
                    }
                }
                else {
                    $body = @{
                        displayName = (
                            "Quarterly group membership review - " +
                            "$groupName"# [$groupId]"
                        )

                        descriptionForAdmins = (
                            "Recurring membership review for group $groupId."
                        )

                        descriptionForReviewers = (
                            'Review whether each member still needs ' +
                            'access to this group.'
                        )

                        scope = @{
                            '@odata.type' = (
                                '#microsoft.graph.accessReviewQueryScope'
                            )
                            query     = "/groups/$groupId/transitiveMembers"
                            queryType = 'MicrosoftGraph'
                        }

                        reviewers = @(
                            @{
                                query     = "/groups/$groupId/owners"
                                queryType = 'MicrosoftGraph'
                            }
                        )

                        settings = @{
                            mailNotificationsEnabled        = $true
                            reminderNotificationsEnabled    = $true
                            justificationRequiredOnApproval = $true
                            recommendationsEnabled          = $false
                            defaultDecisionEnabled          = $false
                            autoApplyDecisionsEnabled       = $false
                            instanceDurationInDays          = $DurationDays

                            recurrence = @{
                                pattern = @{
                                    type       = 'absoluteMonthly'
                                    interval   = $IntervalMonths
                                    dayOfMonth = $start.Day
                                }

                                range = @{
                                    type      = 'noEnd'
                                    startDate = $startDate
                                }
                            }
                        }
                    }

                    if ($PSCmdlet.ShouldProcess(
                        "$groupName ($groupId)",
                        'Create recurring access review'
                    )) {
                        $json = $body |
                            ConvertTo-Json -Depth 12 -Compress

                        $created = Invoke-MgGraphRequest `
                            -Method POST `
                            -Uri '/v1.0/identityGovernance/accessReviews/definitions' `
                            -ContentType 'application/json' `
                            -Body $json `
                            -ErrorAction Stop

                        $result = [pscustomobject]@{
                            GroupId      = $groupId
                            Group        = $groupName
                            Status       = 'Created'
                            DefinitionId = $created.id
                            Error        = ''
                        }
                    }
                    else {
                        $result = [pscustomobject]@{
                            GroupId      = $groupId
                            Group        = $groupName
                            Status       = 'WhatIf: not created'
                            DefinitionId = ''
                            Error        = ''
                        }
                    }
                }
            }
        }
    }
    catch {
        $result = [pscustomobject]@{
            GroupId      = $inputGroupId
            Group        = $groupName
            Status       = 'Failed'
            DefinitionId = ''
            Error        = $_.Exception.Message
        }
    }

    $results += $result
    $processed++

    $label = if ($result.Group) {
        $result.Group
    }
    else {
        $result.GroupId
    }

    Write-Host "[$processed/$total] $($result.Status): $label"

    if ($result.Error) {
        Write-Host "  Error: $($result.Error)"
    }

    Write-Progress `
        -Activity 'Creating group access reviews' `
        -Status "Completed $processed of $total" `
        -PercentComplete (100 * $processed / $total)
}

Write-Progress -Activity 'Creating group access reviews' -Completed

$createdCount = @($results | Where-Object { $_.Status -eq 'Created' }).Count
$skippedCount = @($results | Where-Object { $_.Status -like 'Skipped*' }).Count
$failedCount = @($results | Where-Object { $_.Status -eq 'Failed' }).Count

Write-Host ''
Write-Host (
    "Finished: $processed processed; " +
    "$createdCount created; " +
    "$skippedCount skipped; " +
    "$failedCount failed."
)

$results | Format-List GroupId, Group, Status, DefinitionId, Error