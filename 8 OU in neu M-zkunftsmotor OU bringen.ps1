# Set domain
$domain = "M-zukunftsmotor.local"

# List of OUs to create under the new OU
$ouList = @(
    "BuHa12A",
    "BuHa12B",
    "BuHa13A",
    "BuHa13B",
    "IT14A",
    "IT14B",
    "IT15A",
    "IT15B"
)

# Name of the new OU
$newParentOU = "M-zukunftsmotor"

# Check if the new parent OU exists
$parentOUExists = Get-ADOrganizationalUnit -Filter "Name -eq '$newParentOU'" -SearchBase "DC=M-zukunftsmotor,DC=local" -ErrorAction SilentlyContinue

if (-not $parentOUExists) {
    # Create the parent OU
    New-ADOrganizationalUnit -Name $newParentOU -Path "DC=M-zukunftsmotor,DC=local"
    Write-Host "Created parent OU: $newParentOU"
} else {
    Write-Host "Parent OU already exists: $newParentOU"
}

# Loop to create each OU under the new parent OU
foreach ($ouName in $ouList) {
    # Check if OU exists
    $ouExists = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -SearchBase "OU=$newParentOU,DC=M-zukunftsmotor,DC=local" -ErrorAction SilentlyContinue

    if (-not $ouExists) {
        # Create OU under the new parent OU
        New-ADOrganizationalUnit -Name $ouName -Path "OU=$newParentOU,DC=M-zukunftsmotor,DC=local"
        Write-Host "Created OU: $ouName under $newParentOU"
    } else {
        Write-Host "OU already exists: $ouName"
    }
}
