# Set domain
 
$domain = "M-zukunftsmotor.local"

# List of OUs to create
 
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

# Loop to create each OU
 
foreach ($ouName in $ouList) {
 
    # Check if OU exists
 
    $ouExists = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -SearchBase "DC=M-zukunftsmotor,DC=local" -ErrorAction SilentlyContinue
 
    if (-not $ouExists) {
 
        # Create OU
 
        New-ADOrganizationalUnit -Name $ouName -Path "DC=M-zukunftsmotor,DC=local"
 
        Write-Host "Created OU: $ouName"
 
    } else {
 
        Write-Host "OU already exists: $ouName"
 
    }
 
}
 
 