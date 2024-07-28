Connect-AzureAD
# Fetch all SKUs
$skus = Get-AzureADSubscribedSku

# Display SKUs
$skus | ForEach-Object {
    Write-Host "SKU ID: $($_.SkuId) - SKU Part Number: $($_.SkuPartNumber)"
}

# Fetch a user's license information
$user = Get-AzureADUser -ObjectId "UPN"
$licenses = Get-AzureADUserLicenseDetail -ObjectId $user.ObjectId

# Display user license information
$licenses | ForEach-Object {
    Write-Host "Service Plan: $($_.ServicePlans)"
}
