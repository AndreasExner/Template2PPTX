# ============================================================
# Test: ReplaceImages — Slide1_Picture1
# ============================================================
# Downloads car.jpg from Azure Blob Storage and sends it
# to the ReplaceImages endpoint to replace shape "Slide1_Picture1".
# ============================================================

$apiBase   = "https://ttemplate2pptx.azurewebsites.net/api"
$apiKey    = "<YOUR_FUNCTION_KEY>"
$template  = "Test\Template_Demo.pptx"
$output    = "Test\Slide1_Picture1_result.pptx"

# --- Image source ---
$imageUrl  = "https://sttemplate2pptx.blob.core.windows.net/templates/car.jpg?sp=r&st=2026-02-26T11:32:03Z&se=2026-02-26T19:47:03Z&spr=https&sv=2024-11-04&sr=b&sig=r2KEkPlSWrP5rgzNqGgeOn8YlsqzzfsaRXqI52pVRow%3D"
$imageFile = "Test\car.jpg"

# --- Step 1: Download image ---
Write-Host "Downloading car.jpg ..."
Invoke-WebRequest -Uri $imageUrl -OutFile $imageFile
Write-Host "  -> Saved to $imageFile ($(((Get-Item $imageFile).Length / 1KB).ToString('N1')) KB)"

# --- Step 2: Call ReplaceImages ---
Write-Host "Calling ReplaceImages ..."
$form = @{
    template         = Get-Item $template
    Slide1_Picture1  = Get-Item $imageFile
}
Invoke-RestMethod -Uri "$apiBase/ReplaceImages?code=$apiKey" `
                  -Method Post `
                  -Form $form `
                  -OutFile $output

Write-Host "  -> Result saved to $output ($(((Get-Item $output).Length / 1KB).ToString('N1')) KB)"
Write-Host "Done!"
