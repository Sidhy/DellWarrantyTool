# Dell Warranty Check
# V1.0.1
# Sidhy (Jorrit)


$Script:DebugProxy = $false # Set to true to use local proxy for debugging
$Script:UseSystemProxy = $true
$Script:UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:104.0) Gecko/20100101 Firefox/104.0"

$Script:Url_Base = "https://www.dell.com"
$Script:Url_MyProducts = "${Url_Base}/support/mps/en-uk/myproducts"
$Script:Url_ListProducts = "${Url_Base}/support/mps/en-uk/myproductgridrecords"
$Script:Url_RemoveProduct = "${Url_Base}/support/mps/en-uk/removefromproduct"
$Script:Url_AddProduct = "${Url_Base}/support/mps/en-uk/saveproductsusingtag"


#region disable ssl check
if ($Script:DebugProxy)
{
add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
}
#endregion

class Asset
{
    [string]$ProductName
    [string]$ServiceTag
    [string]$ShipDate
    [string]$WarrantyType
    [string]$WarrantyEndDate = "Unknown"

    Asset($ProductName, $ServiceTag, $ShipDate, $WarrantyType, $WarrantyEndDate) {
       $this.ProductName = $ProductName
       $this.ServiceTag = $ServiceTag
       $this.ShipDate = $ShipDate
       $this.WarrantyType = $WarrantyType
       $this.WarrantyEndDate = $WarrantyEndDate 
    }
}

function IIf($If, $IfTrue, $IfFalse) {
    If ($If) {If ($IfTrue -is "ScriptBlock") {&$IfTrue} Else {$IfTrue}}
    Else {If ($IfFalse -is "ScriptBlock") {&$IfFalse} Else {$IfFalse}}
}

$Script:AssetList = @{}



#region Cookie Loader

Add-Type -AssemblyName System.Web
$Script:WebSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
function ProcessCookie 
{
    param(
        $Cookie
    )
    $Cookie = $Cookie.Split('; ')
    foreach($line in $Cookie)
    {
        if ($line.Length -gt 0)
        {
            try
            {
                $var = $line.split("=")[0]

                $start = $line.IndexOf("=")+1
                $end = $line.Length-$start
                $val = $line.Substring($start,$end)
                $val = [System.Web.HttpUtility]::UrlEncode($val) # Url encoding
                $c = New-Object System.Net.Cookie($var,$val ,'/','.dell.com')
                $Script:WebSession.Cookies.Add($c)
            }
            catch
            { continue }

        }
    }
    Write-Host "Cookies loaded"
}

function Set-DellCookie 
{
    Write-Host "Open your browser and login to dell after that copy / paste the session cookie here"
    $cookie = Read-InputFromConsole

    if ($null -ne $cookie)
    {
        ProcessCookie($cookie)
    }
}


#endregion


function Read-InputFromConsole ($MaxLength = 65536)
{
    [System.Console]::SetIn([System.IO.StreamReader]::new([System.Console]::OpenStandardInput($maxLength), [System.Console]::InputEncoding, $false, $maxLength))
    return [System.Console]::ReadLine()
}



#region Functions
function SearchAsset {
    param($ServiceTag, $AssetList)

    foreach ($asset in $AssetList.Assets)
    {
        $asset.ServiceTag
    }

}
#endregion

function DellPostRequest 
{
    param($url, $data)

    $headers = @{ 
        "Host" = "www.dell.com";
        "Accept" = "*/*";
        'Accept-Language' = 'en-US,en;q=0.5';
        'Accept-Encoding'='gzip, deflate';
        'Referer'='https://www.dell.com/support/mps/en-uk/myproducts';
        'Content-Type'='application/json; charset=utf-8';
        'Mps-Lock'='True';
        'X-Requested-With'='XMLHttpRequest';
        'Sec-Fetch-Dest'='empty';
        'Sec-Fetch-Mode'='cors';
        'Sec-Fetch-Site'='same-origin';
    }
	
	$retry_count = 0
	while ($retry_count -lt 3)
	{
        try 
        {
			if ($Script:DebugProxy)
			{
				$r = Invoke-WebRequest -Uri $url -WebSession $Script:WebSession -UserAgent $Script:UserAgent -Headers $headers -Method Post -Body $data -Proxy "http://127.0.0.1:8080" -TimeoutSec 10
			}
			elseif ($Script:UseSystemProxy)
			{
				$proxy = ([System.Net.WebRequest]::GetSystemWebproxy()).GetProxy($url)
				$r = Invoke-WebRequest -Uri $url -WebSession $Script:WebSession -UserAgent $Script:UserAgent -Headers $headers -Method Post -Body $data -Proxy $proxy -ProxyUseDefaultCredentials -TimeoutSec 10
			}
			else 
			{
				$r = Invoke-WebRequest -Uri $url -WebSession $Script:WebSession -UserAgent $Script:UserAgent -Headers $headers -Method Post -Body $data -TimeoutSec 10
			}

			if ($r.StatusCode -eq 200)
			{
				return $r
			}
        }
        catch [System.Net.WebException] {
            Write-Warning "DellPostRequest  WebException:" 
            Write-Warning $_
             $retry_count++
	    }
        catch {
            Write-Error "DellPostRequest  Error:" 
            Write-Error $_
            return $null
        }
	}
	
}

function ProcessAssetList 
{
    param ($assets)
    foreach ($asset in $assets.Assets)
    {
        if (![string]::IsNullOrEmpty($asset.ServiceTag))
        {   
            $productname = (IIf [string]::IsNullOrEmpty($asset.ProductName) "UNKNOWN" $asset.ProductName)
            $servicetag = (IIf [string]::IsNullOrEmpty($asset.ServiceTag) "UNKNOWN" $asset.ServiceTag)
            $shipdate = (IIf [string]::IsNullOrEmpty($asset.ShipDate) "UNKNOWN" $asset.ShipDate)
            $warrantytype = (IIf [string]::IsNullOrEmpty($asset.WarrantyType) "UNKNOWN" $asset.WarrantyType)
            $warrantyenddate = (IIf [string]::IsNullOrEmpty($asset.WarrantyEndDate) "UNKNOWN" $asset.WarrantyEndDate)

            $Script:AssetList.Add($asset.ServiceTag, [Asset]::new($productname, $servicetag, $shipdate, $warrantytype, $warrantyenddate))
        }
    }
}

function Save-DellAssetsToCSV
{
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Path
    )
    $assets = $Script:AssetList
    $assets.Values | Export-Csv -NoTypeInformation -Path $Path
}

function LoadDellAssets
{
    $Script:AssetList = @{} # Reset asset list
    $pagesize = 150
    # List Products    
    
    $init = $true
    $requests = 1
    $RequestPage = 1
    $progress = 0
    while ($true)
    {
        Write-Progress -Activity "Loading Assets from Dell" -Status "$progress% Complete:" -PercentComplete $progress
        $data = '{"GroupId":0,"GroupName":"","KeyWordSearch":"","ProductNames":[],"NickNames":[],"WarrantyTypes":[],"Paging":{"PageNumber":'+$RequestPage+',"PageSize":' + $pagesize + '},"SortBy":{"SortBy":"ProductName","SortOrder":"Asc"}}'
        $r = DellPostRequest $Script:Url_ListProducts $data 

        if ($null -ne $r)
        {
            $assets = ConvertFrom-Json $r.Content
            $total = $assets.Paging.TotalCount
            $size = $assets.Paging.PageSize
            $page = $assets.Paging.PageNumber

            ProcessAssetList $assets

            if ($init)
            {
                
                if ($total -le $size) { return $true }
                else 
                {
                    $requests = [Math]::Floor($total / $size)
                    if (($total % $size) -gt 0)
                    {
                        $requests += 1
                    }

                    $requests
                }
                $init = $false
            }
            elseif ($RequestPage -ge $requests) {
                Write-Progress -Activity "Loading Assets from Dell" -Status "$progress% Complete:" -PercentComplete $progress -Completed
                return $true
            }

        }
        else { break }
        $progress = [Math]::Floor((100/$requests) * $RequestPage)
        $RequestPage += 1
        Start-Sleep 0.5
    }
    return $false
}

#Remove Product
function RemoveDellProduct
{
    param(
        $ProductName,
        $ServiceTag
    )
    $data = '{"ProductName":"'+$ProductName+'","ServiceTag":"'+$ServiceTag+'","GroupId":0}'
    $r = DellPostRequest $Script:Url_RemoveProduct $data 

    if ($null -ne $r)
    {
        Write-Host "Successfully removed $ServiceTag from Dell"
    }
    return $false
}

#Add Product
function AddDellProduct
{
    param($ServiceTag)
    $data = '{"products":[{"ServiceTag":"'+$ServiceTag+'","AssignedTo":"","Groups":null}]}'
    $r = DellPostRequest $Script:Url_AddProduct $data 

    if ($null -ne $r)
    {
        $j = $r.Content | ConvertFrom-Json
        if (1 -eq $j.ResponseCode)
        {
            return $true
        }
    }

    return $false
}


function Clear-DellProducts
{
    param ($Backup = $true)
    $r = Read-Host "Are you sure you want to remove all your products from Dell? (Y/N)"
    if ($r.ToLower() -eq 'y')
    {
        $dtf = Get-Date -Format "yyyymmddHHmmss"
        Write-Host "Writing backup to backup-${dtf}.csv"

        Save-DellAssetsToCSV "backup-${dtf}.csv"

        $counter = 0
        $assets = $Script:AssetList
        foreach ($key in $assets.Keys)
        {
            $progress = [Math]::Floor((100 / $assets.Count) * $counter)
            Write-Progress -Activity "Removing Assets from Dell" -Status "$progress% Complete:" -PercentComplete $progress
            $productname = $assets[$key].ProductName
            $servicetag = $assets[$key].ServiceTag
            $r = RemoveDellProduct $productname $servicetag
            Write-Host $r.StatusCode
            $counter++
            Start-Sleep 0.5
        }
        Write-Progress -Activity "Removing Assets from Dell" -Status "100% Complete:" -PercentComplete 100 -Completed
        $Script:AssetList = @{}
    }
}

function LoadDellProductsFromFile 
{
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Path
    )
    if (Test-Path $Path) {
        $servicetags = Get-Content $Path

        $total = $servicetags.Count
        $counter = 0
        foreach ($servicetag in $servicetags)
        {
            $progress = [Math]::Floor((100 / $total) * $counter)
            Write-Progress -Activity "Adding Assets to Dell" -Status "$progress% ($counter/$total) Complete:" -PercentComplete $progress
            $counter++
            $servicetag = $servicetag.Trim()
            if ([string]::IsNullOrEmpty($servicetag)) { continue }
            if ($Script:AssetList.Contains($servicetag)) { continue }
            $r = AddDellProduct($servicetag)
            if ($false -eq $r) 
            {
                Write-Host "Failed to add servicetag: $servicetag"
            }
            
        }
        Write-Progress -Activity "Adding Assets to Dell" -Status "100% Complete:" -PercentComplete 100 -Completed
    }
    else 
    {
        Write-Error "Unable to find file $Path"
    }
}



function SaveDellAssetsMenuEntry
{
    $dtf = Get-Date -Format "yyyymmddHHmmss"
    $fname = "DellAssetInfo-${dtf}.csv"
    Write-Host "Saving currently loaded assets to file: ${fname}"
    Save-DellAssetsToCSV $fname

    Read-Host "Press enter to continue."
}

function UploadDellAssets
{
    $fname = "ServiceTags.txt"
    Write-Host "Upload Dell ServiceTags from file:"
    $r = Read-Host "Press enter to use the Default [${fname}]:"
    if (-not([string]::IsNullOrEmpty($r)))
    {
        $fname = $r
    }


    if (Test-Path $fname)
    {
        LoadDellProductsFromFile $fname
    }
    else {
        Write-Error "ERROR: $fname does not exit"
    }
    Read-Host "Press enter to continue."
}

function DellAssetMenu
{
    $exit = $false
    while (!$exit)
    {

        Clear-Host
        Write-Host "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
        Write-Host ""
        Write-Host "1. Set Session Cookie"
        Write-Host "2. Load assets from Dell"
        Write-Host "3. Upload assets to Dell"
        Write-Host "4. Save asset info to file"
        Write-Host ""
        Write-Host "8. Clear all assets from Dell"
        Write-Host "0. Exit"
        Write-Host "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
        $input = Read-Host "Please make your choice"
        switch ($input)
        {
            1 { Set-DellCookie }
            2 { LoadDellAssets ; Read-Host "Press enter to continue" }
            3 { UploadDellAssets }
            4 { SaveDellAssetsMenuEntry }
            8 { Clear-DellProducts }
            0 { exit }
            default { continue }
        }
    }
}

DellAssetMenu

