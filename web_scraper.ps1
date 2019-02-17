cd $PSScriptRoot

#http://www.wccnet.edu/search/employee/search/search/
Add-Type -AssemblyName System.Web

$ie = New-Object -com InternetExplorer.Application
$ie.silent = $true
$ie.navigate2("http://www.wccnet.edu/search/employee/search/search/")
while($ie.Busy) {Start-Sleep -Seconds 1}

$doc = $ie.Document
$departmentElements = $doc.getElementById("department")

$departments = @()
foreach($ele in $departmentElements) {
    if($ele.value) {
        $departments += $ele.value
    }
}

<#$filter = @(
    "Business & Computer Technologies",
    "Industrial Technology",
    "Math, Science & Engineering Technology"
    "Hum, Soc & Behav Sciences",
    "Health Sciences"
    "PR & Marketing"
    "Learning Resources"
)#>

$filter = @(
    "Administration & Finance",
    "Business & Computer Technologies",
    "Entrepreneurship Center",
    "Pooled Business Services - SBDC",
    "Small Business Development Ctr",
    "PR & Marketing"
)

#Loop through departments and gather emails
foreach($dep in $departments) {
    if($dep -in $filter) {
        Clear-Content -Path "./emails-${dep}.csv" -Force
        Write-Host ""
        Write-Host "Department: [${dep}]"
        $ie.Navigate2("http://www.wccnet.edu/search/employee/department/" + [System.Web.HttpUtility]::UrlEncode(${dep}) + "/search/search/")
        while($ie.Busy) {Start-Sleep -Seconds 1}
        $emailElements = $ie.Document.getElementsByClassName('more')

        foreach($detail in $emailElements) {
            $detail.previousSibling.previousSibling.previousSibling.previousSibling.click()
            while(($detail.getElementsByTagName('a') | Where {$_.protocol -eq 'mailto:'}).pathname -eq $null) {Start-Sleep -Seconds .1}
            $email = ($detail.getElementsByTagName('a') | Where {$_.protocol -eq 'mailto:'}).pathname
            Add-Content -Value $email -Path "./emails-${dep}.csv" -Force
            Write-Host $email  
        }
    }
}

$ie.Quit()