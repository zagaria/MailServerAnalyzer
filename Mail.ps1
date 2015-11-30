#requires -version 2.0
#Script parametr
param([string]$File="", [string]$Bulk = "40",[string]$Seconds = "1")

Write-Host "-------------------------------------------------------------------" -foregroundcolor "green"
Write-Host "                      Mail Servers Analyzer 1.0                    " -foregroundcolor "green"
Write-Host "-------------------------------------------------------------------" -foregroundcolor "green"

#Path for script dir
$DIR = $pwd.Path

#INIT
$Start = 0
$CACHE = @{""=""}
$EmailRegex = '^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$';

#For silent mode, if not use you see all messages (include info and errors all native functions) 
$ErrorActionPreference = 'silentlycontinue' 

#Set Mail Server Type
$MailServerType = @{"hotmail.com"="Outlook.com";
                    "google.com"="Gmail";
                    "eo.outlook.com"="Office365";
                    "microsoftonline.com"="Office365";
                    "autodiscover"="Office365";
                    "outlook.com"="Office365";
                    "onmicrosoft.com"="Office365";
                    "mx.yandex.ru"="Yandex";
                    "mxs.mail.ru"="Mail.Ru";
                    "mail.ru"="Mail.Ru";
                    "yahoo.com"="Yahoo";
                    "yahoodns.net"="Yahoo";
                    "googlemail.com"="Gmail"}

#Init STAT
$Statictics = @{"Outlook.com"=0;
                "Gmail"=0;
                "Office365"=0;
                "Yandex"=0;
                "Yahoo"=0;
                "Other"=0;
                "Error"=0;}


if ($file -eq "")
{
    Write-Host "----------------------------Help-----------------------------------" -foregroundcolor "green"
	Get-Content "$DIR\README.txt"
	Write-Host "----------------------------END-----------------------------------" -foregroundcolor "green"
	exit
}

Write-Host "Run Script. Check parametrs..."  -foregroundcolor "green"

#Prepare to start
function Init()
{
	$Start = CheckLast
	LoadCache
    ReadData $File $Start
}

<#

#>
function CheckLast()
{
	$LastPosition = Get-Content $DIR"/config.txt"
	if ($temp -eq "")
    {
		$Start = 0
	}
	else
	{
		$Start = $LastPosition
	}
	return [int]$Start
}

function LoadCache()
{
	#Load CACHE
	$temp = Get-Content $DIR"/cache.txt"
	$load = New-Object System.Collections.ArrayList
	$load.AddRange($temp)
	Good "Wait. Load cache"
	#Write-Host $load
	$count = $load.Count
	for ($i = 0;$i -le $count;$i++)
	{
		
		$split = $load[$i].Split('=')
		
		if ($split.length -ge 0)
		{
			
			if ($CACHE.Contains($temp[0]))
			{
				continue;
			}
			else
			{
				#Write-Host $split[0]
				$CACHE.Add($split[0],$split[1])
			}
		}
	}
	Good "Load $count complete."
}

function Statictics($data)
{
	if ($Statictics.Contains($data))
	{
		$Statictics.$data +=1;
	}
	else
	{
		$Statictics.Add($data,1);
	}
	#ReWrite Statictics after Bulk
	New-Item "$DIR\stat.txt" -type file -force -value "" >$null
	$Statictics.getEnumerator() | Select name, value | Add-Content "stat.txt"
}

<#
Write in console text with green color
#>
function Good ($Text,$Color="green")
{
	Write-Host $Text -foregroundcolor $Color
}

<#
Write in console text with red color
#>
function Error ($Text,$Color="red")
{
	Write-Host $Text -foregroundcolor $Color
}

<#
Detect Mail Server based on MX Records, Domain and Domain AutoDiscover
#>
function GetMailServerType($mxrecord,$domain)
{
   foreach ($ServerType in $MailServerType.Keys)
   {
      if ($mxrecord.ToLower().Contains($ServerType))
	  {
		return $MailServerType.Get_Item($ServerType)
	  }
   }
   foreach ($ServerType in $MailServerType.Keys)
   {
      if ($domain.Contains($ServerType))
	  {
        return $MailServerType.Get_Item($ServerType)
	  }
   }
   $error.Clear()
   
   # Add to domain 'autodiscover.' for detect Office365 Servers, check if exist this record in dns
   $domain = "autodiscover."+$domain
   
   #Get CNAME DNS Record
   $resolve = Resolve-DnsName -Name $domain -Type CNAME
   
   #If exist then
   if ($resolve.Name -ne "" -And $error.Count -eq "0")
   {
      return "Maybe Office365"
   }
   #Else
   return "Other"
}

<#
 Analyzing DNS Records (MX AND CNAME) and analyze domain name
#>
function AnalyzingDNSResult($DNSResult,$RealDomain,$PreviousResult='')
{
  if ($DNSResult -eq ""){Continue}
  
  $MailServerType = GetMailServerType $DNSResult $RealDomain
  
  if ($MailServerType -eq "Other") 
  {
  #Try detect based DNS for autodiscover.$RealDomain 
	$AutoDiscoverDomain = "autodiscover."+$RealDomain
	$DNSResult = DetectMail $AutoDiscoverDomain
	$MailServerType = GetMailServerType $DNSResult $RealDomain
  }
  
  If ($MailServerType -eq $PreviousResult){return $PreviousResult}
  
  $DataToBaseFile = $Email+","+$DNSResult+","+$MailServerType 
  WriteToFile $DataToBaseFile
  Statictics($MailServerType)
  
  return $MailServerType
}

<#
 Analyzing Email Based DNS
#>
function Analyzing ($Email)
{
    #Use RegEx for check correct Email Address
     if ($Email -NotMatch $EmailRegex) {
        continue
	}
     $GetRealDomain = $Email.Split('@')
     if ($GetRealDomain.Count -eq 2)
	 {
		   #Detect mail server type
		   $RealDomain = $GetRealDomain[1]
		   #Empty mxrecord for check based on domain
		   $Type = GetMailServerType '' $RealDomain
		   if ($Type -ne "Other")
		   {
				$ResolveDNSResult = Resolve-DnsName -Name $realdomain -Type MX
				Statictics($Type)
				$DataToBaseFile = $Email+","+$ResolveDNSResult.NameExchange+","+$Type 
				WriteToFile $DataToBaseFile
		        Continue;
		   }
		   #If first detect (based only domain) was false - trying detect based on DNS records MX (or CNAME) 
		   $DNSResult = DetectMail $RealDomain
		   $NumberElementsInDNSResult = $DNSResult.Count-1
		   if ($DNSResult -isnot [system.array])#IF return data is not an Array Type
		   {
			    AnalyzingDNSResult $DNSResult $RealDomain
		   }
		   else
		   {
		        $PreviousResult = ''
		        for ($y = 0;$y -le $NumberElementsInDNSResult;$y++)
				{
		           $PreviousResult = AnalyzingDNSResult $DNSResult[$y] $RealDomain $PreviousResult
				}
		   }
   }

}
<#
 Set progress and set request delay
#>
function Progress($CurrentPosition,$LastPosition)
{
        #For very security dns server use Sleep
        if ($CurrentPosition % $Bulk -eq 0 -and $CurrentPosition -ne 0)
		{
		  $Percent = [math]::Round($CurrentPosition/$LastPosition,4)*100
		  Write-Host "Analyzing: $CurrentPosition from $LastPosition ($Percent %)" 
		  Write-Host "Waiting $Seconds seconds..."  -foregroundcolor "green"
		  New-Item "$DIR\config.txt" -type file -force -value $i >$null
		  $Statictics
		  Sleep ($Seconds)
		}

}

<#
---------------------------------------
This function read email from text file
---------------------------------------

#>
function ReadData($FileName="data.txt",$StartPosition=0)
{
   $Emails = @(Get-Content $FileName)#Read all data from file
   $NumberEmails = $Emails.Count-1
   Write-Host "Found $NumberEmails emails. Start from position $Start"
   for ($i=$StartPosition;$i -le $NumberEmails;$i++)#Parse on one record
   {
        Progress $i $NumberEmails
		Analyzing $Emails[$i]
   }
}


<#
------------------------------
This function search MX record
------------------------------

Result: Return string or Array

#>
function DetectMail($domain)
{
    $Date = Get-Date
	$data = "";
		if ($CACHE.Contains($domain) -And $domain.Contains('autodiscover') -eq "False")
		{
			return $CACHE.Get_Item($domain)
		}
	Try
	{
		if ($domain.Contains('autodiscover'))
		{
		    $all = Resolve-DnsName -Name $domain -Type CNAME
		    $data = $all.NameHost
		}
		else
		{
			$all = Resolve-DnsName -Name $domain -Type MX
			$data = $all.NameExchange
		}
	}
	Catch
	{
	    #If error, write to log
        WriteToFile $Date" >> Error MX "$domain $DIR"/log.txt"
		return "Error"
	}
	Finally
	{
		#DEBUG
		foreach ($value in $data)
		{
		  #$result = GetMailServerType($value)
		  $text = $domain+"="+$value+"`r"
		  WriteToFile $text $DIR"/cache.txt"
		  $CACHE.Add($domain,$value)
		  #Write-Host $result
		}
	}
	WriteToFile $Date" >> Success MX "$domain" MX:"$data $DIR"/log.txt"
	return $data
}


#Write To File
function WriteToFile($Text,$PathToFile="$DIR\Base.csv")
{
#If file exist than
If (Test-Path $PathToFile)
{
	$TextBytes = [Text.Encoding]::GetEncoding("windows-1251").GetBytes($Text+"`r")
	$fs = New-Object IO.FileStream($PathToFile,[IO.FileMode]::Open,[Security.AccessControl.FileSystemRights]::AppendData,[IO.FileShare]::Read,8,[IO.FileOptions]::None)
	$fs.Write($TextBytes,0,$TextBytes.Count)
	$fs.Close()
}
else
{
    #If file not exist and filename contains "base.csv" - create file and write Field Name
    if ($PathToFile.Contains("Base.csv"))
	{
		$Text = "Email,MX Record,Server`r$text"
	}
	New-Item $PathToFile -type file -force -value $Text"`r" | Out-Null
}
}
#Start Script
Init
