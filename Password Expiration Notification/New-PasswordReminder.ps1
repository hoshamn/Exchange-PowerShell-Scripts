<# 
.SYNOPSIS
		Notifies users that their password is about to expire.

.DESCRIPTION
    Let's users know their password will soon expire. Details the steps needed to change their password, and advises on what the password policy requires. Accounts for both standard Default Domain Policy based password policy and the fine grain password policy available in 2008 domains.

.NOTES
    Version    	      	: v2.9 - See changelog at http://www.ehloworld.com/596
    Wish list						: Set $DaysToWarn automatically based on Default Domain GPO setting
    										: Description for scheduled task
    										: Verify it's running on R2, as apparently only R2 has the AD commands?
    										: Determine password policy settings for FGPP users
    										: better logging
    Rights Required			: local admin on server it's running on
    Sched Task Req'd		: Yes - install mode will automatically create scheduled task
    Lync Version				: N/A
    Exchange Version		: 2007 or later
    Author       				: M. Ali (original AD query), Pat Richard, Lync MVP
    Email/Blog/Twitter	: pat@innervation.com 	http://www.ehloworld.com @patrichard
    Dedicated Post			: http://www.ehloworld.com/318
    Disclaimer   				: You running this script means you won't blame me if this breaks your stuff.
    Acknowledgements 		: (original) http://blogs.msdn.com/b/adpowershell/archive/2010/02/26/find-out-when-your-password-expires.aspx
    										: (date) http://technet.microsoft.com/en-us/library/ff730960.aspx
												:	(calculating time) http://blogs.msdn.com/b/powershell/archive/2007/02/24/time-till-we-land.aspx
												: http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/23fc5ffb-7cff-4c09-bf3e-2f94e2061f29/
												: http://blogs.msdn.com/b/adpowershell/archive/2010/02/26/find-out-when-your-password-expires.aspx
												: (password decryption) http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/f90bed75-475e-4f5f-94eb-60197efda6c6/
												: (determine per user fine grained password settings) http://technet.microsoft.com/en-us/library/ee617255.aspx
  	Assumptions					: ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
    Limitations					: 
    Known issues				: 												

.LINK     
    http://www.ehloworld.com/318

.INPUTS
		None. You cannot pipe objects to this script
		
.PARAMETER Demo
		Runs the script in demo mode. No emails are sent to the user(s), and onscreen output includes those who are expiring soon.

.PARAMETER Preview
		Sends a sample email to the user specified. Usefull for testing how the reminder email looks.
		
.PARAMETER PreviewUser
		User name of user to send the preview email message to.

.PARAMETER Install
		Create the scheduled task to run the script daily. It does NOT create the required Exchange receive connector.

.PARAMETER NoImages
		When specified, sends the email with no images, but keeps all other HTML formatting.
		
.EXAMPLE 
		.\New-PasswordReminder.ps1
		
		Description
		-----------
		Searches Active Directory for users who have passwords expiring soon, and emails them a reminder with instructions on how to change their password.

.EXAMPLE 
		.\New-PasswordReminder.ps1 -demo
		
		Description
		-----------
		Searches Active Directory for users who have passwords expiring soon, and lists those users on the screen, along with days till expiration and policy setting

.EXAMPLE 
		.\New-PasswordReminder.ps1 -Preview -PreviewUser [username]
		
		Description
		-----------
		Sends the HTML formatted email of the user specified via -PreviewUser. This is used to see what the HTML email will look like to the users.

.EXAMPLE 
		.
		
		Description
		-----------
		Creates the scheduled task for the script to run everyday at 6am. It will prompt for the password for the currently logged on user. It does NOT create the required Exchange receive connector.

#> 
#Requires -Version 2.0 

[CmdletBinding(SupportsShouldProcess = $True)]
param(
	[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)] 
	[switch]$Demo,
	[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)] 
	[switch]$Install,
	[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)] 
	[string]$PreviewUser,
	[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)] 
	[switch]$NoImages
)
Write-Verbose "Setting variables"
[string]$Company_en = "NOUR MAIL Service"
[string]$Company_ar = "خدمة نور ميل"
[string]$OwaUrl = "https://mail.nourmail.com.sa/owa"
[string]$PSEmailServer = "10.214.25.202"
[string]$EmailFrom = "notifications@nourmail.com.sa"
# Set the following to blank to exclude it from the emails
#[string]$HelpDeskPhone = "+966 11 821 6033   Ext: 111"
# Set the following to blank to remove the link from the emails
[string]$HelpDeskURL = ""
[string]$TranscriptFilename = $MyInvocation.MyCommand.Name + " " + $env:ComputerName + " {0:yyyy-MM-dd hh-mmtt}.log" -f (Get-Date)
[int]$global:UsersNotified = 0
[int]$DaysToWarn = 10
# Below path should be accessible by ALL users who may receive emails. This includes external/mobile users.
[string]$HImagePath = "https://sp8sm17mrcn3rstwre09xx1e-wpengine.netdna-ssl.com/wp-content/uploads/2018/12/TRSDevCo-Final-Identity-English-NoColor.png"
[string]$ScriptName = $MyInvocation.MyCommand.Name
[string]$ScriptPathAndName = $MyInvocation.MyCommand.Definition
[string]$ou = "DC=nourmail,DC=local"
# Change the following to alter the format of the date in the emails sent
# See http://technet.microsoft.com/en-us/library/ee692801.aspx for more info
[string]$DateFormat = "d"

if ($PreviewUser){
	$Preview = $true
}

Write-Verbose "Defining functions"
function Set-ModuleStatus { 
	[cmdletBinding(SupportsShouldProcess = $true)]
	param	(
		[parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true, HelpMessage = "No module name specified!")] 
		[string]$name
	)
	if(!(Get-Module -name "$name")) { 
		if(Get-Module -ListAvailable | ? {$_.name -eq "$name"}) { 
			Import-Module -Name "$name" 
			# module was imported
			return $true
		} else {
			# module was not available (Windows feature isn't installed)
			return $false
		}
	}else {
		# module was already imported
		return $true
	}
} # end function Set-ModuleStatus

function Remove-ScriptVariables {  
	[cmdletBinding(SupportsShouldProcess = $true)]
	param(
		[string]$path
	)
	$result = Get-Content $path |  
	ForEach { 
		if ( $_ -match '(\$.*?)\s*=') {      
			$matches[1]  | ? { $_ -notlike '*.*' -and $_ -notmatch 'result' -and $_ -notmatch 'env:'}  
		}  
	}  
	ForEach ($v in ($result | Sort-Object | Get-Unique)){		
		Remove-Variable ($v.replace("$","")) -EA 0
	}
} # end function Remove-ScriptVariables

function Install	{
	[cmdletBinding(SupportsShouldProcess = $true)]
	param()
	# http://technet.microsoft.com/en-us/library/cc725744(WS.10).aspx
	$error.clear()
	Write-Host "Creating scheduled task `"$ScriptName`"..."
	$TaskCreds = Get-Credential("$env:userdnsdomain\$env:username")
	$TaskPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($TaskCreds.Password))
	schtasks /create /tn $ScriptName /tr "$env:windir\system32\windowspowershell\v1.0\powershell.exe -command $ScriptPathAndName" /sc Daily /st 09:21 /ru $TaskCreds.UserName /rp $TaskPassword | Out-Null
	if (! $error){
		Write-Host "Installation complete!" -ForegroundColor green
	}else{
		Write-Host "Installation failed!" -ForegroundColor red
	}
	remove-variable taskpassword
	exit
} # end function Install

function Get-ADUserPasswordExpirationDate {
	[cmdletBinding(SupportsShouldProcess = $true)]
	Param (
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, HelpMessage = "Identity of the Account")]
		[Object]$accountIdentity
	)
	PROCESS {
		Write-Verbose "Getting the user info for $accountIdentity"
		$accountObj = Get-ADUser $accountIdentity -properties PasswordExpired, PasswordNeverExpires, PasswordLastSet, name, mail , DisplayName, Description
		# Make sure the password is not expired, and the account is not set to never expire
    Write-Verbose "verifying that the password is not expired, and the user is not set to PasswordNeverExpires"
    if (((!($accountObj.PasswordExpired)) -and (!($accountObj.PasswordNeverExpires))) -or ($PreviewUser)) {
    	Write-Verbose "Verifying if the date the password was last set is available"
    	$passwordSetDate = $accountObj.PasswordLastSet     	
      if ($passwordSetDate -ne $null) {
      	$maxPasswordAgeTimeSpan = $null
        # see if we're at Windows2008 domain functional level, which supports granular password policies
        Write-Verbose "Determining domain functional level"
        if ($global:dfl -ge 4) { # 2008 Domain functional level
          $accountFGPP = Get-ADUserResultantPasswordPolicy $accountObj
          if ($accountFGPP -ne $null) {
          	$maxPasswordAgeTimeSpan = $accountFGPP.MaxPasswordAge
					} else {
						$maxPasswordAgeTimeSpan = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
					}
				} else { # 2003 or ealier Domain Functional Level
					$maxPasswordAgeTimeSpan = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
				}				
				if ($maxPasswordAgeTimeSpan -eq $null -or $maxPasswordAgeTimeSpan.TotalMilliseconds -ne 0) {
					$DaysTillExpire = [math]::round(((New-TimeSpan -Start (Get-Date) -End ($passwordSetDate + $maxPasswordAgeTimeSpan)).TotalDays),0)
					if ($preview){$DaysTillExpire = 10}
					if ($DaysTillExpire -le $DaysToWarn){
						Write-Verbose "User should receive email"
						$PolicyDays = [math]::round((($maxPasswordAgeTimeSpan).TotalDays),0)
						if ($demo)	{Write-Host ("{0,-25}{1,-8}{2,-12}" -f $accountObj.Name, $DaysTillExpire, $PolicyDays)}
            # start assembling email to user here
						$EmailName = $accountObj.DisplayName
						$ArName	= $accountObj.description						
						$DateofExpiration = (Get-Date).AddDays($DaysTillExpire)
						$DateofExpiration = (Get-Date($DateofExpiration) -f $DateFormat)						

Write-Verbose "Assembling email message"						
[string]$emailbody = @"






<html>

<head>
	<meta charset="UTF-8">
	<style>
		html {
			high: 100%;
		}
		
		body {
			padding-left: 9%;
			padding-right: 9%;
			padding-top: 1%;
			padding-bottom: 1%;
		}
		
		h3 {
			font-size: 15px;
			font-family: Tahoma, 'Segoe UI', 'Segoe WP', 'Segoe UI WPC', Arial, sans-serif;
		}
		
		p {
			font-size: 14px;
			font-family: Tahoma, 'Segoe UI', 'Segoe WP', 'Segoe UI WPC', Arial, sans-serif;
			text-align: justify;
			padding: 5px 0px 5px 0;
		}
		
		li {
			font-size: 14px;
			font-family: Tahoma, 'Segoe UI', 'Segoe WP', 'Segoe UI WPC', Arial, sans-serif;
			text-align: justify;
			padding: 5px 0px 0 0;
		}
		
		#ar_content h2 {
			font-family: 'Segoe UI', 'Segoe WP', 'Segoe UI WPC', Tahoma, Arial, sans-serif;
		}
		
		#ar_content {
			direction: rtl; 
			height:100%;
		}
		
		#en_content {
            direction: ltr; 
			height:100%;  
		}
		
		#headpic  .left {
			float:center;
            		text-align:center;
		}
		
		/*#headpic  .right {
			 
			 float:right;
            text-align:right;
		}*/
		
		#footer {
			width: 100%;
			background: #16894E;
			height: 85px;
			margin: auto;
			position: relative;
			font-family: 'Segoe UI', 'Segoe WP', 'Segoe UI WPC', Tahoma, Arial, sans-serif;
			position: bottom;
		}

			
		.container .ftContent {
			 
			padding: 10px;
			margin-top: -30px;
			direction: ltr;
		}
		
		.one {
		 
			text-align: right;
			direction: rtl;
			 
		}
		
		.two {
		 
			text-align: left;
			direction: ltr;
			  
		}

		.three {
		 
			text-align: right;
			direction: rtl;
			 
		}
		
		.four {
		 
			text-align: left;
			direction: ltr;
			  
		}
		
		
		#ar_content ol {
			padding-right: 1em;
		}
		
		#ar_content ol>li {
			padding-right: 1em;
		}
		
		#en_content ol {
			padding-left: 1em;
		}
		
		#en_content ol>li {
			padding-left: 1em;
		}
		
		.title {
			clear: both;
			text-align: center;
			color: #16894e;
			font-weight: bold;
			font-size: 22px;
			margin-bottom: 0px;
			padding: 0px;
		}
		
		.header {
			padding: 10px;
		}
		
		 
		
		.tblContent {
			width: 100%;
			direction: rtl;
		}
		.ftContent {
			width:100%;
			direction: rtl;
		}
		
	</style>

	<!--[if (gte IE 6)&(lte IE 8)]>
	<style>
		.footer {
			 
			margin-top:0px;
		}
	</style>
	<![endif]-->
</head>

<body>

	
<div id="container" class="container">

		<table class="tblContent">
			<tr>
				<td style='width:48%'  >

					<div id="ar_content" class="one">

						<p>مرحبا $EmailName,</p>
						<p> تحية طيبة من فريق نور ميل,</p>
						<p> بناءً على سياسات وإجراءات أمن المعلومات في إحكام, وحرصاً منا على استمرارية الخدمة بأعلى مقاييس الأمان. فإنه يجب على المستخدم القيام بتغيير كلمة المرور كل 90 يوم وذلك تفادياَ لحدوث أي اختراق على البريد الالكتروني.</p>	
							
							<p> ونود تذكيركم بأن كلمة المرور الحالية الخاصة بكم سوف تنتهي صلاحيتها بعد <span style="background-color: red; color: white; font-weight: bold;">&nbsp;$DaysTillExpire&nbsp;</span> يوم، في <span style="font-weight: bold;">&nbsp;$DateofExpiration&nbsp;</span>.</p>
							<p>الرجاء تغيير كلمة المرور لتجنب الإيقاف المؤقت لحسابك. يمكن تغيير كلمة المرور من خلال اتباع الخطوات التالية::</p>

						<ol>
							
							

<li> 	الدخول على بوابة البريد الالكتروني من خلال الرابط التالي خدمة <a href="https://cp.nourmail.com.sa/ForgetPassword.aspx">نور ميل للبريد الإلكتروني</a>	</li>
<li> 	ادخل اسم المستخدم كاملا ثم أدخل كلمة المرور الحالية.</li>
<li> 	اضغط على علامة "الإعدادات".</li>
<li> 	الضغط على " خيارات".</li>
<li> 	أختيار " عام".</li>
<li> 	أختيار "حسابي".</li>
<li> 	اضغط على "تغيير كلمة المرور" في اسفل الصفحه.</li>
<li> 	ادخال كلمة المرور الحالية.</li>
<li> 	ادخال كلمة المرور الجديدة مرتين.</li>
<li>	اضغط على حفظ. </li>
</ol>
                        
                            <p><span style="font-weight: bold">ملاحظة:</span> ستحتاج إلى استخدام كلمة المرور الجديدة الخاصة بك عند تسجيل الدخول إلى حسابات البريد الإلكتروني على برنامج Microsoft Outlook أو جهاز الكمبيوتر اللوحي أو الهاتف الجوال. </p>

						<h3>القواعد الواجب اتباعها عند اختيار كلمة المرور الخاصة بك في الشركة: $company_ar :-</h3>


						<ol>
							<li>يجب ان لا تقل كلمة المرور عن $MinPasswordLength.</li>
							<li>يجب ان لا تحتوي كلمة المرور على اسمك الاول او الاخير او اسم المستخدم الخاص بك.</li>
							<li>يجب تغيير كلمة المرور كل $PolicyDays يوم.</li>
							<li>لا يمكن إعادة استخدام أي من كلمات المرور $PasswordHistory الأخيرة.</li>
							<li>لا يمكن تغيير كلمة المرور أكثر من مرة واحدة يوميا.</li>
							<li>يجب ان تحقق كلمة المرور على الأقل ثلاث من الشروط التالية:-</li>
							<ul><b>
								<li>1 حرف كبير (A-Z)</li>
								<li>1 حرف صغير (a-z)</li>
								<li>1 رقم (0-9)</li>
								<li>1 رمز خاص ($,@,#,..)</li>

							</b></ul>

						</ol>

						

					</div>
				</td>
                <td style='width:4%'>
                &nbsp; 
                </td>
				<td  id="" style='width:48%'>
					<div id="en_content" class="two">

						<p>Hello $EmailName, </p>
						<p>Greetings from Nourmail Team, </p> 
				<p>Based on information security policies and procedures in Nourmail, and in order to maintain the service on its highest security standards. The user is required to change the password every 90 days in order to avoid security breach on the email account. </p> 
                <p>We would like to remind you that your password will expire in  <span style="background-color: red; color: white; font-weight: bold;">&nbsp;$DaysTillExpire&nbsp;</span> day(s), on <span style="font-weight: bold;">&nbsp;$DateofExpiration</span>.</p>
				<p>Kindly change your password to avoid the temporarily suspension of your account. You can change your password by following these steps :</p>

						<ol>
							
							<li>Log into <a href="https://cp.nourmail.com.sa/ForgetPassword.aspx">Nourmail Service URL</a></li>
							<li>Enter your full email ID and your current password.</li>
							<li>Click "Settings" sign.</li>
							<li>Select the "Option".</li>
							<li>Select the " General”.</li>
							<li>Select the " My Account”.</li>
							<li>At the bottom of the page select “Change Your Password</li>
							<li>Enter your current password.</li>
							<li>Then enter your new password twice.</li>
							<li>After that click "Save".</li>
							
						</ol>

                           <p><span style="font-weight: bold">NOTE:</span>You will need to use your new password when logging into Microsoft Outlook, Tablet or Cellphone Email Clients.</p>

						<h3>$company_en Password Policy:</h3>

						<ol>
							<li>password must have a minimum of a $MinPasswordLength characters.</li>
							<li>password must not contain parts of your first, last, or logon name.</li>
							<li>password must be changed every $PolicyDays days.</li>
							<li>You cannot reuse any of your last $PasswordHistory passwords</li>
							<li>Your password cannot be changed more than one time per day.</li>
							<li>Password must contain at least three of the following conditions:-</li>

							<ul><b>
								<li>1 upper case character (A-Z)</li>
								<li>1 lower case character (a-z)</li>
								<li>1 numeric character (0-9)</li>
								<li>1 special character ($,@,#,..)</li>
							</b></ul>
						</ol>

						<!-- <h3>Best Regards,<br /> Nourmail Team<br /> -->
						 
					</div>
				</td>
			</tr>

		<tr>
                
		<td style='width:48%' >
		<div id="ft_right" class="three">                    
			#<img   src=" ">
                </td>
		</div>

		<td style='width:4%'>
                &nbsp; 
                </td>

		<td style='width:48%'>
		<div id="ft_left" class="four">
          #          <img   src="">
                </td>
		</div>
		     </tr>


		</table>


						
"@
if ($accountFGPP -eq $null){ 
	$emailbody += @"
			
"@							

if ($PasswordComplexity){
	Write-Verbose "Password complexity"
	$emailbody += @"
							
"@
}
$emailbody += @"
						
"@
}
if (!($NoImages)){
$emailbody += @"								
							
"@
}
if ($HelpDeskURL){
$emailbody += @"															
							
"@
}
if (!($NoImages)){
$emailbody += @"
								
"@
}
$emailbody += @"
	</body>
</html>
"@
						if (!($demo)){
							$emailto = $accountObj.mail
							if ($emailto){
								Write-Verbose "Sending demo message to $emailto"
								Send-MailMessage -To $emailto -Subject "ALERT: Your password expires in $DaysTillExpire day(s)  |  تنبيه: كلمة المرور الخاصه بك تنتهي خلال $DaysTillExpire يوم" -Body $emailbody -From $EmailFrom -Priority High -BodyAsHtml -Encoding utf8
								$global:UsersNotified++
							}else{
								Write-Verbose "Can not email this user. Email address is blank"
							}
						}
					}
				}
			}
		}
	}
} # end function Get-ADUserPasswordExpirationDate

if ($install){
	Write-Verbose "Install mode"
	Install
	Exit
}

Write-Verbose "Checking for ActiveDirectory module"
if ((Set-ModuleStatus ActiveDirectory) -eq $false){
	$error.clear()
	Write-Host "Installing the Active Directory module..." -ForegroundColor yellow
	Set-ModuleStatus ServerManager
	Add-WindowsFeature RSAT-AD-PowerShell
	if ($error){
		Write-Host "Active Directory module could not be installed. Exiting..." -ForegroundColor red; 
		if ($transcript){Stop-Transcript}
		exit
	}
}
Write-Verbose "Getting Domain functional level"
$global:dfl = (Get-AdDomain).DomainMode
#Get-ADUser -filter * -properties PasswordLastSet,EmailAddress,GivenName -SearchBase "DC=nourmail,DC=local" |foreach {
if (!($PreviewUser)){
	if ($ou){
		Write-Verbose "Filtering users to $ou"
		# $users = Get-AdUser -filter * -SearchScope subtree -SearchBase $ou -ResultSetSize $null
        $users = Get-AdUser -ldapfilter '(!(name=*$))' -SearchScope subtree -SearchBase $ou -ResultSetSize $null
	}else{
		# $users = Get-AdUser -filter * -ResultSetSize $null
		$users = Get-AdUser -ldapfilter '(!(name=*$))' -ResultSetSize $null
	}
}else{
	Write-Verbose "Preview mode"
	$users = Get-AdUser $PreviewUser
}
if ($demo){
	Write-Verbose "Demo mode"
	# $WhatIfPreference = $true
	Write-Host "`n"
	Write-Host ("{0,-25}{1,-8}{2,-12}" -f "User", "Expires", "Policy") -ForegroundColor cyan
	Write-Host ("{0,-25}{1,-8}{2,-12}" -f "========================", "=======", "===========") -ForegroundColor cyan
}

Write-Verbose "Setting event log configuration"
[object]$evt = new-object System.Diagnostics.EventLog("Application")
[string]$evt.Source = $ScriptName
$infoevent = [System.Diagnostics.EventLogEntryType]::Information
[string]$EventLogText = "Beginning processing"
# $evt.WriteEntry($EventLogText,$infoevent,70)

Write-Verbose "Getting password policy configuration"
$DefaultDomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy
[int]$MinPasswordLength = $DefaultDomainPasswordPolicy.MinPasswordLength
# this needs to look for FGPP, and then default to this if it doesn't exist
[bool]$PasswordComplexity = $DefaultDomainPasswordPolicy.ComplexityEnabled
[int]$PasswordHistory = $DefaultDomainPasswordPolicy.PasswordHistoryCount

ForEach ($user in $users){
If ($user.Enabled)
{
	Get-ADUserPasswordExpirationDate $user.samaccountname
}
}

Write-Verbose "Writing summary event log entry"
$EventLogText = "Finished processing $global:UsersNotified account(s). `n`nFor more information about this script, run Get-Help .\$ScriptName. See the blog post at http://www.ehloworld.com/318."
#$evt.WriteEntry($EventLogText,$infoevent,70)

# $WhatIfPreference = $false

Remove-ScriptVariables -path $ScriptPathAndName