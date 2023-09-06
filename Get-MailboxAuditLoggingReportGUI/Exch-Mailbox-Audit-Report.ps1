########################################################################
# Created By  : Tzahi Kolber
# Version     : v1.0.0.0
# Last Update : 21/01/2020 20:25
########################################################################
# Disclaimer:
# The sample scripts are not supported under any Microsoft standard support program or service.
# The sample scripts are provided AS IS without warranty of any kind.
# Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.
########################################################################


### Path location for saving the Reports ###

#$PathLocation = "C:\Scripts"

$PathLocation = [Environment]::GetFolderPath("Desktop")

##########################################


#Generated Form Function
function GenerateForm {

#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
#endregion


#region Generated Form Objects
$tooltip1 = New-Object System.Windows.Forms.ToolTip
$form = New-Object System.Windows.Forms.Form
$ReportLBL = New-Object System.Windows.Forms.Label
$Button1 = New-Object System.Windows.Forms.Button
$CBOwner = New-Object System.Windows.Forms.CheckBox
$CBAdmin = New-Object System.Windows.Forms.CheckBox
$CBDelegate = New-Object System.Windows.Forms.CheckBox
$MBXlistBox = New-Object System.Windows.Forms.ListBox
$SelMBXLBL = New-Object System.Windows.Forms.Label
$StdateLBL = New-Object System.Windows.Forms.Label
$ENDdateLBL = New-Object System.Windows.Forms.Label
$dateTimePicker2 = New-Object System.Windows.Forms.DateTimePicker
$dateTimePicker1 = New-Object System.Windows.Forms.DateTimePicker
$textBox = New-Object System.Windows.Forms.TextBox
$LBLPATH = New-Object System.Windows.Forms.Label
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

#Provide Custom Code for events specified in PrimalForms.
$handler_form_Load= 
{
#TODO: Place custom script here

}
# For Exchange Online usage, remove the # key before line 57 and add # key before line 58
#$mbx=(Get-mailbox -ResultSize unlimited  | ? {$_.AuditEnabled -eq "True"}).alias
$mbx=(Get-mailbox -ResultSize unlimited -Filter {AuditEnabled -eq "True"}).alias

$Button1_OnClick= 
{

#TODO: Place custom script here


$mbx | % {$textbox.AutoCompleteCustomSource.AddRange($_) }

    $Button1.Enabled = $false
    $Button1.Text = 'Running...'
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor

       If ($CBOwner.Checked -eq $true) {$o = "Owner"}
       If ($CBAdmin.Checked -eq $true) {$b = "Admin"}
       If ($CBDelegate.Checked -eq $true) {$d = "Delegate"}

$mbxch = $textBox.Text
If ($mbx -contains $mbxch) {

$StartDT=$dateTimePicker1.Value.ToString("MM/dd/yyyy")
$EndDT=$dateTimePicker2.Value.ToString("MM/dd/yyyy")

Write-Host "Mailbox for Auditing report is $mbxch"
Write-Host "Start Date: $StartDT"
Write-Host "End Date: $EndDT"

$logonTp=$o+$d+$b

switch ($logonTp)
    {
        "OwnerAdminDelegate" {$logont="Owner,Admin,Delegate"}
        "Owner" {$logonT="Owner"}
        "Admin" {$logonT="Admin"}
        "Delegate" {$logonT="Delegate"}
        "OwnerAdmin" {$logonT="Owner,Admin"}
        "AdminOwner" {$logonT="Owner,Admin"}
        "AdminDelegate" {$logonT="Admin,Delegate"}
        "DelegateAdmin" {$logonT="Admin,Delegate"}
        "OwnerDelegate" {$logonT="Owner,Delegate"}
        "DelegateOwner" {$logonT="Owner,Delegate"}
        "$null"      {$logont="Owner,Admin,Delegate" ; [System.Windows.Forms.MessageBox]::Show('Since you did not select a report type, a report with all Audit types will be generated. Click OK to continue') ; $Button1.Enabled = $false}
    }

Write-Host -ForegroundColor yellow $logonT

$AuditInfo = Search-MailboxAuditLog $mbxch -LogonTypes $logonT -StartDate $StartDT -EndDate $EndDT -ShowDetails -resultsize 25000

if ($($AuditInfo.Count) -gt 0)
{
 
    $AuditInfo | Export-CSV "$PathLocation\Audit-Report.csv" -NoTypeInformation -Encoding UTF8
 #   $AuditInfo | Out-GridView
    $report = @()
    foreach ($usr in $AuditInfo)
    {
	$int = ((($usr.ClientInfoString).split(";") | select -First 1).split("=") | select -Last 1)         
	if ($int -eq "MSExchangeRPC") {$int = "Outlook"}
        $Objrep = New-Object PSObject
        $Objrep | Add-Member NoteProperty -Name "Mailbox" -Value $usr.MailboxResolvedOwnerName
        $Objrep | Add-Member NoteProperty -Name "Mailbox UPN" -Value $usr.MailboxOwnerUPN
        $Objrep | Add-Member NoteProperty -Name "Time-stamp" -Value $usr.LastAccessed
        $Objrep | Add-Member NoteProperty -Name "LogonType" -Value $usr.LogonType
        $Objrep | Add-Member NoteProperty -Name "Accessed By" -Value $usr.LogonUserDisplayName
        $Objrep | Add-Member NoteProperty -Name "Operation" -Value $usr.Operation
        $Objrep | Add-Member NoteProperty -Name "Result" -Value $usr.OperationResult
        $Objrep | Add-Member NoteProperty -Name "Folder" -Value $usr.FolderPathName
        $Objrep | Add-Member NoteProperty -Name "Client IP Address" -Value $usr.ClientIPAddress
        $Objrep | Add-Member NoteProperty -Name "Client Interface" -Value $int 
        if ($usr.ItemSubject)
        {
            $Objrep | Add-Member NoteProperty -Name "Subject" -Value $usr.ItemSubject
        }
        else
        {
            $Objrep | Add-Member NoteProperty -Name "Subject" -Value $usr.SourceItemSubjectsList
        }

        $report += $Objrep
    }
        
    $htmlbody = $report | ConvertTo-Html -Fragment

	$htmlhead="<html>
				<style>
				BODY{font-family: Arial; font-size: 11pt;}
				H1{font-size: 24px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H2{font-size: 20px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H3{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				table {border-collapse: separate; background-color: #F2F2F2; border: 3px solid #103E69; caption-side: bottom;font-size: 8pt;}
				TH{border: 1px solid #969595; background: #33c1ff; padding: 5px; color: #000000;}
				TD{border: 1px solid #969595; padding: 5px; }
				td.pass{background: #B7EB83;}
				td.warn{background: #FFF275;}
				td.fail{background: #FF2626; color: #ffffff;}
				td.info{background: #85D4FF;}
				</style>
				<body>
                <p><b>Report of mailbox audit log for $mbxch between $StartDT and $EndDT.</p>"
		
	$htmltail = "</body></html>"	

	$htmlreport = $htmlhead + $htmlbody + $htmltail

    $htmlreport | Out-File "$PathLocation\Audit-Mailbox-Report.html" -Encoding UTF8
    $ReportLBL.ForeColor = "Blue"
    $ReportLBL.Text = "Report was generated"

    }

    Else
     {
     Write-Host -ForegroundColor Red "No Auditing data was found for $mbxch"
     $ReportLBL.ForeColor = "Red"
     $ReportLBL.Text = "No report was generated" 
     }
       }
       Else {[System.Windows.Forms.MessageBox]::Show("The mailbox doesn't exists or auditing was not activated - Please choose a valid mailbox with Auditing enabled")}
        $Button1.Text = 'Generate'
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
    $Button1.Enabled = $true
        }
    
$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$form.WindowState = $InitialFormWindowState
}

#----------------------------------------------
#region Generated Form Code
$System_Drawing_Size = New-Object System.Drawing.Size
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$System_Drawing_Size.Height = 442
$System_Drawing_Size.Width = 525
$form.ClientSize = $System_Drawing_Size
$form.DataBindings.DefaultDataSourceUpdateMode = 0
$form.Name = "form"
$form.Text = "Mailbox Audit Reporting ver 1.0 By Tzahi Kolber"
$form.add_Load($handler_form_Load)

$ReportLBL.DataBindings.DefaultDataSourceUpdateMode = 0
$ReportLBL.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8,1,3,0)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 240
$System_Drawing_Point.Y = 355
$ReportLBL.Location = $System_Drawing_Point
$ReportLBL.Name = "ReportLBL"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 30
$System_Drawing_Size.Width = 150
$ReportLBL.Size = $System_Drawing_Size
$ReportLBL.TabIndex = 11
$ReportLBL.Text = "Report Status"

$form.Controls.Add($ReportLBL)


$Button1.DataBindings.DefaultDataSourceUpdateMode = 0
$Button1.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9.75,1,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 44
$System_Drawing_Point.Y = 355
$Button1.Location = $System_Drawing_Point
$Button1.Name = "Button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 43
$System_Drawing_Size.Width = 104
$Button1.Size = $System_Drawing_Size
$Button1.TabIndex = 10
$Button1.Text = "Generate"
$Button1.UseVisualStyleBackColor = $True
$Button1.add_Click($Button1_OnClick)
$form.Controls.Add($Button1)

$CBOwner.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 350
$System_Drawing_Point.Y = 279
$CBOwner.Location = $System_Drawing_Point
$CBOwner.Name = "CBOwner"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 104
$CBOwner.Size = $System_Drawing_Size
$CBOwner.TabIndex = 8
$CBOwner.Text = "Owner"
$CBOwner.UseVisualStyleBackColor = $True
$form.Controls.Add($CBOwner)


$CBAdmin.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 250
$System_Drawing_Point.Y = 279
$CBAdmin.Location = $System_Drawing_Point
$CBAdmin.Name = "CBAdmin"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 104
$CBAdmin.Size = $System_Drawing_Size
$CBAdmin.TabIndex = 7
$CBAdmin.Text = "Admin"
$CBAdmin.UseVisualStyleBackColor = $True
$form.Controls.Add($CBAdmin)

$CBDelegate.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 154
$System_Drawing_Point.Y = 279
$CBDelegate.Location = $System_Drawing_Point
$CBDelegate.Name = "CBDelegate"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 104
$CBDelegate.Size = $System_Drawing_Size
$CBDelegate.TabIndex = 6
$CBDelegate.Text = "Delegate"
$CBDelegate.UseVisualStyleBackColor = $True
$form.Controls.Add($CBDelegate)

$textBox.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 140
$System_Drawing_Point.Y = 166
$textBox.Location = $System_Drawing_Point
$textBox.Name = "textBox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 186
$textBox.Size = $System_Drawing_Size
$textBox.TabIndex = 0
$textbox.AutoCompleteSource = 'CustomSource'
$textbox.AutoCompleteMode='SuggestAppend'
$textbox.AutoCompleteCustomSource=$autocomplete
$form.Controls.Add($textBox)

$SelMBXLBL.DataBindings.DefaultDataSourceUpdateMode = 0
$SelMBXLBL.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,1,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 34
$System_Drawing_Point.Y = 166
$SelMBXLBL.Location = $System_Drawing_Point
$SelMBXLBL.Name = "SelMBXLBL"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 100
$SelMBXLBL.Size = $System_Drawing_Size
$SelMBXLBL.TabIndex = 4
$SelMBXLBL.Text = "Select Mailbox"
$form.Controls.Add($SelMBXLBL)

$ENDdateLBL.DataBindings.DefaultDataSourceUpdateMode = 0
$ENDdateLBL.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,1,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 34
$System_Drawing_Point.Y = 92
$ENDdateLBL.Location = $System_Drawing_Point
$ENDdateLBL.Name = "ENDdateLBL"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 100
$ENDdateLBL.Size = $System_Drawing_Size
$ENDdateLBL.TabIndex = 3
$ENDdateLBL.Text = "End Date"
$form.Controls.Add($ENDdateLBL)

$StdateLBL.DataBindings.DefaultDataSourceUpdateMode = 0
$StdateLBL.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,1,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 34
$System_Drawing_Point.Y = 41
$StdateLBL.Location = $System_Drawing_Point
$StdateLBL.Name = "StdateLBL"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 100
$StdateLBL.Size = $System_Drawing_Size
$StdateLBL.TabIndex = 2
$StdateLBL.Text = "Start Date"
$form.Controls.Add($StdateLBL)

$dateTimePicker2.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 140
$System_Drawing_Point.Y = 95
$dateTimePicker2.Location = $System_Drawing_Point
$dateTimePicker2.Name = "dateTimePicker2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$dateTimePicker2.Size = $System_Drawing_Size
$dateTimePicker2.TabIndex = 1
$form.Controls.Add($dateTimePicker2)

$dateTimePicker1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 140
$System_Drawing_Point.Y = 41
$dateTimePicker1.Location = $System_Drawing_Point
$dateTimePicker1.Name = "dateTimePicker1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 200
$dateTimePicker1.Size = $System_Drawing_Size
$dateTimePicker1.TabIndex = 0
$form.Controls.Add($dateTimePicker1)


$LBLPATH.DataBindings.DefaultDataSourceUpdateMode = 0
$LBLPATH.FlatStyle = 1
$LBLPATH.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,0,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 30
$System_Drawing_Point.Y = 220
$LBLPATH.Location = $System_Drawing_Point
$LBLPATH.Name = "LBLPATH"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 30
$System_Drawing_Size.Width = 250
$LBLPATH.Size = $System_Drawing_Size
$LBLPATH.TabIndex = 0
$LBLPATH.Text = "Reports location is: $PathLocation"
$form.Controls.Add($LBLPATH)


$tooltip1.SetToolTip($button1, "Click here to generate the Auditing report")
$tooltip1.SetToolTip($LBLPATH, "Report location path")
$tooltip1.SetToolTip($textBox, "Start typing an Audit enabled mailbox alias")
$tooltip1.IsBalloon = $True

# Adding ICON

    $iconBase64 = 'AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAABAAABMLAAATCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAqisABqk2B0eoNwe
VqTkJwKo6DNqrPQ/yqz0P86o6DNupOgnBpzgHl6Y5B0i2SQAHAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArT0KGa
s9CJeuQRHysEgY/7FKGv+wRxf/r0YW/65EE/+uRBP/r0UV/7BHF/+xShr/sEgY/65BEfKsPAiZpzsKGgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAn0AACK5CDISxSBX1tVEf/7JKFv+vQg3/rkEL/65BC/+uQQv/rkEL/65BC/+uQQv/rkEL/65BC/+vQg3/skoW/7VRH/+xSBX1rkIMhLZJAAcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAALVKCxixSQ/It1Yi/7VRG/+wRw7/sEYN/7BGDf+wRg3/sEYN/7BGDf+wRg3/sEYN/7BGDf+wRg3/sEYN/7BGDf+wRg3/sEcO/7RQGv+3ViL/sUkPx7VKCxgAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC1ShEttFEW5LteKP+0UBX/s0wQ/7NMEP+zTBD/s0wQ/7NMEP+zTBD/s0wQ/7NMEP+zTBD/s0wQ/7NMEP+zTBD/s0wQ/7NMEP+zTBD/s0wQ/7RQFf+7Xij/tF
EW5LVKES0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAtVUVGLdVGeS+Yyv/t1IU/7ZREv+2URL/tlES/7ZREv+2URL/tlES/7ZREv+2URL/tlES/7ZREv+2URL/tlES/7ZREv+2URL/tlES/7ZRE
v+2URL/tlES/7dSFP++Yyv/uVYa5LVVFRgAAAAAAAAAAAAAAAAAAAAAAAAAALZJJAe5WBjHwGow/7lXFv+4VhT/uFYU/7hWFP+4VhT/uFYU/7hWFP+4VhT/uFYU/7hWFP+4VhT/uFYU/7hWFP+4VhT/
uFYU/7hWFP+4VhT/uFYU/7hWFP+4VhT/uFYU/7lXFv/AajD/uFgYybZJJAcAAAAAAAAAAAAAAAAAAAAAvFsWgcNuMf+9Xxv/u1sW/7tbFv+7Wxb/u1sW/7tbFv+7Wxb/u1sW/7tbFv+7Wxb/u1sW/7t
bFv+7Wxb/u1sW/7tbFv+7Wxb/u1sW/7tbFv+7Wxb/u1sW/7tbFv+7Wxb/u1sW/71fG//DbjH/u1sVhAAAAAAAAAAAAAAAAL9gFRjCaST1w24r/75hGP++YRj/vmEY/75hGP++YRj/vmEY/75hGP/Qjl
j/+fDm//nw5v/58Ob/9eja/75hGP++YRj/vmEY/9COWP/16Nn/9ejZ//Xo2f/cqH//vmEY/75hGP++YRj/vmEY/8NuK//CaiX2wmYUGQAAAAAAAAAAwWcalcp8Ov/BZxv/wWYa/8FmGv/BZhr/wWYa/
8FmGv/BZhr/wWYa/9OTXf/68ef/+vHn//rx5//57uP/wWYa/8FmGv/JeTb/9unb//rx5//68ef/682y/8JoHv/BZhr/wWYa/8FmGv/BZhr/wWcb/8p8Ov/BZhuZAAAAANWAKwbGcynyyHgw/8NrHP/D
axz/w2sc/8NrHP/Daxz/w2sc/8NrHP/Daxz/1Jde//ry6f/68un/+vLp//nw5f/Daxz/xXAj//DYwv/68un/+vLp//Ph0P/Idiz/w2sc/8NrHP/Daxz/w2sc/8NrHP/Daxz/yHgw/8Z0KvPVgCsGxnE
cSM6DPP/HcyL/xnAe/8ZwHv/GcB7/xnAe/8ZwHv/GcB7/xnAe/8ZwHv/WmV//+vPr//rz6//68+v/+fHo/8ZwHv/mwJz/+vPr//rz6//47uT/0IlE/8ZwHv/GcB7/xnAe/8ZwHv/GcB7/xnAe/8ZwHv
/HciL/zoM8/8ZxHEjJdiKX0YtE/8l2If/JdiH/yXYh/8l2If/JdiH/yXYh/8l2If/JdiH/yXYh/9meYf/79Oz/+/Ts//v07P/78+r/26Zu//v07P/79Oz/+/Ts/9uhZ//JdiH/yXYh/8l2If/JdiH/y
XYh/8l2If/JdiH/yXYh/8l2If/Ri0T/yXYil8x/KcPTjkP/y3sj/8t7I//LeyP/y3sj/8t7I//LeyP/y3sj/8t7I//LeyP/2qFi//v17v/79e7/+/Xu//v17v/58ej/+/Xu//v17v/lvZD/y3sj/8t7
I//LeyP/y3sj/8t7I//LeyP/y3sj/8t7I//LeyP/y3sj/9KOQv/NfifC0IYw3dSQQP/OgCX/zoAl/86AJf/OgCX/zoAl/86AJf/OgCX/zoAl/86AJf/dp2j/+/bx//v38f/79/H/+/fx//v38f/79/H
/+O7i/9KKN//QhCz/zoAl/86AJf/OgCX/zoAl/86AJf/OgCX/zoAl/86AJf/OgCX/1I9A/9GGMdzUjzf11pE9/9GFJ//RhSf/0YUn/9GFJ//RhSf/0YUn/9GFJ//TizH/1pE9/+Kxdf/8+PP//Pjz//
z48//8+PP/9unZ//z48//8+PP/68yk/9aTQP/WkT3/04sx/9GFJ//RhSf/0YUn/9GFJ//RhSf/0YUn/9GFJ//WkTz/1I849NaUO/TYl0D/04sp/9OLKf/Tiyn/04sp/9OLKf/UjCz/2JdA/9maRv/Zm
kb/47Z5//z59f/8+fX//Pn1//z59f/eq2P/+/Xu//z59f/8+fX/57+I/9maRv/Zmkb/15c//9SMLP/Tiyn/04sp/9OLKf/Tiyn/04sp/9eXP//WlTvz2JY23NygSf/WkCv/1pAr/9aQK//WkCv/15Iv
/9ufR//coUv/3KFL/9yhS//mun3//fn2//359v/9+fb//fn2/9yhS//nv4T//fn2//359v/89/L/5LVy/9yhS//coUv/259H/9eSLv/WkCv/1pAr/9aQK//WkCv/3J9I/9mXONvZlzHC4KhT/9mVLf/
ZlS3/2ZUt/9mVLv/eo0n/4KdR/+CnUf/gp1H/4KdR/+i+gP/9+vj//fr4//36+P/9+vj/4KdR/+CnUf/v1Kv//fr4//36+P/78+n/47Bl/+CnUf/gp1H/3qNI/9mVLf/ZlS3/2ZUt/9mVLf/gqFL/2p
gywduaLpXjsFz/25ov/9uaL//bmi//3aE9/+KtV//irVf/4q1X/+KtV//irVf/6cSF//37+f/9+/n//fv5//37+f/irVf/4q1X/+OvWf/25s///fv5//37+f/47Nv/5LFe/+KtV//irVf/3aA7/9uaL
//bmi//25ov/+OwXP/bmi6V3qAzRuSyWf/fozj/3qAy/96gMv/jr1H/5bRd/+W0Xf/ltF3/5bRd/+W0Xf/pv3f//vz6//78+v/+/Pr/9N+6/+W0Xf/ltF3/5bRd/+e3Zf/pv3b/6b92/+m/dv/nuGb/
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'
    $iconBytes = [Convert]::FromBase64String($iconBase64)
    $stream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
    $stream.Write($iconBytes, 0, $iconBytes.Length);
    #$iconImage = [System.Drawing.Image]::FromStream($stream, $true)
    $Form.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())

#Save the initial state of the form
$InitialFormWindowState = $form.WindowState
#Init the OnLoad event to correct the initial state of the form
$form.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form.ShowDialog()| Out-Null

} #End Function


GenerateForm
