
#"backup OWA Customization"

New-Item "c:\ver3\Client Access\all" -ItemType Directory

New-Item "c:\Ver3\Client Access\languageselection" -ItemType Directory

New-Item "c:\ver3\Client Access\ExpiredPassword"  -ItemType Directory



Copy-Item 'C:\Program Files\Microsoft\Exchange Server\V15\FrontEnd\HttpProxy\owa\auth\logon.aspx' 'c:\ver3\Client Access\all\'
Copy-Item 'C:\Program Files\Microsoft\Exchange Server\V15\FrontEnd\HttpProxy\owa\auth\15.2.792\themes\resources\logon.css' '\Ver3\Client Access\languageselection\'                        


Copy-Item 'C:\Program Files\Microsoft\Exchange Server\V15\FrontEnd\HttpProxy\owa\auth\15.2.792\themes\resources\Sign_in_arrow.png' 'C:\Ver3\Client Access\languageselection\'

Copy-Item 'C:\Program Files\Microsoft\Exchange Server\V15\FrontEnd\HttpProxy\owa\auth\15.2.792\themes\resources\Sign_in_arrow_rtl.png' 'c:\Ver3\Client Access\languageselection\'
Copy-Item 'C:\Program Files\Microsoft\Exchange Server\V15\FrontEnd\HttpProxy\owa\auth\15.2.792\themes\resources\owa_text_blue.png'  'c:\Ver3\Client Access\languageselection\'

Copy-Item 'C:\Program Files\Microsoft\Exchange Server\V15\FrontEnd\HttpProxy\owa\auth\ExpiredPassword.aspx' 'c:\ver3\Client Access\ExpiredPassword\'


#"backup Iphone customization"

New-Item "c:\ver3\iphone\FrontEnd\HttpProxy\Sync" -ItemType Directory

New-Item "c:\ver3\iphone\ClientAccess\Sync" -ItemType Directory


Copy-Item 'C:\Program Files\Microsoft\Exchange Server\V15\FrontEnd\HttpProxy\Sync\web.config' 'c:\ver3\iphone\FrontEnd\HttpProxy\Sync\'

Copy-Item 'C:\Program Files\Microsoft\Exchange Server\V15\ClientAccess\Sync\web.config' 'c:\ver3\iphone\ClientAccess\Sync\'


#"backup Scripting agent"


New-Item "c:\ver3\CmdletExtensionAgents" -ItemType Directory



Copy-Item 'C:\Program Files\Microsoft\Exchange Server\V15\Bin\CmdletExtensionAgents\ScriptingAgentConfig.xml' 'c:\ver3\CmdletExtensionAgents\'

#"backup Rerouting"

New-Item "c:\ver3\Custom" -ItemType Directory

Copy-Item 'C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\agents\Custom\Microsoft.Exchange.SBR.Config' 'c:\ver3\Custom\'


