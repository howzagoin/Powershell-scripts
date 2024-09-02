# Ensure the ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

# Define credentials and connection details
$AdminUsername = 'welkin@firstfinancial.com.au'
$AdminPassword = 'Dob82498'
$UserToSearch = 'james.wrigley@firstfinancial.com.au'

# Connect to Exchange Online
$SecurePassword = ConvertTo-SecureString $AdminPassword -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential ($AdminUsername, $SecurePassword)
Connect-ExchangeOnline -Credential $Credentials

Write-Host "Successfully connected to Exchange Online."

# Define the list of email addresses to add to the safe sender list
$SafeSenders = @(
    "a.griffin@stanfordbrown.com.au",
    "aawhite456@gmail.com",
    "abradse@gmail.com",
    "accelerateservice@tal.com.au",
    "adam.m.stawski@gmail.com",
    "adam.trimboli@btfinancialgroup.com",
    "adam@adamelliot.com.au",
    "adam@cynch.com.au",
    "adamelliot@icloud.com",
    "admin@adviceforlife.com.au",
    "admin@adviserratings.com.au",
    "admin@assuredsupport.com.au",
    "admin@hopedermatology.com.au",
    "admin@hub24.com.au",
    "admin@simplyaskit.com.au",
    "admin@timeadvice.com.au",
    "adrian.flores@momentummedia.com.au",
    "adriana.brink@inwealth.com.au",
    "adriana.colaneri@gmail.com",
    "adviser@macquarie.com",
    "adviser@onepath.com.au",
    "advisers@genlife.com.au",
    "adviserserve@boardroomlimited.com.au",
    "aelliot@tesla.com",
    "afeigin@fastmail.fm",
    "agirolami@bhtpartners.com.au",
    "aharel@seek.com.au",
    "ahbeetsang@hotmail.com",
    "aimee@aimeeharel.com.au",
    "alanna.richmond@mk.com.au",
    "alaynnaelliot@hotmail.com",
    "alex.mattingly@outlook.com",
    "alexander.ramsey@flexdapps.com",
    "alexander.walker@salesforce.com",
    "alice@onefellswoop.com.au",
    "alicep@onefellswoop.com.au",
    "alison.haigh@kaplan.edu.au",
    "alison.matthews@kodacapital.com",
    "alisonworland1@gmail.com",
    "alister@tempuspartners.com.au",
    "alkichemist@gmail.com",
    "allenthelma49@gmail.com",
    "allison.haigh@kaplan.edu.au",
    "amanda.duong@mckeanpark.com.au",
    "amy.brandmeier@xero.com",
    "ana.sandova.l1989@hotmail.com",
    "andrew.chapman@mac.com",
    "andrew.dobson@findex.com.au",
    "andrew.oxley@umowlai.com.au",
    "andrew@omegadigital.com.au",
    "andrew@wheatleyfinance.com",
    "andrew_kalopedis@ssga.com",
    "andrewcoops76@gmail.com",
    "andygilbs@gmail.com",
    "angelaabbey@bigpond.com",
    "ania.przekwas@aia.com",
    "anna@pathwayam.com.au",
    "anne.noakes@agedcaresteps.com.au",
    "anthony@tonywhitegroup.com.au",
    "apeters@hallmarc.com.au",
    "arbonk@guardiannm.com.au",
    "archiex@rosepartners.com.au",
    "ash.mcauliffe@mercer.com",
    "ash@empiraa.com",
    "ash@ilovesmsf.com",
    "ashleigh@a-esque.com",
    "ashley.mcauliffe@rmit.edu.au",
    "astoker@sccs.com.au",
    "astoker@sxiq.com.au",
    "au.uwpreassess@aia.com",
    "au.vic_taspreassess@aia.com",
    "auservices@metlife.com",
    "authorities@cbusmail.com.au",
    "authority@australiansuper.com.au",
    "avbrown2@bigpond.com"
)

# Get the current safe sender list and remove duplicates
$CurrentSafeSenders = (Get-MailboxJunkEmailConfiguration -Identity $UserToSearch).TrustedSendersAndDomains
$SafeSenders = $SafeSenders | Sort-Object -Unique

# Ensure that we only add new safe senders
$NewSafeSenders = ($CurrentSafeSenders + $SafeSenders) | Sort-Object -Unique

# Update the safe sender list
Set-MailboxJunkEmailConfiguration -Identity $UserToSearch -TrustedSendersAndDomains $NewSafeSenders

Write-Host "Safe sender list updated successfully for $UserToSearch."
