param(
$mailbox="your_email@company.com",
$itemsView=1000,
$password="Your_password"
)

[Reflection.Assembly]::LoadFile(".\2.0\Microsoft.Exchange.WebServices.dll")

$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1

$s = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$s.Credentials =  New-Object Net.NetworkCredential($mailbox,$password)
$s.Url = new-object Uri("https://webmail.company.com/EWS/Exchange.asmx");


$iv = new-object Microsoft.Exchange.WebServices.Data.ItemView($itemsView)

#Retrieve undeleted messages
$iv.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow 

$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$psPropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text;
$iv.PropertySet = $psPropertySet

$msgs = $s.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$iv)


foreach ($msg in $msgs.Items)
{ 
	$msg.Load()
	$body = $msg.Body.Text        
	write-host "BODY $body"
}



