# RPC Logs Exchange Outlook Builds

# Exchange 2010 Path:
#$logpath = 'C:\Program Files\Microsoft\Exchange Server\V14\Logging\RPC Client Access'

# Exchange 2013 and later Path:
$logpath = 'C:\Program Files\Microsoft\Exchange Server\V15\Logging\RPC Client Access'

$files = Get-ChildItem $logpath -ea 1 |Where-Object {$_.LastWriteTime -ge (Get-Date).AddDays(-2)}
$logs = $files | ForEach {Get-Content $_.FullName}| Where-Object {$_ -notlike '#*'}
$result = $logs |ConvertFrom-Csv -Header date-time,session-id,seq-number,client-name,organization-info,client-software,client-software-version,client-mode,client-ip,server-ip,protocol,application-id,operation,rpc-status,processing-time,operation-specific,failures
$uniqueClients = $result | Where-Object {$_.'client-software' -eq 'OUTLOOK.EXE'}| select client-name,client-ip,client-software-version,client-software | sort client-name -unique
$fr = @()
$uniqueClients | %{ $rt = get-recipient $_."client-name" ; $obj=""|select Email,Build,RecType ; $obj.Email=$rt.PrimarySMTPAddress;$obj.Build=$_."client-software-version";$obj.RecType=$rt.RecipientTypeDetails; $fr += $obj;}
$fr | export-csv Outlook_BuildsEmail.csv -notype
