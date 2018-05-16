$o = New-Object -comobject outlook.application
$n = $o.GetNamespace("MAPI")
#$f = $n.PickFolder()
$f = $n.Folders.Item("username@domain.com").Folders.item("Reports").Folders.item("Application_Folder").Folders.item("Systems")
$filepath = "C:\Users\username\Desktop\"
#write-host $f
if ($args){
	foreach ($arg in $args){
		$arg = $arg.split("=")
	}

	if ($arg[0].Equals("-d")){
		$RptCount = $arg[1]
		$dates = @()
		for ($i=0; $i -lt $RptCount; $i++){
			$dates = $dates + ("{0:yyyy-MM-dd}" -f [datetime](get-date).AddDays(-$i))
		
		}
	}else{ 
		$RptCount = 1 
		$dates = (get-date -format "yyyy-MM-dd")
		}
}else{
$dates = (get-date -format "yyyy-MM-dd")
}

$f.Items| foreach {
	$t = "{0:yyyy-MM-dd}" -f [datetime]$_.creationtime
	if ($dates.contains($t)){
		$_.attachments|foreach {
			$a = $_.filename
			$now = (get-date -format "yyyy-MM-dd")		
			$saveReport = (join-path $filepath ($t + " report.pdf"))
			$_.saveasfile($saveReport)		
		}
	}	
}
