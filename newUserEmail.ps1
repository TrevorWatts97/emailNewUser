$emailList = [System.Collections.Generic.List[string]]::new()
$emailList.Add("trwatts@scprt.com")
$emailList.Add("dmarshall@scprt.com")
$flag = $false
while($flag -eq $false){
    $answer = read-host "Answer"
    switch($answer){
        #add email
        1{
        $addEmail = read-host "New Email"
        $emailList.Add($addEmail)
        break
        }
        #delete email
        2{
        for($i = 0; $i -le ($emailList.count - 1); $i++){
            $string = '' + ($i + 1)+ '. ' + $emailList[$i]
            echo $string
        }
        $removeEmail = read-host "Delete Email"
        $emailList.remove($emailList[$removeEmail - 1])
        break
        }
        #confirm emails
        3{"confirm"
        $flag = $true
        break
        }
        
        default{"Not a valid answer"}
    }
}
$sendTo = $null
foreach($email in $emailList){
    $sendTo += $email + ";"
}
echo $emailList.count
echo $sendTo