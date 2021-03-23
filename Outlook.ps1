$Outlook = New-Object -ComObject Outlook.Application

$MailArray = get-content -path C:\Users\username\email\emailadresses.txt -Raw #email addresses should be delimited with ; and a space 

$mailbody  = get-content -path C:\Users\username\email\mailbody.txt

$outfileName = "C:\Users\username\gpreport.html"

$files = Get-ChildItem -path C:\Users\username\email



foreach ($f in $files){
  # for this to work we need to have a file with the email address as the name 
  $filenameWE=[System.IO.Path]::GetFileNameWithoutExtension($f)

  if($MailArray -match $filenameWE){

    write-host("found")#or write the output to a log file to keep track 

    $massage=get-content -path $f.FullName

    $Mail = $Outlook.CreateItem(0)

    $mail.Attachments.Add($outfileName)

    $Mail.Subject = "Hello this is one of the best days of your life"
    #this is how you would replace something in your mail body 
    $newmailbody = $mailbody.replace('*youwillneverfindthis*' ,$massage)

    $Mail.HTMLBody = "$newmailbody"

    $Mail.to = "$filenameWE"

    $Mail.Save() #should be replaced by send 

  }else{

  write-host("This is else statement")

  }

}
