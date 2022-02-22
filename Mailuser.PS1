"Error: did not find recipient"

       1) Mailuser erstellen
         ===================

                     CLOUD
                     -----

           $alias = "Dein.Email-Alias"
[GUID]$objectGUID = "ebd564ef-7033-491b-aa94-815b9fcd36cc"
          $domain = "FIRMA.de"
      $ONMSdomain = "FIRMA.onmicrosoft.com"
             $PWD = "!anges_P@ssw0rt"

         $Param = @{ Name = $alias     
     ExternalEmailAddress = "SMTP:$alias@$domain"
       PrimarySmtpAddress = "$alias@$domain"
MicrosoftOnlineServicesID = "$alias@$ONMSdomain"
                 Password = ConvertTo-SecureString $PWD -AsPlainText -Force
              ImmutableId = [system.convert]::ToBase64String(([GUID]$objectGUID).ToByteArray()) }

              New-MailUser @Param

2) Exchangeguid von Onprem setzen
   ==============================

Set-MailUser $alias -ExchangeGuid 1cdcb6c3-5290-43ec-8e3f-580ff3ee15a9


           ONPREM
           ------

1) ObjectGUID für Identity Hardmatch
======================================
get-aduser "USERNAME" | FT ObjectGUID


1) ExchangeGUID für Exchange Hardmatch
======================================
a) AD
      $mbx = get-aduser "USERNAME" -Properties msExchMailboxGuid
[guid]$mbx.msExchMailboxGuid
-----------------------------------------
B) Exchange Management shell

Get-Mailbox "USER" |ft name,exchangeguid

1cdcb6c3-5290-43ec-8e3f-580ff3ee15a9


