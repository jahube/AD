# create Mailuser - "Error: did not find recipient"

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