########   Anleitung   ########

# 1. Ghostscript 64Bit install: https://www.ghostscript.com/releases/gsdnld.html
# 2. Variablen unten anpassen
# 3. Ausführen

## Variablen please change
$PDFOrdner = "C:\Users\dominik\Desktop\PDF - verkleinert\"
$DatumVon = Get-Date -Year 2015 -Month 03 -Day 01 -Hour 00 -Minute 00
$DatumBis = Get-Date -Year 2023 -Month 03 -Day 01 -Hour 00 -Minute 00
$EMailAn = "abc@abc.com"


####################################################################################################################

### Script Start ###

####################################################################################################################


### Variablen dont change ###
cd C:
$tempOrdner = ($env:TEMP + "\pdfcompresed123")
$GhostScript = ((Get-ChildItem 'C:\Program Files\gs' -Directory | Sort-Object LastWriteTime -Descending)[0].Fullname) + "\bin\gswin64.exe"
$SAVED = 0

### TEMP Ordner create
if (Test-Path $tempOrdner)
{Write-Host "TEMP Ordner exists!"}
else
{
New-Item -Path $tempOrdner -ItemType directory
}

## Sort
$PDFFiles = Get-ChildItem -Path $PDFOrdner | where { $DatumVon -lt $_.LastWriteTime -and $DatumBis -gt $_.LastWriteTime}

## copie to temp
foreach ($PDFFile in $PDFFiles)
{

    $Arguments = "-sDEVICE=pdfwrite -dSAFER -dCompatibilityLevel=1.4 -dEmbedAllFonts=true -dSubsetFonts=true -dColorImageDownsampleType=/Bicubic -dColorImageResolution=130 -dGrayImageDownsampleType=/Bicubic -dGrayImageResolution=130 -dMonoImageDownsampleType=/Bicubic -dMonoImageResolution=130 -dPDFSETTINGS=/screen -dNOPAUSE -dQUIET -dBATCH -sOutputFile=`"" + $tempOrdner + "\" + $PDFFile.Name + "`" `"" + $PDFOrdner +  $PDFFile.Name + "`""
    Start-Process $GhostScript -ArgumentList $Arguments
    Wait-Process gswin64 -Timeout 10 -WarningAction Continue
    $Error.Clear()
    if ((Get-Process -Name gswin64 -ErrorAction SilentlyContinue).Responding) 
    {   
        Write-Host "Ghostscript Problem"
        Stop-Process -Name gswin64
        sleep -Seconds 10
        continue;
    }

    if (Test-Path -Path ($tempOrdner + "\" + $PDFFile))
    {
        $File1 = Get-ChildItem -Path ($tempOrdner + "\" + $PDFFile)
        $FileSize1 = $File1.Length/1MB

        $File2 = Get-ChildItem -Path ($PDFOrdner + "\" + $PDFFile)
        $FileSize2 = $File2.Length/1MB

        if ($FileSize1 -lt $FileSize2)
            {
                Write-Host "good its smaller"
                Copy-Item -Path $File1 -Destination $File2 -Recurse -Force
                Remove-Item -Path $File1 -Recurse -Force
                Write-Host "PDF has been shrinked:" $PDFFile
            }
        else
            {
                Write-Host "PDF doesent get shrinked:" $PDFFile
                Remove-Item -Path $File1 -Recurse -Force
            }
            
            $SAVED = $FileSize2 - $FileSize1 + $SAVED
            
    }

    else
    {
        Write-Host ("file cant copy: " + ($tempOrdner + "\" + $PDFFile))
        Remove-Item -Path $File1 -Recurse -Force
    }
}


Get-ChildItem $tempOrdner | Remove-Item -Force


	# Mailflow 
	  $smtpServer = "smtp.abc.com"
	  $smtpFrom = "shrinkpdf@abc.com"
	  $smtpTo = $EMailAn
	  $messageSubject = "Status: shrink PDF"
	  
	  $message = New-Object System.Net.Mail.MailMessage $smtpFrom, $smtpTo
	  $message.Subject = $messageSubject
	  $message.isbodyHTML = $true
	  $message.body = 'You saved ' + $SAVED + ' MB! <p>
Ordner: ' + $PDFOrdner
	  
	  $smtp = New-Object Net.Mail.SmtpClient($smtpServer)
	  $smtp.send($message)
