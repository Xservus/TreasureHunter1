#___________                                                     ___ ___               __                  ____ 
#\__    ___/______   ____ _____    ________ _________   ____    /   |   \ __ __  _____/  |_  ___________  /_   |
#  |    |  \_  __ \_/ __ \\__  \  /  ___/  |  \_  __ \_/ __ \  /    ~    \  |  \/    \   __\/ __ \_  __ \  |   |
#  |    |   |  | \/\  ___/ / __ \_\___ \|  |  /|  | \/\  ___/  \    Y    /  |  /   |  \  | \  ___/|  | \/  |   |
#  |____|   |__|    \___  >____  /____  >____/ |__|    \___  >  \___|_  /|____/|___|  /__|  \___  >__|     |___|
####Find files of interest in the user profile path
# mrR3b00t
# Version 0.1
# Created by Daniel Card (Xservus Limited)
# use at your own risk ;)
# this might set DLP systems off in a crazy way so be careful when you run this!
#####################################
#This one is for Cral3r!!!!!!!!!!
####################################

$extentsions = "*.xlsx","*.docx","*.doc","*.pptx","*.ppt"
$Logfile = "c:\temp\treasure_hunter.log"
$hunterpath = 'C:\Users'

foreach($extentsion in $extentsions){

#hunt for each extention
#write-host $extentsion
#Get-ChildItem -Path 'C:\Users' -Recurse -Force -ErrorAction SilentlyContinue -Filter ($extentsion) | Select-Object FullName


###############check for office files with passwords in them###################
## thanks to https://sharepoint.stackexchange.com/questions/68987/check-for-password-protected-office-documents

$Items = Get-ChildItem -Path $hunterpath -Recurse -Force -ErrorAction SilentlyContinue -Filter ($extentsion)| Select-Object FullName

foreach($Item in $Items){

write-host $Item.FullName -ForegroundColor Yellow 

 # Open the file and select first ~2kb
            #$Binary = $Item.FullName.file.OpenBinary()

            # Read the entire file to an array of bytes.
            $bytes = [System.IO.File]::ReadAllBytes($Item.FullName)
            # Decode first 2000 bytes to a text assuming ASCII encoding.
            $text = [System.Text.Encoding]::ASCII.GetString($bytes, 0, 2000)

            #incase you want to read the header
            #write-host $text

            # Test for pattern indicating encrypted Office 2007/+ format file
           if($text -match "E.n.c.r.y.p.t.e.d.P.a.c.k.a.g.e") {

               write-host "Office password protected file detected" -ForegroundColor Magenta

                # Record encrypted file detection in log
                $log = "Encrypted file found," + $Item.FullName
                $log | Out-File $Logfile -Append
                                
           }

            #close the for each item in items loop
            }


}
