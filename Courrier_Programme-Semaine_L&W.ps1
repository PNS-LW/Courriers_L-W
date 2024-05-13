#Indique le numéro de la semaine au cours de l'année
    $CountDay = 0
        IF
            ((Get-Date -UFormat "%A %d %B %Y") -like 'lundi *')
                {
                    $Week_Number = Get-Date
                    $Day = Get-Date -UFormat "%A %d %B %Y"
                }

        ELSE
            
            {
                Do
                        {
                            $CountDay--
                            $Day = ((Get-Date).AddDays(+$countday)) | Get-Date -UFormat "%A %d %B %Y"
                        }
                    
                    Until($Day -like 'lundi *')

                    $Week_Number = ((Get-Date).AddDays(+$CountDay))
                    $Day = ((Get-Date).AddDays(-$CountDay)) | Get-Date -UFormat "%A %d %B %Y"
            }

    $Week_Number = ($Week_Number.DayOfYear / 7) + 1
    #$Week_Number = ($Week_Number.DayOfYear / 7)
    $Week_Number = '{0:f0}' -f $Week_Number


#Obtient le fuseau horaire UTC
    $Get_UTC_Global = Get-TimeZone
    $Get_UTC = $Get_UTC_Global | Select-Object BaseUtcOffset

        IF($Get_UTC -like '*01*')

            {
                $UTC_Folder = 'UTC01'
            }
        
        ELSEIF($Get_UTC -like '*02*')

            {
                $UTC_Folder = 'UTC02'
            }





#Chemin du dossier Global
    $Path_Source = ((Split-Path -Path $MyInvocation.MyCommand.Definition -Parent) + "\")

#Indique l'année en cours
    $Year = Get-Date -Format ("yyyy")

#Indique le chemin des fichiers de destination
    $Global_Destination = ($Path_Source + "Courriers_Finaux\Programme_Semaine\" + $Year + "\" + ($Year + "_Semaine-" + $Week_Number))
    $TMP_Folder = $Path_Source + 'TMP'
    $File_Week = ($Global_Destination + "\" + "Annee-" + $Year + "_" + "Programme_Semaine-" + $Week_Number + ".txt")
    $File_Suite01_Week = ($TMP_Folder + "\Annee-" + $Year + "_" + "Suite01_Programme_Semaine-" + $Week_Number + ".txt")
    $File_Suite_Week = ($Global_Destination + "\Annee-" + $Year + "_" + "Suite_Programme_Semaine-" + $Week_Number + ".txt")

    $File_Content_Week = ($TMP_Folder + "\" + "Content_Week")
    $File_Content02_Week = ($TMP_Folder + "\" + "Content02_Week")
    $File_Content03_week = ($TMP_Folder + "\" + "Content03_Week")
    $File_Content04_week = ($TMP_Folder + "\" + "Content04_Week")
    $File_Content01_UTC = ($TMP_Folder + "\" + "Content01_UTC.txt")
    $File_Content02_UTC = ($TMP_Folder + "\" + "Content02_UTC.txt")


function Reset-Folder_TMP {

                                   Get-ChildItem -Path $Path_Source -Filter '.txt' -Directory | ForEach-Object { Remove-Item -Path ($_.fullname) -Force -Recurse -Confirm:$false }
                                   New-Item -Path $TMP_Folder -ItemType Directory -force

                              }

Reset-Folder_TMP

#Creation du dossier qui va recevoir les fichiers de destination
    New-Item -Path $Global_Destination -ItemType Directory -force

#Indique le chemin des fichiers sources qui alimenteront les fichier finaux en texte
    $Week_Number | % {if($_ % 2 -eq 0 ) {$paire_impaire = 'paire'} ELSE {$paire_impaire = 'impaire'}}
    $Template_Source = ($Path_Source + 'Templates\Programme_Semaine\' + $UTC_Folder + '\' + $UTC_Folder + '_Template_Semaine_' + $paire_impaire + '.txt')
    $Template_Source_Suite = ($Path_Source + 'Templates\Programme_Semaine\' + $UTC_Folder +'\' + $UTC_Folder + '_Template_Semaine_' + $paire_impaire + '_Suite.txt')


#Alimente le premier fichier final

    $count = 0
    $count2 = $countday

    $File_Content_week = ($File_Content_week + ".txt")

    (get-content $Template_Source) | Foreach-Object {


    $_.replace('NumberWeek', $Week_Number).replace('YYYY',$Year).replace('XXX0',$Day)

    } | Set-Content $File_Content_week

Do{

    $count++
    $count2++
    $File_Content02_week = ($TMP_Folder + "\" + "Content02_Week")
    $File_Content02_week = ($File_Content02_week + ".txt")
    $Day_Plus = ((Get-Date).AddDays(+$count2))
    $Day_Plus = Get-Date $Day_Plus -UFormat "%A %d %B %Y"

    (get-content $File_Content_week) | Foreach-Object {

    $_.replace(('XXX' + $count),$Day_Plus)

    } | Set-Content $File_Content02_week

    Remove-Item $File_Content_week

    Start-Sleep -Seconds 1

    (get-content $File_Content02_week) | Set-Content $File_Content_week

}Until($count -eq 4)

Copy-Item $File_Content02_week $File_Week
      

#Alimente le fichier suite final

    $count--
    $count2--

    $File_Content03_week = ($File_Content03_week + ".txt")

    (get-content $Template_Source_Suite) | Foreach-Object {


    $_.replace('NumberWeek', $Week_Number).replace('YYYY',$Year)

    } | Set-Content $File_Content03_week

    Do{

    $count++
    $count2++
    $File_Content04_Week = ($TMP_Folder + "\" + "Content04_Week")
    $File_Content04_week = ($File_Content04_week + ".txt")
    #$Day = Get-Date -UFormat "%A %d %B %Y"
    $Day_Plus = ((Get-Date).AddDays(+$count2))
    $Day_Plus = Get-Date $Day_Plus -UFormat "%A %d %B %Y"

    (get-content $File_Content03_week) | Foreach-Object {

    $_.replace(('XXX' + $count2),$Day_Plus)

    } | Set-Content $File_Content04_week

    Remove-Item $File_Content03_week

    Start-Sleep -Seconds 2

    (get-content $File_Content04_week) | Set-Content $File_Content03_week
    

}Until($count -eq 7)

(get-content $File_Content04_week) | Set-Content $File_Suite_Week

Remove-Item $TMP_Folder -Recurse -force -Confirm:$false

explorer.exe $Global_Destination