<#======= DEKLARACJE ZMIENNYCH I FUNKCJI GLOBALNYCH==========#>
function Unistall-TeamViewer
{
    quser | Select-String Disc | ForEach {logoff ($_.tostring() -split ' +')[2]} #Wylogowanie pozostałych użytkowników poza bieżącym w celu całkowitego odinstalowania programu
    Write-Host Dezinstalacja teamviewera rozpoczeta -BackgroundColor Red

    if(Test-Path "C:\Program Files (x86)\TeamViewer\uninstall.exe")
    {
        Start-Process -FilePath 'C:\Program Files (x86)\TeamViewer\uninstall.exe' -ArgumentList /S
    }
    
    if(Test-Path -Path C:\Users\Public\Desktop\TeamViewer.lnk)
    {
        Write-Host Usuwanie skrótu TeamViewera -BackgroundColor Red
        Remove-Item C:\Users\Public\Desktop\TeamViewer.lnk –Force
    }
}

$isTeamViewerInstalled = Get-Package -Name *teamviewer*
$main = "C:\TeamViewer\"
$temp = "C:\TeamViewer\temp"
$mainEXE = "C:\TeamViewer\TeamViewerQS-idc69yz48b.exe"
$tempEXE = "C:\TeamViewer\temp\TeamViewerQS-idc69yz48b.exe"
<#-----------------------------------------------------------#>



<#---- PROCEDURA --------#>
<#  1. Sprawdź czy TeamViewer jest zainstalowany na komputerze i odinstaluj  #>
if($isTeamViewerInstalled)
{
    Unistall-TeamViewer
}
else
{
    Write-Host Brak zainstalowanego teamviewera -BackgroundColor Red
}

<#  2. Sprawdź czy istenieje odpowiednie drzewo folderów  #>
if (Test-Path -Path $main)
{
        if(Test-Path $temp)
        {
            Write-Host Wszystkie katalogi są na miejscu -BackgroundColor Red
        }
        else
        {
            Write-Host Tworzenie katalogu $temp -BackgroundColor Red
            New-Item -ItemType Directory -Force -Path $temp
        }
}
else
{
    Write-Host Tworzenie katalogu $main -BackgroundColor Red
    New-Item -ItemType Directory -Force -Path $main
    New-Item -ItemType Directory -Force -Path $temp
}

<#  3. Pobierz najnowszą wersję TeamViewer QS  #>
if (Test-Path $tempEXE)
{
    Write-Host Usuwanie poprzedniej wersji TeamViewer QS z lokalizacji temp jeżeli nie został usunięty wcześniej -BackgroundColor Red
    Remove-Item $tempEXE -Force
    sleep 5
}
    Invoke-WebRequest -Uri https://download.teamviewer.com/download/TeamViewerQS.exe -OutFile $tempEXE
    $secondsElapsed = 0

    while(1)
    {
        if(Test-Path  $tempEXE)
        {
            Write-Host Pobrano najnowszą wersję TeamViewera z internetu -BackgroundColor Red
            break
        }
        else
        {
            Write-Host Oczekiwanie na pobranie TeamViewera  $secondsElapsed -BackgroundColor Red
            $secondsElapsed++
            sleep 1

            if($secondsElapsed -eq 180)
            {
                Write-Host Pobrano najnowszą wersję TeamViewera z serwera server -BackgroundColor Red

                if(Test-Path \\server\TeamviewerQS\TeamViewerQS-idc69yz48b.exe)
                {
                    Copy-Item \\server\TeamviewerQS\TeamViewerQS-idc69yz48b.exe -Destination $tempEXE
                }
                else
                {
                    Write-Host Nie pobrano z internetu i nie pobrano z serwera -BackgroundColor Red
                    exit
                }
            }
        }
    }

<#  4.Podmiana TeamViewera na nowszą wersję   #>
if(Test-Path $mainEXE)
{
    Write-Host Wykryto poprzednią wersję TeamViewera -BackgroundColor Red

    if((Get-Process TeamViewer).id)
    {
        $nid = (Get-Process TeamViewer).id
       
        Write-Output Czekam na zamknięcie TeamViewera -BackgroundColor Red
        Wait-Process -Id $nid
    }
    Remove-Item $mainEXE
}
if(Test-Path $tempEXE)
{
    Write-Host Wykryto TeamViewera w katalogu temp -BackgroundColor Red
    Copy-Item $tempEXE -Destination $mainEXE
    Write-Host Kopiuje TeamViewera z katalogu temp do main -BackgroundColor Red
    Write-Host Czekam 5 sekund -BackgroundColor Red
    sleep 5
    Remove-Item $tempEXE
    Write-Host Usunięto TeamViewera z katalogu temp -BackgroundColor Red
}
else
{
    Write-Host Błąd przy podmienianiu na nowszą wersje -BackgroundColor Red
}

<#  5.Sprawdzenie czy istnieje skrót na pulpicie i stworzenie go jeżeli nie istnieje  #>
if(Test-Path -Path $env:Public\Desktop\Zdalna_pomoc.lnk)
{
    Write-Host Skrót jest poprawny, koniec pracy -BackgroundColor Red
}
else
{
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("$env:Public\Desktop\Zdalna_pomoc.lnk")
    $Shortcut.TargetPath = $mainEXE
    $Shortcut.Save()
    Write-Host Utworzono skrót na puplicie, koniec pracy. -BackgroundColor Red
}