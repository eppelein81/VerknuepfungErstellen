cls
#Variablen setzen
$VerzeichnisVerknuepfung=$home+"\desktop"
$Verzeichnis = Read-Host "In welchem Verzeichnis befindet sich die Datei:"
cd $Verzeichnis
ls
$Datei = Read-Host "Für welche Datei soll eine Verknüpfung erstellt werden:"
$Link = Read-Host "Wie soll die Verknüpfung heißen:"

#Prüfung der Variablen
[System.Windows.Forms.MessageBox]::Show("Es wird die Verknüpfung "+$Link+" für die Datei "+$Verzeichnis+"\"+$Datei+" in dem Verzeichnis "+$VerzeichnisVerknuepfung+" erstellt.","Titel",1)

#Verknüpfung wird erstellt
$WshShell = New-Object -comObject WScript.Shell
$link = $wshshell.CreateShortcut($VerzeichnisVerknuepfung+"\"+$Link+".lnk")
$link.targetpath = $Verzeichnis+"\"+$Datei
$link.save()

#Verzeichnisauflistung nach der Verknüpfung
cd $VerzeichnisVerknuepfung
ls
"Hier muss die Verknüpfung nun mit drinnen stehen!"
