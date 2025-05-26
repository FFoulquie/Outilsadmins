@echo off
ECHO Arret du service Windows Update
net stop wuauserv
rmdir c:\windows\SoftwareDistribution /s /q
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v SusClientId /f
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v SusClientIdValidation /f
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v PingID /f
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v AccountDomainSid /f
CLS
ECHO Demarrage du service Windows Update
net start wuauserv

ECHO Force la prise de contact entre le poste client et le serveur
wuauclt.exe /resetauthorization /detectnow
wuauclt.exe /detectnow
wuauclt.exe /reportnow
cls
Color C 
ECHO Operation terminee 

Pause