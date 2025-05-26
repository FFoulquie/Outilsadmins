title Sauvegarde des parametres WiFi
Pushd "%~dp0"
color 9
@echo off
mkdir wifi
netsh wlan export profile key = clear folder = WiFi
echo for %%%f in (.\*.xml) do ( > Wifi\restoWiFi.bat
echo netsh wlan add profile filename=".\%%%f" >> Wifi\restoWiFi.bat
echo ) >> Wifi\restoWiFi.bat
color 2

echo termine
