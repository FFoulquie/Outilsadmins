title Sauvegarde des parametres WiFi
Pushd "%~dp0"
color 9
@echo off
mkdir wifi
netsh wlan export profile key = clear folder = WiFi


echo termine
