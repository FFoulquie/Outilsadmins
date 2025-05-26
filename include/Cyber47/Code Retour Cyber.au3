#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\..\..\..\..\..\..\Icones\Metro Icons\cadenas.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <Date.au3>


$calcul = number(@YDAY) * number(@mday) * Number(@HOUR)
MsgBox(0,'Code Retour Cyber',"Le code est :  " & $calcul)