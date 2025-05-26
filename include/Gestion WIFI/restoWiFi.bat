cd \temp\wifi
for %%f in (.\*.xml) do ( 
netsh wlan add profile filename=".\%%f" 
) 
