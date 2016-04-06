## Starwind SAN Server Miner 
This is a little Powershell script that fetch metric's values from Starwind SAN Server.

Actual release 0.9.0

Tested on:
- Windows Server 2012 R2, StarWind 8, PowerShell 4


Support objects:
- _Server_ - Starwind SAN Server info;
- _Target_ - Starwind iSCSI target;
- _Device_ - Starwind iSCSI Device.

Actions:
- _Discovery_ - Make Zabbix's LLD JSON;
- _Get_       - Get metric from collection's item;
- _Avg_       - Calculate average of metric values from collection of items;
- _Min_       - Find minimal value of metrics from collection of items;
- _Max_       - Find maximal value of metrics from collection of items;
- _Last_      - Get last value of metric from collection of items;
- _Sum_       - Sum metrics of collection's items;
- _Count_     - Count collection's items.

Zabbix's LLD available to:
- _Target_;
- _Device_.

Virtual keys for 'Server' object is:
- _PerformanceData.CPU_ - ;
- _PerformanceData.RAM_ - .

Virtual keys for 'Target' object is:
- _Initiator_ - Number of connected initiators.

Virtual keys for 'Server', 'Target', 'Device' object is:
- _PerformanceData.ReadBandwidth_, _PerformanceData.WriteBandwidth_, _PerformanceData.TotalBandwidth_ - ;
- _PerformanceData.TotalIOPs_ - .


###How to use standalone

    # Make Zabbix's LLD JSON for iSCSI Devices.
    powershell.exe -NoProfile -ExecutionPolicy "RemoteSigned" -File "swm.ps1" -Action "Discovery" -ObjectType "Device"    

    # Return average write bandwith for Target with ID=0x00000000004601D0 calculated for last hour
    ... "swm.ps1" -Action "Avg" -ObjectType "Target" -Key "PerformanceData.ReadBandwidth" -Id "0x00000000004601D0" -TimePeriod "60"

###How to use with Zabbix
1. Just include [zbx\_swm.conf](https://github.com/zbx-sadman/StarWind/tree/master/Zabbix_Templates/zbx_swm.conf) to Zabbix Agent config;
2. Put _swm.ps1_ to _C:\zabbix\scripts_ dir. If you want to place script to other directory, you must edit _zbx\_swm.conf_ to properly set script's path; 
3. Set Zabbix Agent's / Server's _Timeout_ to more that 3 sec (may be 10 or 30);
4. Import [template](https://github.com/zbx-sadman/StarWind/tree/master/Zabbix_Templates) to Zabbix Server;
6. Enjoy.

**Note**
Do not try import Zabbix v2.4 template to Zabbix _pre_ v2.4. You need to edit .xml file and make some changes at discovery_rule - filter tags area and change _#_ to _<>_ in trigger expressions. I will try to make template to old Zabbix.

###Hints
- To see available metrics, run script without "-Key" option: _... "swm.ps1" -Action "Get" -Object "Server"_;
- To measure script runtime use _Verbose_ command line switch;
- To get on Zabbix Server side properly UTF-8 output when have non-english (for example Russian Cyrillic) symbols in Computer Group's names, use  _-consoleCP **your_native_codepage**_ command line option. For example to convert from Russian Cyrillic codepage (CP866): _... "swm.ps1" ... -consoleCP CP866_;

