<#                                          
    .SYNOPSIS  
        Return Starwind SAN Server's metrics value, calculate its, make LLD-JSON for Zabbix

    .DESCRIPTION
        Return Starwind SAN Server's metrics value, calculate its, make LLD-JSON for Zabbix

    .NOTES  
        Version: 0.9.0
        Name: Starwind SAN Server Miner
        Author: zbx.sadman@gmail.com
        DateCreated: 06APR2016
        Testing environment: Windows Server 2012 R2, StarWind 8, PowerShell 4

    .LINK  
        https://github.com/zbx-sadman

    .PARAMETER Action
        What need to do with collection or its item:
            Discovery - Make Zabbix's LLD JSON;
            Get       - Get metric from collection's item;
            Avg       - Calculate average of metric values from collection of items;
            Min       - Find minimal value of metrics from collection of items;
            Max       - Find maximal value of metrics from collection of items;
            Last      - Get last value of metric from collection of items;
            Sum       - Sum metrics of collection's items;
            Count     - Count collection's items.

    .PARAMETER ObjectType
        Define rule to make collection:
            Server  - Starwind SAN Server info;
            Target  - Starwind iSCSI target;
            Device  - Starwind iSCSI Device;

    .PARAMETER Key
        Define "path" to collection item's metric 

        Virtual keys for 'Server' object:
            PerformanceData.CPU -
            PerformanceData.RAM -

        Virtual keys for 'Target' object:
            Initiator - Number of connected initiators.

        Virtual keys for 'Server', 'Target', 'Device' object:
            PerformanceData.ReadBandwidth, PerformanceData.WriteBandwidth, PerformanceData.TotalBandwidth - 
            PerformanceData.TotalIOPs -

    .PARAMETER TimePeriod
        How much minutes contains time period for selecting data for PerformanceData.* virtual key. 
        For example, value equal 5 mean that will processed all data from (Now-5min) to (Now).
        0 - take only last PerformanceData record;
        1..60 - take all PerformanceData records for specified mins;
        >60 - have no sense.

    .PARAMETER ID
        Used to select only one item from collection

    .PARAMETER ErrorCode
        What must be returned if any process error will be reached

    .PARAMETER ConsoleCP
        Codepage of Windows console. Need to properly convert output to UTF-8

    .PARAMETER DefaultConsoleWidth
        Say to leave default console width and not grow its to $CONSOLE_WIDTH

    .PARAMETER Verbose
        Enable verbose messages

    .EXAMPLE 
        powershell.exe -NoProfile -ExecutionPolicy "RemoteSigned" -File "swm.ps1" -Action "Discovery" -ObjectType "Device"

        Description
        -----------  
        Make Zabbix's LLD JSON for iSCSI Devices.

    .EXAMPLE 
        ... "swm.ps1" -Action "Avg" -ObjectType "Target" -Key "PerformanceData.ReadBandwidth" -Id "0x00000000004601D0" -TimePeriod "60"

        Description
        -----------  
        Return average write bandwith for Target with ID=0x00000000004601D0 calculated for last hour

#>


Param (
   [Parameter(Mandatory = $False)] 
   [ValidateSet('Discovery', 'Get', 'Count', 'Avg', 'Last', 'Max', 'Min')]
   [String]$Action,

   [Parameter(Mandatory = $False)]
   [ValidateSet('Server', 'Target', 'Device')]
   [Alias('Object')]
   [String]$ObjectType,

   [Parameter(Mandatory = $False)]
   [String]$Key,

   [Parameter(Mandatory = $False)]
   [String]$Id,

   [Parameter(Mandatory = $False)]
   [Int]$TimePeriod,

   [Parameter(Mandatory = $False)]
   [String]$ErrorCode,

   [Parameter(Mandatory = $False)]
   [String]$ConsoleCP,

   [Parameter(Mandatory = $False)]
   [Switch]$DefaultConsoleWidth
)

#Set-StrictMode –Version Latest

# Set US locale to properly formatting float numbers while converting to string
[System.Threading.Thread]::CurrentThread.CurrentCulture = "en-US"

# Width of console to stop breaking JSON lines
Set-Variable -Name "CONSOLE_WIDTH" -Value 255 -Option Constant

Add-Type -TypeDefinition "public enum SwPerfCounterType { CpuAndRam = 0, Bandwidth, Iops }";
Add-Type -TypeDefinition "public enum SwPerfTimeInterval { LastHour = 0, LastDay }";

####################################################################################################################################
#
#                                                  Function block
#    
####################################################################################################################################
#
#  Select object with Property that equal Value if its given or with Any Property in another case
#
Function PropertyEqualOrAny {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject,
      [PSObject]$Property,
      [PSObject]$Value
   );
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
         # IsNullorEmpty used because !$Value give a erong result with $Value = 0 (True).
         # But 0 may be right ID  
         If (($Object.$Property -Eq $Value) -Or ([string]::IsNullorEmpty($Value))) { $Object }
      }
   } 
}

#
#  Prepare string to using with Zabbix 
#
Function PrepareTo-Zabbix {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject,
      [String]$ErrorCode,
      [Switch]$NoEscape,
      [Switch]$JSONCompatible
   );
   Begin {
      # Add here more symbols to escaping if you need
      $EscapedSymbols = @('\', '"');
      $UnixEpoch = Get-Date -Date "01/01/1970";
   }
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
         If ($Null -Eq $Object) {
           # Put empty string or $ErrorCode to output  
           If ($ErrorCode) { $ErrorCode } Else { "" }
           Continue;
         }
         # Need add doublequote around string for other objects when JSON compatible output requested?
         $DoQuote = $False;
         Switch (($Object.GetType()).FullName) {
            'System.Boolean'  { $Object = [int]$Object; }
            'System.DateTime' { $Object = (New-TimeSpan -Start $UnixEpoch -End $Object).TotalSeconds; }
            Default           { $DoQuote = $True; }
         }
         # Normalize String object
         $Object = $( If ($JSONCompatible) { $Object.ToString() } else { $Object | Out-String }).Trim();
         
         If (!$NoEscape) { 
            ForEach ($Symbol in $EscapedSymbols) { 
               $Object = $Object.Replace($Symbol, "\$Symbol");
            }
         }

         # Doublequote object if adherence to JSON standart requested
         If ($JSONCompatible -And $DoQuote) { 
            "`"$Object`"";
         } else {
            $Object;
         }
      }
   }
}

#
#  Convert incoming object's content to UTF-8
#
Function ConvertTo-Encoding ([String]$From, [String]$To){  
   Begin   {  
      $encFrom = [System.Text.Encoding]::GetEncoding($from)  
      $encTo = [System.Text.Encoding]::GetEncoding($to)  
   }  
   Process {  
      $bytes = $encTo.GetBytes($_)  
      $bytes = [System.Text.Encoding]::Convert($encFrom, $encTo, $bytes)  
      $encTo.GetString($bytes)  
   }  
}

#
#  Make & return JSON, due PoSh 2.0 haven't Covert-ToJSON
#
Function Make-JSON {
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject, 
      [array]$ObjectProperties, 
      [Switch]$Pretty
   ); 
   Begin   {
      [String]$Result = "";
      # Pretty json contain spaces, tabs and new-lines
      If ($Pretty) { $CRLF = "`n"; $Tab = "    "; $Space = " "; } Else { $CRLF = $Tab = $Space = ""; }
      # Init JSON-string $InObject
      $Result += "{$CRLF$Space`"data`":[$CRLF";
      # Take each Item from $InObject, get Properties that equal $ObjectProperties items and make JSON from its
      $itFirstObject = $True;
   } 
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) {
         # Skip object when its $Null
         If ($Null -Eq $Object) { Continue; }

         If (-Not $itFirstObject) { $Result += ",$CRLF"; }
         $itFirstObject=$False;
         $Result += "$Tab$Tab{$Space"; 
         $itFirstProperty = $True;
         # Process properties. No comma printed after last item
         ForEach ($Property in $ObjectProperties) {
            If (-Not $itFirstProperty) { $Result += ",$Space" }
            $itFirstProperty = $False;
            $Result += "`"{#$Property}`":$Space$(PrepareTo-Zabbix -InputObject $Object.$Property -JSONCompatible)";
         }
         # No comma printed after last string
         $Result += "$Space}";
      }
   }
   End {
      # Finalize and return JSON
      "$Result$CRLF$Tab]$CRLF}";
   }
}

#
#  Return value of object's metric defined by key-chain from $Keys Array
#
Function Get-Metric { 
   Param (
      [Parameter(ValueFromPipeline = $True)] 
      [PSObject]$InputObject, 
      [Array]$Keys
   ); 
   Process {
      # Do something with all objects (non-pipelined input case)  
      ForEach ($Object in $InputObject) { 
        If ($Null -Eq $Object) { Continue; }
        # Expand all metrics related to keys contained in array step by step
        ForEach ($Key in $Keys) {              
           If ($Key) {
              $Object = Select-Object -InputObject $Object -ExpandProperty $Key -ErrorAction SilentlyContinue;
              If ($Error) { Break; }
           }
        }
        $Object;
      }
   }
}

#
#  Exit with specified ErrorCode or Warning message
#
Function Exit-WithMessage { 
   Param (
      [Parameter(Mandatory = $True, ValueFromPipeline = $True)] 
      [String]$Message, 
      [String]$ErrorCode 
   ); 
   If ($ErrorCode) { 
      $ErrorCode;
   } Else {
      Write-Warning ($Message);
   }
   Exit;
}

function New-SWServer {
   Param (
      [String]$SWHost = '127.0.0.1',
      [Int]$SWPort = 3261,
      [String]$Username = 'root',
      [String]$Password = 'starwind'
   )
   $StarWindX = New-Object -ComObject StarWindX.StarWindX;
   $Server = $StarWindX.CreateServer($SWHost, $SWPort);
   $AuthInfo = $Server.AuthentificationInfo;
   $AuthInfo.Login = $Username;
   $AuthInfo.Password = $Password;
   $AuthInfo.IsChap = $False;
   
   $Server;
}

####################################################################################################################################
#
#                                                 Main code block
#    
####################################################################################################################################

Write-Verbose "$(Get-Date) Creating COM-object"
# Connect to local server with default props
$SWServer = New-SWServer;

Write-Verbose "$(Get-Date) Trying to connect to local StarWind Server"
Try {
   $SWServer.Connect()
   If ($SWServer.Connected) {
      Write-Verbose "$(Get-Date) Connection established";
   }
}
Catch {
#   $server.Disconnect();
   Exit-WithMessage -Message $_.Exception.Message -ErrorCode $ErrorCode;
}

# split key to subkeys
$Keys = $Key.Split(".");
$Now = Get-Date;
Write-Verbose "$(Get-Date) Creating collection of specified object: '$ObjectType'";
Switch ($ObjectType) {
   'Server' { 
      $Objects = $SWServer; 
#      Add-Member -Force -InputObject $Objects -MemberType NoteProperty -Name "FullVersion" -Value $Objects.Version.ToString();
   }
   'Target' {
      $Objects = PropertyEqualOrAny -InputObject $SWServer.Targets -Property ID -Value $Id;
      # Add initiators arrays to target
      # May be that procedure need to go to separate 'Key'-case block?
      Write-Verbose "$(Get-Date) Initiators";
      ForEach ($Target in $Objects) {
        Add-Member -InputObject $Target -MemberType 'NoteProperty' -Name 'Initiators' -Value @($SWServer.GetInitiators($Target.Id));
      }
 
   }
   'Device' { 
      $Objects = PropertyEqualOrAny -InputObject $SWServer.Devices -Property DeviceID -Value $Id;
   }
}  


#$Objects | fl *;
#exit;

Write-Verbose "$(Get-Date) Analyzing key";
$Objects = $( 
   Switch ($Keys[0]) {
      'Initiator' {
         Write-Verbose "$(Get-Date) 'Initiator' key detected";
         # Non-optimal code. need to rework
         If ('Target' -Eq $ObjectType) {
            ForEach ($Object in $Objects) { $Object.Initiators; }
         } Else {
           $Null;
         }
         $Keys[0] = '';
      }
      'PerformanceData' {
         Write-Verbose "$(Get-Date) 'PerformanceData' key detected";

         $PerfCounterType = $(
             If ('CPU' -Eq $Keys[1] -OR 'RAM' -Eq $Keys[1]) {
               Write-Verbose "$(Get-Date) Do 'CpuAndRam' query";
               [SwPerfCounterType]::CPUandRAM;
             } ElseIf ('WriteBandwidth' -Eq $Keys[1] -OR 'ReadBandwidth' -Eq $Keys[1] -OR 'TotalBandwidth' -Eq $Keys[1]) {
               Write-Verbose "$(Get-Date) Do 'Bandwidth' query";
               [SwPerfCounterType]::Bandwidth;
             } ElseIf ('TotalIOPs' -Eq $Keys[1]) {
               Write-Verbose "$(Get-Date) Do 'IOps' query";
               [SwPerfCounterType]::IOPs;
             } Else {
              #Error   
             }
         );

         $PerfTimeInterval = [SwPerfTimeInterval]::LastHour;
         $TimePeriod = $( If ('Last' -Eq $Action) { 0 } ElseIf ($TimePeriod -As [Int]) { $TimePeriod } Else { 1 });
         $LastMin = $Now.AddMinutes( 0-$TimePeriod );
         Write-Verbose "$(Get-Date) Take '$ObjectType's '$Property' PerfData for $TimePeriod mins ($LastMin ... $Now)";
         ForEach ($Object in $Objects) { 
            If ($Null -Eq $Object) { Continue; }
            Switch ($ObjectType) {
               'Server' { $PerformanceData = $SWServer.QueryServerPerformanceData($PerfCounterType, $PerfTimeInterval); }
               'Target' { $PerformanceData = $SWServer.QueryTargetPerformanceData($PerfCounterType, $Object.Name, $PerfTimeInterval); }
               'Device' { $PerformanceData = $SWServer.QueryDevicePerformanceData($PerfCounterType, $Object.Name, $PerfTimeInterval); }
            }
            If (0 -Eq $TimePeriod ) {
               $PerformanceData = $PerformanceData | Select-Object -Last 1
            } Else {
               $PerformanceData = $PerformanceData | ? { $_.DateTime -gt $LastMin };
            }
            $PerformanceData
         }
         $Keys[0] = '';
      }
      Default { $Objects; } 
   }
);


#$Objects | fl *;
#exit;

Write-Verbose "$(Get-Date) Collection created, begin processing its with action: '$Action'";
$Result = $(
   If (('Discovery' -Ne $Action -OR 'Get' -Ne $Action) -And (-Not $Objects)) { 
     Write-Verbose "$(Get-Date) No objects in collection - all calculation functions must return 0";
     0; Break;
   } 
   Switch ($Action) {
      # Discovery given object, make json for zabbix
      'Discovery' {
          Switch ($ObjectType) {
             'Target' { $ObjectProperties = @("ALIAS", "NAME", "ID");  }
             'Device' { $ObjectProperties = @("NAME", "DEVICETYPE", "DEVICEID", "TARGETID", "EXISTS");  }
          }
          Write-Verbose "$(Get-Date) Generating LLD JSON";
          Make-JSON -InputObject $Objects -ObjectProperties $ObjectProperties -Pretty;
      }
      # Get metrics or metric list
      'Get' {
         If ($Null -Eq $Objects) {
            Exit-WithMessage -Message "No objects in collection" -ErrorCode $ErrorCode;
         }
         If ($Keys) { 
            Write-Verbose "$(Get-Date) Getting metric related to key: '$Key'";
            PrepareTo-Zabbix -InputObject (Get-Metric -InputObject $Objects -Keys $Keys) -ErrorCode $ErrorCode;
         } Else { 
            Write-Verbose "$(Get-Date) Getting metric list due metric's Key not specified";
            Out-String -InputObject $Objects;
         };
      }
      # Count selected objects
      'Count' { 
         Write-Verbose "$(Get-Date) Counting objects";  
         # if result not null, False or 0 - return .Count
         @($Objects).Count;
      }
      'Last' {
         Write-Verbose "$(Get-Date) Take last value";  
         Get-Metric -InputObject $Objects -Keys $Keys;
      }
      'Avg' { 
         Write-Verbose "$(Get-Date) Calculate average of objects metrics";  
         $Result = 0;
         ForEach ($Object in $Objects) {
            $Result += Get-Metric -InputObject $Object -Keys $Keys;
         }
         $Result / @($Objects).Count;
      }
      'Max' { 
         Write-Verbose "$(Get-Date) Find max value of objects metrics";  
         $Result = 0;
         ForEach ($Object in $Objects) {
            $NextValue = Get-Metric -InputObject $Object -Keys $Keys;
            If ($Result -Lt $NextValue) { $Result = $NextValue; } 
         }
         $Result;
      }
      'Min' { 
         Write-Verbose "$(Get-Date) Find min value of objects metrics";  
         $FirstItem = $True;
         ForEach ($Object in $Objects) {
            $NextValue = Get-Metric -InputObject $Object -Keys $Keys;
            If ($FirstItem -OR $Result -Gt $NextValue) { $Result = $NextValue; } 
            $FirstItem = $False;
         }
         $Result;
      }
   }  
);

# Convert string to UTF-8 if need (For Zabbix LLD-JSON with Cyrillic chars for example)
if ($consoleCP) { 
   Write-Verbose "$(Get-Date) Converting output data to UTF-8";
   $Result = $Result | ConvertTo-Encoding -From $consoleCP -To UTF-8; 
}

# Break lines on console output fix - buffer format to 255 chars width lines 
if (!$defaultConsoleWidth) { 
   Write-Verbose "$(Get-Date) Changing console width to $CONSOLE_WIDTH";
   mode con cols=$CONSOLE_WIDTH; 
}

Write-Verbose "$(Get-Date) Finishing";

$Result;
