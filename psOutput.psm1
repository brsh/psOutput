##
## Leveragd from 
## http://www.the-little-things.net/blog/2015/10/03/powershell-thoughts-on-module-design/
##
#region Private Variables
# Current script path
[string]$ScriptPath = Split-Path (get-variable myinvocation -scope script).value.Mycommand.Definition -Parent
#endregion Private Variables
 
#region Methods
 
# Dot sourcing private script files
Get-ChildItem $ScriptPath/private -Recurse -Filter "*.ps1" -File | Foreach { 
    . $_.FullName
}
 
# Load and export methods

function Write-Heading {
    <# 
    .SYNOPSIS 
        Consistent Header output formatting. 
 
    .DESCRIPTION 
        Writes a "header" to default output (screen). Colors the text and background (full screen width) to highlight the information.
 
    .PARAMETER  Text
        Type in the message you want here. Default (empty string) is the current date/time.
 
    .PARAMETER  Char
        This will be the "spacer" character. Default is space.
 
    .PARAMETER  Fore
        Foreground color. Default is DarkYellow (or White, if the console is DarkYellow).
 
    .PARAMETER  Back
        Background color. Default is Black (or DarkMagenta, if the console is Black)
 
    .EXAMPLE 
        PS C:\> Write-Heading -Text "Testing the function" -Char - -Fore White -Back Red

        Writes "Testing the Function------..." with White text on Red background
    #> 

    #Make a consisten heading
    #Shades the background of the header to the width of the screen
	param( 
        [Parameter(Position=0,Mandatory=$False)] 
        [System.String] 
        $Text = (Get-Date).ToLongDateString() + " - " + (Get-Date).ToLongTimeString(),
        [Parameter(Position=1,Mandatory=$False)]
        [Char]
        $Char = " ",
        [Parameter(Position=2,Mandatory=$False)]
        [System.ConsoleColor]
        $Fore = (DefaultHeadingForeground),
        [Parameter(Position=3,Mandatory=$False)]
        [System.ConsoleColor]
        $Back = (DefaultHeadingBackground)
		)
	
    $Width = (($Host.UI.RawUI.WindowSize.Width) - $Text.Length - 1)

    $Text += Write-Repeating -Character $Char -Width $Width

    Write-Host $Text -ForegroundColor $Fore -BackgroundColor $Back
}

function Write-ItemName {
    <# 
    .SYNOPSIS 
        Consistent Item Naming output formatting. 
 
    .DESCRIPTION 
        Writes a consistent "report item title" to default output (screen), coloring the text and providing x number of tabs (default is 1) without a new line. Meant to be used with Write-Item
 
    .PARAMETER  Text
        Type in the message you want here. Default (empty string) is the current date/time.
 
    .PARAMETER  Tabs
        How many tabs to offset the Item text
 
    .EXAMPLE 
        PS C:\> Write-ItemName -Text "Item:" -tabs 2

        Writes "Item:" with 2 tabs after

    .EXAMPLE 
        PS C:\> Write-ItemName -Text "Item:" -tabs 2; Write-Item "This is an example."

        Writes "Item:        This is an example."
    #> 
  
    #Writes an Item Name 
    #Basically, the label ... think of it as a property or module name of an object
    #Puts a colon and a certain number of tabs for column-ing
	param( 
        [Parameter(Position=0)] 
        [System.String] 
        $Text = (Get-Date).ToLongDateString() + " - " + (Get-Date).ToLongTimeString(),
        [Parameter(Position=1)]
        [Int16]
        $Tabs = 1
		)
	$FG = "White"
    $sTab = (Write-Repeating "`t" $Tabs)
	write-host $Text$sTab -foregroundcolor $FG -NoNewline
}

function Write-Item {
    <# 
    .SYNOPSIS 
        Consistent Item output formatting. 
         
    .DESCRIPTION 
        Writes a consistent "report item" to default output (screen). Meant to be used with Write-ItemName
 
    .PARAMETER  Text
        Type in the message you want here. Default (empty string) is the current date/time.
 
    .EXAMPLE 
        PS C:\> Write-Item -Text "This is boring text."

        Writes "Item:" with 2 tabs after

    .EXAMPLE 
        PS C:\> Write-ItemName -Text "Item:" -tabs 2; Write-Item "This is an example."

        Writes "Item:        This is an example."
    #> 
  	param( 
        [Parameter(Position=0)] 
        [System.String] 
        $Text = (Get-Date).ToLongDateString() + " - " + (Get-Date).ToLongTimeString()
		)
	$FG = "Gray" 
	write-host	$Text -foregroundcolor $FG		
}

function Write-Repeating {
    <# 
    .SYNOPSIS 
        Repeat a character x times
         
    .DESCRIPTION 
        Just a simple function that repeats a character (or series) an x number of times. Default is m-dash (─) for the width of the screen. (An m-dash next to an m-dash seems like a continuous line).
 
    .PARAMETER  Character
        The character(s) to repeat. Note: if you specify a weird number of characters, you might not fill the screen/required space.
 
    .PARAMETER  Width
        How many times to repeat. Default is screen width.

    .EXAMPLE 
        PS C:\> Write-Repeating x 5

        Writes "xxxxx"

    .EXAMPLE 
        PS C:\> Write-Repeating

        Writes "─────────..."
    #>

    #Repeat some characts a certain number of times
    #Defaults to the m-dash at full console width
    #Example: x 6 would display xxxxxx
    param( 
        [Parameter(Position=0)] 
        [System.String] 
        $Character = "─",
        [Parameter(Position=1)] 
        [System.Int16] 
        $Width = (($Host.UI.RawUI.WindowSize.Width) - 1)
    ) 

    #set the line width based on console width divided by number of chars in the string
    #so we can set multi-character repeats... ex. xxxxx or x x x x
    #BUT not wrap to the next line
    $Width = (($Width) / [Math]::Floor([Decimal]($Character.Length)))
    
    #Let's keep it all on one line, please.
    while ( (($width) * ($Character.Length)) -gt (($Host.UI.RawUI.WindowSize.Width) - 1)) { $width -= 1 }

    #output the line
    $($Character.ToString() * $Width)
}

function Write-Center {
    <# 
    .SYNOPSIS 
        Center text on the screen 
 
    .DESCRIPTION 
        Takes the input string and pads it with characters (default is spaces) so the text fits in the approximate middle of the console.
 
    .PARAMETER  Text 
        Type in the message you want here. 

    .PARAMETER  Char 
        The spacer character (defaults to space).
         
    .EXAMPLE 
        PS C:\> write-center -Message "Testing the function" 
         
    .EXAMPLE 
        PS C:\> write-center -Message "Testing the function" -char -
#> 

    #Center text on the console line
    #Can use spaces (default) or other character as the spacer
    param(
        [Parameter(Position=0,Mandatory=$True)]
        [String]
        $Text,
        [Parameter(Position=1,Mandatory=$False)]
        [Char]
        $Char = " "
    )
    #Get the width of the console minus the text and divide in half to get the starting spacer
    $Width = [Math]::Floor(((($Host.UI.RawUI.WindowSize.Width) - 1) - ($Text.Length)) / 2 )
    #Temp var to hold the output
    $retval = (Write-Repeating $Char.ToString() $Width)
    #Make sure we have enough characters to fill the line (add some if necessary)
    if ( ( ( ($retval.length) * 2) + ($Text.Length) ) -lt (($Host.UI.RawUI.WindowSize.Width) - 1)) {
        $retval += $Text + (Write-Repeating $Char.ToString() ($Width + 1))
    }
    else {
        $retval += $Text + $retval
    }
    $retval    
}

function Write-Right {
    <# 
    .SYNOPSIS 
        Right-justify text on the screen 
 
    .DESCRIPTION 
        Takes the input string and pads it with characters (default is spaces) so the text fits to the right side of the console.
 
    .PARAMETER  Text 
        Type in the message you want here. 

    .PARAMETER  Char 
        The spacer character (defaults to space).
         
    .EXAMPLE 
        PS C:\> write-right -Message "Testing the function" 
         
    .EXAMPLE 
        PS C:\> write-right -Message "Testing the function" -char -
#> 

    #Right justify text on the console line
    #Can use spaces (default) or other character as the spacer)
    param(
        [Parameter(Position=0,Mandatory=$True)]
        [String]
        $Text,
        [Parameter(Position=1,Mandatory=$False)]
        [Char]
        $Char = " "
    )
    #get the width of the console minus the test
    $Width = (($Host.UI.RawUI.WindowSize.Width) - ($Text.Length) - 1)

    $retval = (Write-Repeating $Char.ToString() $Width) + $Text

    $retval    
}

function Write-Trace {
  <#
    .Synopsis
      Write a message in Trace32 format
    .Description
      This cmdlet takes a given message and formats that message such that it's compatible with
      the Trace32 log viewer tool used for reading/parsing System Center log files.
      
      The date and time (to the millisecond) is determined at the time that this cmdlet is called.
      Several optional arguments can be provided, to define the Component generating the log
      message, the File that is generating the message, the Thread ID, and the Context under which
      the log entry is being made.
    .Parameter Message
      The actual message to be logged.
    .Parameter Component
      The Component generating the logging event.
    .Parameter File
      The File generating the logging event.
    .Parameter Thread
      The Thread ID of the thread generating the logging event.
    .Parameter Context
    .Parameter FilePath
      The path to the log file to be generated/written to. By default this cmdlet looks for a
      variable called "WRITELOGFILEPATH" and uses whatever path is there. This variable can be
      set in the script prior to calling this cmdlet. Alternatively a path to a file may be
      provided.
    .Parameter Type
      The type of event being logged. Valid values are 1, 2 and 3. Each number corresponds to a 
      message type:
      1 - Normal messsage (default)
      2 - Warning message
      3 - Error message
  #>
  [CmdletBinding()]
  param(
    [Parameter( Mandatory = $true )]
    [string] $Message,
    [string] $Component="",
    [string] $File="",
    [string] $Thread="",
    [string] $Context="",
    [string] $FilePath=$WRITELOGFILEPATH,
    [ValidateSet(1,2,3)]
    [int] $Type=1
  )
  
  begin
  {
    $TZBias = (Get-WmiObject -Query "Select Bias from Win32_TimeZone").bias
  }
  
  process
  {
    $Time = Get-Date -Format "HH:mm:ss.fff"
    $Date = Get-Date -Format "MM-dd-yyyy"
    
    $Output  = "<![LOG[$($Message)]LOG]!><time=`"$($Time)+$($TZBias)`" date=`"$($Date)`" "
    $Output += "component=`"$($Component)`" context=`"$($Context)`" type=`"$($Type)`" "
    $Output += "thread=`"$($Thread)`" file=`"$($File)`">"
    
    Write-Verbose "$Time $Date`t$Message"
    Out-File -InputObject $Output -Append -NoClobber -Encoding Default -FilePath $FilePath
  }
}
New-Alias -name wt -value Write-Trace -Description "Write output in trace32 format" -Force

function Out-Speech { 
    <# 
    .SYNOPSIS 
        Text to Speech 
 
    .DESCRIPTION 
        This is a Text to Speech Function made in powershell. 
 
    .PARAMETER  Message 
        Type in the message you want here. 
 
    .EXAMPLE 
        PS C:\> Out-Speech -Message "Testing the function" -Gender 'Female' 
         
    .EXAMPLE 
        PS C:\> "Testing the function 1","Testing the function 2 ","Testing the function 3","Testing the function 4 ","Testing the function 5 ","Testing the function 6" | Foreach-Object { Out-Speech -Message $_ } 
     
    .EXAMPLE 
        PS C:\> "Testing the Pipeline" | Out-Speech 
 
    .INPUTS 
        System.String 
    #> 
    [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true,ValueFromPipeline=$true)] 
        [System.String] 
        $Message, 
        [Parameter(Position=1)] 
        [System.String] 
        [validateset('Male','Female')] 
        $Gender = 'Female' 
    ) 
    begin { 
        try { 
             Add-Type -Assembly System.Speech -ErrorAction Stop 
        } 
        catch { 
            Write-Error -Message "Error loading the requered assemblies" 
        } 
    } 
    process { 
            $voice = New-Object -TypeName 'System.Speech.Synthesis.SpeechSynthesizer' -ErrorAction Stop 
            
            Write-Verbose "Selecting a $Gender voice" 
            $voice.SelectVoiceByHints($Gender) 
             
            Write-Verbose -Message "Start Speaking" 
            $voice.Speak($message) | Out-Null 
    } 
    end { 
    } 
}

new-alias -name say -value Out-Speech -Description "Have the computer _speak_ the output" -Force

function Format-Color([hashtable] $Colors = @{}, [switch] $SimpleMatch) {
    <# 
    .SYNOPSIS 
        Re-color text output 
 
    .DESCRIPTION 
        Parses the text imput (usually via pipeline) and changes the color of lines based on pattern matching. It can do a simple match (where the entire line matches) or a pattern match (where something keys the color change - like the word error). 
        
        You send patterns and colors in as a hash table - for example @{ 'error' = 'red'; 'info' = 'white' }
         
    .PARAMETER  Colors 
        A hashtable of pattern and color - example: @{ '^Command' = 'white'; '^--' = 'white' }

    .PARAMETER  SimpleMatch 
        A switch to turn off pattern matching and require an exact match (the line of text = the entire text specified)
         
    .EXAMPLE 
        PS C:\> get-alias | format-color @{ '^CommandType' = 'white'; '^--' = 'yellow' }

        Takes the output of the get-alias command and colors the line that starts with CommandType (the ^ means start-of-line) as white, and the line that starts with -- as yellow. Anything else is left as default

    .EXAMPLE 
        PS C:\> get-help Format-Color -Examples | format-color @{ 'error' = 'red' }

        Takes the output of the get-content command and for any line that has the word error, colors that line red

    .EXAMPLE 
        PS C:\> get-help Format-Color -Examples | format-color @{ 'error' = 'red' } -simplematch

        Takes the output of the get-content command and for any line that is ONLY the word error, colors that line red
#> 

    #Just a way to recolor some things that don't have color options
    #can handle regex or simpler matching (like just * to recolor everything)
    #from http://www.bgreco.net/powershell/format-color/
    #pass it a hash table of the form @{'pattern1' = 'Color1'[; ...]}
	$lines = ($input | Out-String) -replace "`r", "" -split "`n"
	ForEach ($line in $lines) {
		$color = ''
		ForEach ($pattern in $Colors.Keys){
			if(!$SimpleMatch -and $line -match $pattern) { $color = $Colors[$pattern] }
			elseif ($SimpleMatch -and $line -like $pattern) { $color = $Colors[$pattern] }
		}
		if($color) {
			Write-Host -ForegroundColor $color $line
		} else {
			Write-Host $line
		}
	}
}

New-Alias -Name clr -value Format-Color -Description "Re-color output text" -Force


Export-ModuleMember -function Write-* 
Export-ModuleMember -function Out-* 
Export-ModuleMember -function Format-* 
Export-ModuleMember -alias *

###################################################
## END - Cleanup
 
#region Module Cleanup
$ExecutionContext.SessionState.Module.OnRemove = {
    # cleanup when unloading module (if any)
    # dir alias: | Where-Object { $_.Source -match "psOutput" } | Remove-Item
    dir function: | Where-Object { $_.Source -match "psOutput" } | Remove-Item
}
#endregion Module Cleanup