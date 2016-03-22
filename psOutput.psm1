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
    [CmdletBinding()]
	param( 
        [Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$true)] 
        [Object] 
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
    BEGIN {
        $retval = ""
        $ScreenWidth = ($Host.UI.RawUI.WindowSize.Width - 1)
    }
    PROCESS {
        foreach ($line in $text) {
            $hold = WrapTheLines $line $ScreenWidth
            #And pad the text on the right for the background color change
            foreach ($newline in $hold.split("`n")) {
                $retval += $newline.PadRight($ScreenWidth, $Char)
                $retval += "`n"
            }
        }
    }
    END {
        foreach ($line in $retval.TrimEnd("`n").split("`n")) {
             Write-Host $line -ForegroundColor $Fore -BackgroundColor $Back
        }
    }
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

new-alias -Name repeat -Value Write-Repeating -description "Repeat a character x times" -force

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
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true)]
        [object] $Text,
        [Parameter(Position=1,Mandatory=$False)]
        [Char] $Char = " "
    )
    BEGIN { 
        $ScreenWidth = ($Host.UI.RawUI.WindowSize.Width - 1)
        $retval = "" 
    }
    PROCESS {
        foreach ($line in $text) {
            $hold = WrapTheLines $line $ScreenWidth
            #And pad the text on the left and right for the background color change
            foreach ($newline in $hold.split("`n")) {
                $retval += $newline.PadRight((($ScreenWidth / 2) + ($newline.Length / 2)), $char).PadLeft($ScreenWidth, $char)
                $retval += "`n"
            }
        }
    }
    END {
        try { 
            #return what we put together and remove any extra, trailing newline
            #could be empty, so try it, but do nothing if it errors.
            return $retval.TrimEnd("`n")
        } catch {}
    }
}

new-alias -Name center -Value Write-Center -description "Center text on the screen" -force

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

    #Cmdlet binding for all the reasons
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
        [object]$Text,
        [Parameter(Position=1,Mandatory=$False)]
        [Char] $Char = " "
    )
    BEGIN { 
        #Lock down the screen size for use later
        $ScreenWidth = ($Host.UI.RawUI.WindowSize.Width - 1)
        $retval = ""
    }
    PROCESS {

        foreach ($line in $text) {
            $hold = WrapTheLines $line $ScreenWidth
            #And pad the text on the left
            foreach ($newline in $hold.split("`n")) {
                $retval += $newline.ToString().PadLeft(($ScreenWidth), $char)
                $retval += "`n"
            }
        }
    }
    END {
        try { 
            #return what we put together and remove any extra, trailing newline
            #could be empty, so try it, but do nothing if it errors.
            return $retval.TrimEnd("`n")
        } catch {}
    }
}

new-alias -Name right -Value Write-Right -description "Right-justify text" -force

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

function Write-ColorText {
  <#
    .Synopsis
      Write in color
    .Description
      A simple function to replace Write-Host with more "in-line" color options. Write-Host with -nonewline works fine and dandy, but makes for long lines. This way we can specify a line of text with mixed color in 1 command.

      Originally found at https://stackoverflow.com/questions/2688547/muliple-foreground-colors-in-powershell-in-one-command

      Note: We don't have parameters specified in the function - we parse all params for the specific ones we want (it's actually simpler this way). 

      Color options: 
        -Black
        -DarkBlue
        -DarkGreen
        -DarkCyan
        -DarkRed
        -DarkMagenta
        -DarkYellow
        -Gray
        -Darkgray
        -Blue
        -Green
        -Cyan
        -Red
        -Magenta
        -Yellow
        -White
        -Foreground (default)

    Command structure:
        Write-ColorText [-color] Text [[-color] [text]...]

            (use quotes if you want text with spaces)

    .Parameter Text
      The text to write; use quotes to include spaces
    .Parameter -Color
      The color in which to write it. Can be -Black, -DarkBlue, -DarkGreen, -DarkCyan, -DarkRed, -DarkMagenta, -DarkYellow, -Gray, -Darkgray, -Blue, -Green, -Cyan, -Red, -Magenta, -Yellow, -White, -Foreground
    .Example
      Write-ColorText "This is normal text"

      Outputs "This is normal text" in the current foreground color
    .Example
      Write-ColorText Normal -Red Red -White White -Blue Blue -ForeGround Normal

      Outputs "Normal" in the current foreground color, plus each color word in that color
  #>

    $allColors = ("-Black",   "-DarkBlue","-DarkGreen","-DarkCyan","-DarkRed","-DarkMagenta","-DarkYellow","-Gray",
                  "-Darkgray","-Blue",    "-Green",    "-Cyan",    "-Red",    "-Magenta",    "-Yellow",    "-White",
                   "-Foreground")
    
    $color = "Foreground"
    $nonewline = $false

    foreach($arg in $args)
    {
        if ($arg -eq "-nonewline")
        { 
            $nonewline = $true 
        }
        elseif ($allColors -contains $arg)
        {
            $color = $arg.substring(1)
        }
        else
        {
            if ($color -eq "Foreground")
            {
                Write-Host $arg -nonewline
            }
            else
            {
                Write-Host $arg -foreground $color -nonewline
            }
        }
    }

    Write-Host -nonewline:$nonewline
}

Function Write-Flag {
    <# 
    .SYNOPSIS 
        Date- or Text-Box Marker 
 
    .DESCRIPTION 
        Writes a "bright" marker - either with the current time/date or with custom text. Useful to call attention to a specific item on the screen (I use it in scripts to call attention to something, usually completion or someting I might otherwise miss). 
 
    .PARAMETER  Text 
        Type in the text you want here. An empty string will use the current time/date
 
    .PARAMETER  Fore
        Foreground color

    .PARAMETER  Back
        Background color

    .EXAMPLE 
        PS C:\> Write-Flag

        Outputs a screen-wide box with the current time/date
         
    .EXAMPLE 
        PS C:\> Write-Flag "This is complete"

        Outputs a screen-wide box with the message "This is complete"
     
    .EXAMPLE 
        PS C:\> Write-Flag "Done!" -bw

        Outputs a screen-wide box with the message "Done!" in the current foreground and background colors

    .INPUTS 
        System.String, System-ConsoleColor
    #> 

    param (
        [Parameter(Position=0, Mandatory=$false)] 
        [System.String] 
        $Text = (Get-Date).ToString("dddd -- MMMM d, yyyy -- h:mmtt"),
        [Parameter(Position=1,Mandatory=$False)]
        [System.ConsoleColor]
        $Fore = (DefaultHeadingForeground),
        [Parameter(Position=2,Mandatory=$False)]
        [System.ConsoleColor]
        $Back = (DefaultHeadingBackground)
    )
    $retval = Write-Repeating
    $retval += "`n"
    $retval += (write-center $text)
    $retval += "`n"
    $retval += Write-Repeating
    if (($Fore -eq (DefaultHeadingForeground)) -and ($Back = (DefaultHeadingBackground))) {
        $retval
    } else {
        Write-Host $retval -ForegroundColor $Fore -BackgroundColor $Back
    }
}

new-alias -Name flag -Value Write-Flag -Description "Date- or Text-box marker" -force

function Write-Box {
    <# 
    .SYNOPSIS 
        Date- or Text-Box 
 
    .DESCRIPTION 
        Writes a box - either with the current time/date or with custom text. Useful to call attention to a specific item on the screen. By default, the output is an Object that can be piped to other cmdlets or functions (like my write-right function). You can also specify fore- and back-ground colors, but then the output is text and completely non-pipe-able.
 
    .PARAMETER  Text 
        Type in the text you want here. An empty string will use the current time/date

    .PARAMETER  Fore
        Foreground color

    .PARAMETER  Back
        Background color

    .EXAMPLE 
        PS C:\> Write-Box

        Outputs (as an OBJECT) a text-wide box with the current time/date
         
    .EXAMPLE 
        PS C:\> Write-Box "This is complete"

        Outputs (as an OBJECT) a text-wide box with the message "This is complete"
     
    .EXAMPLE 
        PS C:\> Write-Box "Done!" -fore red -back blue

        Outputs (as TEXT) a text-wide box with the message "Done!" as red text on blue background

    .INPUTS 
        System.String, System-ConsoleColor
    #> 
    param (
        [Parameter(Position=0, Mandatory=$false)] 
        [System.String] 
        $Text = (Get-Date).ToString("dddd -- MMMM d, yyyy -- h:mmtt"),
        [Parameter(Position=2,Mandatory=$False)]
        [System.ConsoleColor]
        $Fore = (DefaultHeadingForeground),
        [Parameter(Position=3,Mandatory=$False)]
        [System.ConsoleColor]
        $Back = (DefaultHeadingBackground)
    )
    #See https://en.wikipedia.org/wiki/Box_Drawing for box shapes
    $length = ($Text.Length + 2)
    $retval = "┌"
    $retval += Write-Repeating -width $length
    $retval += "┐"
    $retval += "`n"
    $retval += "│ $text │"
    $retval += "`n"
    $retval += "└"
    $retval += Write-Repeating -width $length
    $retval += "┘"
    if (($Fore -eq (DefaultHeadingForeground)) -and ($Back = (DefaultHeadingBackground))) {
        $retval
    } else {
        Write-Host $retval -ForegroundColor $Fore -BackgroundColor $Back
    }

}

new-alias -Name box -Value Write-Box -Description "Date- or Text-box marker" -force

function Out-Speech { 
    <# 
    .SYNOPSIS 
        Text to Speech 
 
    .DESCRIPTION 
        This is a Text to Speech Function made in powershell. 
 
    .PARAMETER  Text 
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
        $Text, 
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
            $voice.Speak($Text) | Out-Null 
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

############# Conversions ###############

function ConvertFrom-SID
 {
  param([string]$SID="S-1-0-0")
  $objSID = New-Object System.Security.Principal.SecurityIdentifier($SID)
  $objUser = $objSID.Translate([System.Security.Principal.NTAccount])
  Return $objUser.Value
 }

 function ConvertTo-SID
 {
  param([string]$ID="Null SID")
  $objID = New-Object System.Security.Principal.NTAccount($ID)
  $objSID = $objID.Translate([System.Security.Principal.SecurityIdentifier])
  Return $objSID.Value
 }

new-alias -name FromSID -value ConvertFrom-SID -Description "Get UserName from SID" -Force
new-alias -name ToSID -value ConvertTo-SID -Description "Get SID from UserName" -Force

Function ConvertTo-URLEncode([string]$InText="You did not enter any text!") {
    [System.Reflection.Assembly]::LoadWithPartialName("System.web") | out-null
    [System.Web.HttpUtility]::UrlEncode($InText)
}

Function ConvertFrom-URLEncode([string]$InText="You+did+not+enter+any+text!") {
    [System.Reflection.Assembly]::LoadWithPartialName("System.web") | out-null
    [System.Web.HttpUtility]::UrlDecode($InText)
}

New-Alias -name "URLEncode" -Value ConvertTo-URLEncode -Description "URL encode a string" -Force
New-Alias -name "URLDecode" -Value ConvertFrom-URLEncode -Description "URL decode a string" -Force

Function ConvertTo-Fahrenheit([decimal]$celsius) {
    $((1.8 * $celsius) + 32 )
} 

Function ConvertTo-Celsius($fahrenheit) {
    $( (($fahrenheit - 32)/9)*5 )
}

New-Alias -name "ToF" -Value ConvertTo-Fahrenheit -Description "Convert degrees C to F" -Force
New-Alias -name "ToC" -Value ConvertTo-Celsius -Description "Convert degrees F to C" -Force


Function Convert-AddressToName($addr) {
    [system.net.dns]::GetHostByAddress($addr)
}

Function Convert-NameToAddress($addr) {
    [system.net.dns]::GetHostByName($addr)
}

New-Alias -name "n2a" -value Get-NameToAddress -Description "Get IP Address from DNS by Host Name" -Force
New-Alias -name "a2n" -value Get-AddressToName -Description "Get Host Name from DNS by IP Address" -Force

function ConvertFrom-RomanNumeral {
  <#
    .SYNOPSIS
        Converts a Roman numeral to a number.
    .DESCRIPTION
        Converts a Roman numeral - in the range of I..MMMCMXCIX - to a number. Found at https://stackoverflow.com/questions/267399/how-do-you-match-only-valid-roman-numerals-with-a-regular-expression
    .EXAMPLE
        ConvertFrom-RomanNumeral -Numeral MMXIV
    .EXAMPLE
        "MMXIV" | ConvertFrom-RomanNumeral
  #>
    [CmdletBinding()]
    [OutputType([int])]
    Param (
        [Parameter(Mandatory=$true,
                   HelpMessage="Enter a roman numeral in the range I..MMMCMXCIX",
                   ValueFromPipeline=$true,
                   Position=0)]
        [ValidatePattern("^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$")]
        [string] $Numeral
    )
    Begin {
        $RomanToDecimal = [ordered]@{
            M  = 1000
            CM =  900
            D  =  500
            CD =  400
            C  =  100
            XC =   90
            L  =   50
            X  =   10
            IX =    9
            V  =    5
            IV =    4
            I  =    1
        }
    }
    Process {
        $roman = $Numeral + " "
        $value = 0

        do {
            foreach ($key in $RomanToDecimal.Keys) {
                if ($key.Length -eq 1) {
                    if ($key -match $roman.Substring(0,1)) {
                        $value += $RomanToDecimal.$key
                        $roman  = $roman.Substring(1)
                        break
                    }
                }
                else {
                    if ($key -match $roman.Substring(0,2)) {
                        $value += $RomanToDecimal.$key
                        $roman  = $roman.Substring(2)
                        break
                    }
                }
            }
        } until ($roman -eq " ")
        $value
    }
    End {
    }
}

New-Alias -name "FromRoman" -value ConvertFrom-RomanNumeral -Description "Convert from a roman numeral" -Force

function ConvertTo-RomanNumeral {
  <#
    .SYNOPSIS
        Converts a number to a Roman numeral.
    .DESCRIPTION
        Converts a number - in the range of 1 to 3,999 - to a Roman numeral. Found at https://stackoverflow.com/questions/267399/how-do-you-match-only-valid-roman-numerals-with-a-regular-expression
    .EXAMPLE
        ConvertTo-RomanNumeral -Number (Get-Date).Year
    .EXAMPLE
        (Get-Date).Year | ConvertTo-RomanNumeral
  #>
    [CmdletBinding()]
    [OutputType([string])]
    Param (
        [Parameter(Mandatory=$true,
                   HelpMessage="Enter an integer in the range 1 to 3,999",
                   ValueFromPipeline=$true,
                   Position=0)]
        [ValidateRange(1,4999)] [int] $Number
    )
    Begin {
        $DecimalToRoman = @{
            Ones      = "","I","II","III","IV","V","VI","VII","VIII","IX";
            Tens      = "","X","XX","XXX","XL","L","LX","LXX","LXXX","XC";
            Hundreds  = "","C","CC","CCC","CD","D","DC","DCC","DCCC","CM";
            Thousands = "","M","MM","MMM","MMMM"
        }
        $column = @{Thousands = 0; Hundreds = 1; Tens = 2; Ones = 3}
    }
    Process {
        [int[]]$digits = $Number.ToString().PadLeft(4,"0").ToCharArray() | ForEach-Object { [Char]::GetNumericValue($_) }
        $RomanNumeral  = ""
        $RomanNumeral += $DecimalToRoman.Thousands[$digits[$column.Thousands]]
        $RomanNumeral += $DecimalToRoman.Hundreds[$digits[$column.Hundreds]]
        $RomanNumeral += $DecimalToRoman.Tens[$digits[$column.Tens]]
        $RomanNumeral += $DecimalToRoman.Ones[$digits[$column.Ones]]

        $RomanNumeral
    }
    End {  
    }
}

New-Alias -name "ToRoman" -value ConvertTo-RomanNumeral -Description "Convert to a roman numeral" -Force


Export-ModuleMember -function Write-* 
Export-ModuleMember -function Out-* 
Export-ModuleMember -function Format-* 
Export-ModuleMember -function Convert* 
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