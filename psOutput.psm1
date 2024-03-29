﻿##
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
        Consistent Header output formatting

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
		[Parameter(Position = 0, Mandatory = $False, ValueFromPipeline = $true)]
		[Object]
		$Text = (Get-Date).ToLongDateString() + " - " + (Get-Date).ToLongTimeString(),
		[Parameter(Position = 1, Mandatory = $False)]
		[Char]
		$Char = " ",
		[Parameter(Position = 2, Mandatory = $False)]
		[System.ConsoleColor]
		$Fore = (DefaultHeadingForeground),
		[Parameter(Position = 3, Mandatory = $False)]
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
				Write-Host $newline.PadRight($ScreenWidth, $Char) -ForegroundColor $Fore -BackgroundColor $Back
			}
		}
	}
	END { }
}

new-alias -Name heading -Value Write-Heading -Description "Consistent Header output formatting" -force

function Write-ItemName {
	<#
    .SYNOPSIS
        Consistent Item Naming output formatting

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
	param(
		[Parameter(Position = 0)]
		[System.String]
		$Text = (Get-Date).ToLongDateString() + " - " + (Get-Date).ToLongTimeString(),
		[Parameter(Position = 1)]
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
        Consistent Item output formatting

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
		[Parameter(Position = 0)]
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
	param(
		[Parameter(Position = 0)]
		[System.String]
		$Character = "─",
		[Parameter(Position = 1)]
		[System.Int16]
		$Width = (($Host.UI.RawUI.WindowSize.Width) - 1),
		[switch] $Rainbow = $false
	)

	#set the line width based on console width divided by number of chars in the string
	#so we can set multi-character repeats... ex. xxxxx or x x x x
	#BUT not wrap to the next line
	$Width = (($Width) / [Math]::Floor([Decimal]($Character.Length)))

	#Let's keep it all on one line, please.
	while ( (($width) * ($Character.Length)) -gt (($Host.UI.RawUI.WindowSize.Width) - 1)) { $width -= 1 }

	$allColors = ("-Black", "-DarkBlue", "-DarkGreen", "-DarkCyan", "-DarkRed", "-DarkMagenta", "-DarkYellow", "-Gray",
		"-Darkgray", "-Blue", "-Green", "-Cyan", "-Red", "-Magenta", "-Yellow", "-White",
		"-Foreground")

	#output the line
	if (-not $Rainbow) {
		$($Character.ToString() * $Width)
	} else {
		Write-Rainbow $($Character.ToString() * $Width)
	}
}

new-alias -Name repeat -Value Write-Repeating -description "Repeat a character x times" -force

function Write-Center {
	<#
    .SYNOPSIS
        Center text on the screen

    .DESCRIPTION
        Takes the input string and pads it with characters (default is spaces) so the text fits in the approximate middle of the console.

    .PARAMETER  Text
        Type in the message you want here

    .PARAMETER  Char
        The spacer character (defaults to space)

    .PARAMETER Width
        How many characters wide the centerting should be; console-width is default

    .EXAMPLE
        PS C:\> write-center -Message "Testing the function"

    .EXAMPLE
        PS C:\> write-center -Message "Testing the function" -char -
#>
	[CmdletBinding()]
	param(
		[Parameter(Position = 0, Mandatory = $True, ValueFromPipeline = $true)]
		[object] $Text,
		[Parameter(Position = 1, Mandatory = $False)]
		[Char] $Char = " ",
		[Parameter(Position = 2, Mandatory = $False)]
		[alias("Width")]
		[int] $ScreenWidth = $Host.UI.RawUI.WindowSize.Width
	)
	BEGIN {
		$ScreenWidth -= 1
		$retval = ""
	}
	PROCESS {
		foreach ($line in $text) {
			$hold = WrapTheLines $line $ScreenWidth
			#And pad the text on the left and right for the background color change
			foreach ($newline in $hold.split("`n")) {
				$newline.PadRight((($ScreenWidth / 2) + ($newline.Length / 2)), $char).PadLeft($ScreenWidth, $char)
			}
		}
	}
	END { }
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

    .PARAMETER Width
        How many characters wide the centerting should be; console-width is default

    .EXAMPLE
        PS C:\> write-right -Message "Testing the function"

    .EXAMPLE
        PS C:\> write-right -Message "Testing the function" -char -
#>
	[CmdletBinding()]
	param(
		[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
		[object]$Text,
		[Parameter(Position = 1, Mandatory = $False)]
		[Char] $Char = " ",
		[Parameter(Position = 2, Mandatory = $False)]
		[alias("Width")]
		[int] $ScreenWidth = $Host.UI.RawUI.WindowSize.Width
	)
	BEGIN {
		#Lock down the screen size for use later
		$ScreenWidth -= 1
		$retval = ""
	}
	PROCESS {
		foreach ($line in $text) {
			$hold = WrapTheLines $line $ScreenWidth
			#And pad the text on the left
			foreach ($newline in $hold.split("`n")) {
				$newline.ToString().PadLeft(($ScreenWidth), $char)
			}
		}
	}
	END { }
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
    .Parameter Fil
      The File generating the logging event.
    .Parameter Thread
      The Thread ID of the thread generating the logging event.
    .Parameter Context
    .Parameter FilePath
      The path to the log file to be generated/written to. This function will use an "overall"
      variable called "WRITELOGFILEPATH" and uses whatever path is there. This variable can be
      set in the script prior to calling this cmdlet. Alternatively a path to a file may be
      provided. If not set, this will output to screen.
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
		[string] $Component = "",
		[string] $File = "",
		[string] $Thread = "",
		[string] $Context = "",
		[string] $FilePath = $WRITELOGFILEPATH,
		[ValidateSet(1, 2, 3)]
		[int] $Type = 1
	)

	begin {
		[string] $TZBias = (Get-WmiObject -Query "Select Bias from Win32_TimeZone").bias
		if (-not ($TZBias.StartsWith("-") -or $TZBias.StartsWith("+"))) {
			$TZBias = "+" + $TZBias.ToString()
		}
	}

	process {
		$Time = Get-Date -Format "HH:mm:ss.fff"
		$Date = Get-Date -Format "MM-dd-yyyy"

		$Output = "<![LOG[$($Message)]LOG]!><time=`"$($Time)+$($TZBias)`" date=`"$($Date)`" "
		$Output += "component=`"$($Component)`" context=`"$($Context)`" type=`"$($Type)`" "
		$Output += "thread=`"$($Thread)`" file=`"$($File)`">"

		Write-Verbose "$Time $Date`t$Message"
		if ($FilePath) {
			Out-File -InputObject $Output -Append -NoClobber -Encoding Default -FilePath $FilePath
		} else {
			switch ($type) {
				2 { $color = [System.ConsoleColor]::Yellow; break }
				3 { $color = [System.ConsoleColor]::Red; break }
				Default { $color = [System.ConsoleColor]::White; break }
			}
			write-host $Output -ForegroundColor $color -BackgroundColor Black
		}
	}
}
New-Alias -name wt -value Write-Trace -Description "Write output in trace32 format" -Force

function Write-Rainbow {
	param (
		$Text = 'You should supply a text parameter'
	)
	$bgcolor = $host.ui.RawUI.BackgroundColor.ToString()
	[System.Collections.ArrayList] $allColors = ("Black", "DarkBlue", "DarkGreen", "DarkCyan", "DarkRed", "DarkMagenta", "DarkYellow", "Gray",
		"Darkgray", "Blue", "Green", "Cyan", "Red", "Magenta", "Yellow", "White")
	$allColors.Remove($bgcolor)
	$a = $text.ToCharArray()
	$a | ForEach-Object {
		Write-Host $_ -NoNewline -ForegroundColor $($allColors | get-random)
	}
	write-host ''

}

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

	$allColors = ("-Black", "-DarkBlue", "-DarkGreen", "-DarkCyan", "-DarkRed", "-DarkMagenta", "-DarkYellow", "-Gray",
		"-Darkgray", "-Blue", "-Green", "-Cyan", "-Red", "-Magenta", "-Yellow", "-White",
		"-Foreground")

	$color = "Foreground"
	$nonewline = $false

	foreach ($arg in $args) {
		if ($arg -eq "-nonewline") {
			$nonewline = $true
		} elseif ($allColors -contains $arg) {
			$color = $arg.substring(1)
		} else {
			if ($color -eq "Foreground") {
				Write-Host $arg -nonewline
			} else {
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

    .PARAMETER  Char
        The spacer character (defaults to space).

    .EXAMPLE
        PS C:\> Write-Flag

        Outputs a screen-wide box with the current time/date

    .EXAMPLE
        PS C:\> Write-Flag "This is complete"

        Outputs a screen-wide box with the message "This is complete"

    .INPUTS
        System.String
    #>
	param (
		[Parameter(Position = 0, Mandatory = $false, ValueFromPipeline = $True)]
		[System.String]
		$Text = (Get-Date).ToString("dddd -- MMMM d, yyyy -- h:mmtt"),
		[Parameter(Position = 1, Mandatory = $False)]
		[Char] $Char = " "
	)
	BEGIN {
		#Write the opening line/separator
		Write-Repeating
	}
	PROCESS {
		#Write each line (center takes care of lines longer than screenwidth)
		write-center $text -char $Char
	}
	END {
		#Write in the closing line/separator
		Write-Repeating
	}
}

new-alias -Name flag -Value Write-Flag -Description "Date- or Text-box marker" -force

function Write-Box {
	<#
    .SYNOPSIS
        Date- or Text-Box

    .DESCRIPTION
        Writes a box - either with the current time/date or with custom text. Useful to call attention to a specific item on the screen.

        See https://en.wikipedia.org/wiki/Box_Drawing for box shapes

    .PARAMETER  Text
        Type in the text you want here. An empty string will use the current time/date

    .PARAMETER  Char
        The spacer character (defaults to space)

    .PARAMETER Left
        Left-Justifies the text in the box (default)

    .PARAMETER Right
        Right-Justifies the text in the box

    .PARAMETER Center
        Center-Justifies the text in the box

    .EXAMPLE
        PS C:\> Write-Box

        Outputs a text-wide box with the current time/date

    .EXAMPLE
        PS C:\> Write-Box "This is complete"

        Outputs a text-wide box with the message "This is complete"

    .INPUTS
        System.String
    #>
	[CmdletBinding(DefaultParameterSetName = "Default")]
	param (
		[Parameter(Position = 0, Mandatory = $false, ValueFromPipeline = $true)]
		[System.String]
		$Text = (Get-Date).ToString("dddd -- MMMM d, yyyy -- h:mmtt"),
		[Parameter(Position = 1, Mandatory = $False)]
		[Char] $Char = " ",
		[Parameter(Mandatory = $false, ParameterSetName = "Default")]
		[Switch]
		$Left,
		[Parameter(Mandatory = $false, ParameterSetName = "Right")]
		[Switch]
		$Right,
		[Parameter(Mandatory = $false, ParameterSetName = "Center")]
		[Switch]
		$Center,
		[switch] $Rainbow = $false
	)
	BEGIN {
		$longest = 0
		#Lock down the screen size for use later
		$ScreenWidth = ($Host.UI.RawUI.WindowSize.Width - 5)
		$AllTheLines = @()
		$retval = @()
	}
	PROCESS {
		#I want the box around the whole thing - not lots of little boxes
		#So, I need to process both command line parameters as well as the pipeline
		#cmd-line params are easy. Pipeline is easy. Both together ... not so much
		#I've looked around; I've tried various things.
		#Nothing works as well as abusing the correct form of advanced functions.
		#I'm sorry - using $input in PROCESS works for pipeline, but breaks cmd-line params
		#The "best" option I could find was using process to build a "total"
		#and breaking it out in END. If someone has a better way, please let me know!!
		foreach ($line in $Text) {
			#First get the length to see if it's the longest line (adjusting the tab char along the way)
			$trimmed = $line.ToString().Replace("`t", '    ')
			if ($trimmed.Length -gt $longest) { $longest = $trimmed.Length }
			##Then add the line to the combination var
			$AllTheLines += $trimmed
		}
	}
	END {
		#Check if the longest line is bigger than the screen - and limit it as necessary
		if ($longest -gt $ScreenWidth) { $longest = $ScreenWidth }
		#Draw the top of the box
		$retval += "┌$(Write-Repeating -width ($longest + 2))┐"
		#And then process each line in the collection for output
		$retval += foreach ($line in $AllTheLines) {
			#BUT, break up any line longer than the screenwidth
			$hold = WrapTheLines $line $ScreenWidth
			#And then actually process the lines for output
			foreach ($line in $hold.ToString().Split("`n")) {
				#Box shape is a bar, a space, the text (padded to "longest" width, a space, and a bar
				if ($Center) {
					"│$($Char)$(write-Center $line.ToString() -width ($longest + 1) -char $Char)$Char│"
				} elseIf ($Right) {
					"│$($Char)$(write-Right $line.ToString() -width ($longest + 1) -char $Char)$Char│"
				} else {
					"│$($Char)$($line.ToString().PadRight($longest, $Char))$Char│"
				}
			}
		}
		#Now draw the bottom of the box
		$retval += "└$(Write-Repeating -width ($longest + 2))┘"
		if (-not $Rainbow) {
			$retval
		} else {
			$retval | ForEach-Object { Write-Rainbow $_ }
		}
	}
}

new-alias -Name box -Value Write-Box -Description "Date- or Text-box marker" -force

function Write-Reverse {
	<#
    .SYNOPSIS
        Reverse the text

    .DESCRIPTION
        Putting the fun in function: takes the input texts and prints it backwards (well, reorders the letters so they are in the reverse order - it doesn't print backwards characters, unfortunately). Prints the date backwards by default.

    .PARAMETER  Text
        Type in the text you want here. An empty string will use the current time/date

    .EXAMPLE
        PS C:\> Write-Reverse

        MA32:01 -- 6102 ,82 hcraM -- yadnoM

    .EXAMPLE
        PS C:\> Write-Reverse | Write-Reverse

        Monday -- March 28, 2016 -- 10:24AM

    .EXAMPLE
        PS C:\> get-content .\Test.txt | Write-Reverse

        elif tset a si sihT
        uoy t'nod ,ti evol tsuj uoY

    .INPUTS
        System.String
    #>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false, ValueFromPipeline = $true)]
		[System.String]
		$Text = (Get-Date).ToString("dddd -- MMMM d, yyyy -- h:mmtt")
	)
	BEGIN {
		$ScreenWidth = ($Host.UI.RawUI.WindowSize.Width - 1)
	}
	PROCESS {
		ForEach ($line in $Text) {
			$hold = WrapTheLines $line $ScreenWidth
			ForEach ($newline in $hold.Split("`n")) {
				$newline[$newline.ToString().Length..0] -join ""
			}
		}
	}
	END { }
}

new-alias -name  reverse -Value Write-Reverse -Description "Reverse text" -force

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
		[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
		[System.String]
		$Text,
		[Parameter(Position = 1)]
		[System.String]
		[validateset('Male', 'Female')]
		$Gender = 'Female'
	)
	begin {
		try {
			Add-Type -Assembly System.Speech -ErrorAction Stop
		} catch {
			Write-Error -Message "Error loading the requered assemblies"
		}
	}
	process {
		$voice = New-Object -TypeName 'System.Speech.Synthesis.SpeechSynthesizer' -ErrorAction Stop

		Write-Verbose "Selecting a $Gender voice"
		$voice.SelectVoiceByHints($Gender)

		Write-Verbose -Message "Start Speaking"
		[void]$voice.Speak($Text)
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

        See http://www.bgreco.net/powershell/format-color/ for more info

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
	$lines = ($input | Out-String) -replace "`r", "" -split "`n"
	ForEach ($line in $lines) {
		$color = ''
		ForEach ($pattern in $Colors.Keys) {
			if (!$SimpleMatch -and $line -match $pattern) { $color = $Colors[$pattern] }
			elseif ($SimpleMatch -and $line -like $pattern) { $color = $Colors[$pattern] }
		}
		if ($color) {
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
