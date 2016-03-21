## Helper Functions for psSysInfo Module

function CurrBackground {
    $host.ui.rawui.BackgroundColor
}

function CurrForeground {
    $host.ui.rawui.ForegroundColor
}

function DefaultHeadingBackground {
    if (CurrBackground -match "DarkMagenta") { return "Black" }
    else { return "DarkMagenta" }
}
function DefaultHeadingForeground {
    if (CurrForeground -match "DarkYellow") { return "Yellow" }
    else { return "DarkYellow" }
}

function WrapTheLines {
    [CmdletBinding()]
    param (
        [Parameter(Position=0,Mandatory=$True)] 
        [Object] $text,
        [Parameter(Position=1,Mandatory=$True)]
        [int16] $width = 5
    )
    BEGIN {
        $retval = ""
    }
    PROCESS {
        foreach ($line in $text) {
            $line = $line.ToString()
            #If the line is actualy many lines
            #We break it up by the newline char and call ourself recursively
            if ($line.Contains("`n")) { 
                $retval += WrapTheLines $line.split("`n") $width
            }
            #If the line is longer than the width
            #We break it up by the width
            elseif ($line.length -gt $width) { 
                #Get the line size
                $len = $line.length
                #Set the max size to width
                $Split = $width
                #set our new, empty var
                $new = ""
                #get the number of widths the text line is
                $repeat = [Math]::Floor($len / $Split)
                #and process each section accordingly
                for ($i = 0; $i -lt $repeat; $i++){
                    $retval += $line.Substring($i * $Split, $Split).TrimEnd(" ")
                    $retval += "`n"
                }
                #if there's any text left over...
                if ($remainder = $len % $Split) {
                   $retval += $line.Substring($len - $remainder).TrimEnd(" ")
                   $retval += "`n"
                }

            }
            #Or we have a normal line of text
            else {
                $retval += $line.TrimEnd(" ")
                $retval += "`n"
            }
        }
    }
    END {
        $retval.TrimEnd("`n") #.Split("`n")
    }
}