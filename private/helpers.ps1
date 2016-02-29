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