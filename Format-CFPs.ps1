# Silly script to import CLD proposals and format in markdown
$Proposals = import-csv '/Users/wframe/Downloads/Community Lightning Demo Proposals.csv'
$PropArray = $Proposals | Foreach-Object {
    $Name = $_.'Your Name'
    $Topic = $_.'Topic Title'
    $Abstract = $_.'Topic Abstract'
    $Blog = $_.'Blog Link'
    $GitHub = $_.'GitHub Handle'

    $Twitter = $_.'Twitter Handle'
    if($Twitter) {$Name = "[$Name]($Twitter)"}

    $Append = ""
    if($GitHub) {$Append += "[GitHub]($GitHub)"}
    if($Blog) {$Append += "[Blog]($Blog)"}
    if($GitHub -or $Blog) {$Name = "$Name ($Append)"}

$Text = @"
### $Topic

Proposed by $Name`:

$Abstract
"@
    [pscustomobject]@{
        Name = $Name
        Topic = $Topic
        Text = $Text
    }
} | Sort-Object Name, Topic

@( $PropArray | Select -ExpandProperty Text ) -join "`n`n" | pbcopy