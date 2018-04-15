Function ConvertTo-XlAddress {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$True, ValueFromPipeline=$True)]
        [int]$RowIndex,
        [Parameter(Mandatory=$False)]
        [string]$Delimiter
    )
    begin {
        $ErrorActionPreference = "Stop"
        $Base = 26
        $ASCII = 65
        $store = @()
    }
    process {
        if ($RowIndex -lt 1) { throw "Argument Error: input is grater than 0." }
        $Ret = ""
        do {
            $RowIndex = $RowIndex - 1
            $Ret = [char][int]($RowIndex % $Base + $ASCII) + $Ret
            $RowIndex = [Math]::Floor($RowIndex / $Base)
        } while($RowIndex -gt 0)
        $store += $Ret
    }
    end {
        if ($Delimiter) {
            $store -join $Delimiter
        } else {
            $store
        }
    }
}
