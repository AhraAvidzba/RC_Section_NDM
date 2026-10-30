# -------------------- ДЛЯ ТЕСТОВ --------------------
# Читает точную идентичность процесса. Не закрывает процессы; недоступные сведения оставляет Unknown.
function Test-ObserveOwnIdentityFinal01 {
    param([int]$IdentityID,[string]$ExpectedTicks,[string]$ExpectedName='')
    if($IdentityID-le0-or$ExpectedTicks-notmatch'^\d{18}$'){throw 'Exact own PID and UTC start ticks required'}
    $state=[ordered]@{PID=$IdentityID;StartTicksUTC=$ExpectedTicks;Name=$ExpectedName;State='Unknown';ExactIdentityTerminal=$false;ObservedName=$null;ObservedStartTicksUTC=$null;Proof=$null;Error=$null;ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');ProcessModifiedOrClosed=$false}
    $process=$null
    try{$process=Get-Process -Id $IdentityID -ErrorAction Stop}catch{
        if($_.FullyQualifiedErrorId-like'NoProcessFoundForGivenId*'){$state.State='Absent';$state.Proof='Get-Process confirmed no process for this PID'}else{$state.Error=$_.Exception.Message}
    }
    if($process){
        try{
            $observedName=[string]$process.ProcessName
            if([string]::IsNullOrWhiteSpace($observedName)-or$null-eq$process.StartTime){throw 'Readable executable name and UTC start required'}
            $ticks=[string]$process.StartTime.ToUniversalTime().Ticks
            $state.ObservedName=$observedName;$state.ObservedStartTicksUTC=$ticks
            if($ExpectedName-and$observedName-ine$ExpectedName){throw 'Different executable requires independent CIM confirmation'}
            if($ticks-ceq$ExpectedTicks-and(-not$ExpectedName-or$observedName-ieq$ExpectedName)){$state.State='Live'}else{$state.State='ReusedByDifferentProcess'}
            $state.Proof='Get-Process executable name and exact UTC start ticks'
        }catch{
            try{
                $cim=Get-CimInstance Win32_Process -Filter ('ProcessId='+$IdentityID) -ErrorAction Stop
                if($null-eq$cim-or$null-eq$cim.CreationDate-or[string]::IsNullOrWhiteSpace([string]$cim.Name)){throw 'Independent CIM identity unavailable'}
                $cimName=[IO.Path]::GetFileNameWithoutExtension([string]$cim.Name)
                $cimTicks=[string]$cim.CreationDate.ToUniversalTime().Ticks
                # CIM может округлять время старта: близкое время того же exe не доказывает исчезновение собственного процесса.
                if(($ExpectedName-and$cimName-ine$ExpectedName)-or[math]::Abs([decimal]$cimTicks-[decimal]$ExpectedTicks)-ge10000){
                    $state.State='ReusedByDifferentProcess';$state.ObservedName=$cimName;$state.ObservedStartTicksUTC=$cimTicks
                    $state.Proof='Independent CIM different executable or creation time separated by at least1ms'
                }else{throw 'Exact own identity cannot be distinguished with CIM precision'}
            }catch{$state.State='Unknown';$state.Error=$_.Exception.Message}
        }
    }
    $state.ExactIdentityTerminal=($state.State-in@('Absent','ReusedByDifferentProcess'))
    [pscustomobject]$state
}
