<#
통합고객목록에서 날짜별로 CSV 를 내려받는다.

등록일을 하루씩 밀며 [검색 → 목록 탭 → 전체선택 → 처리 탭 → Excel → 저장] 을 반복한다.

컨트롤을 GetDlgCtrlID 번호로 찾지 않는다. 이 프로그램의 컨트롤 번호는 창을 새로 열 때마다
달라져서, 한 번 뽑아 둔 번호가 다음 실행에서는 맞지 않는다. 대신 창에 적힌 글자와 클래스,
그리고 서로의 위치 관계로 찾는다. 이것들은 창을 다시 열어도 그대로다.

사용:
    powershell -ExecutionPolicy Bypass -File crm4_collect.ps1 -RawRoot C:\CRM자동화\raw
    powershell -ExecutionPolicy Bypass -File crm4_collect.ps1 -RawRoot ... -Start 2026-07-01 -End 2026-08-21
    powershell -ExecutionPolicy Bypass -File crm4_collect.ps1 -RawRoot ... -Diagnose    클릭 없이 인식만 확인
#>

param(
    # 내려받은 CSV 가 쌓이는 곳. 없는 폴더면 만들지 않고 멈춘다.
    [Parameter(Mandatory = $true)]
    [string]$RawRoot,

    # 수집할 등록일 범위. 주지 않으면 어제 하루만 받는다.
    #
    # 지난 날짜를 지금 뽑으면 그날 뽑았을 때가 아니라 오늘 기준으로 갱신된 상태가 나온다.
    # 그 사이 유입경로가 바뀌거나 지워진 건이 섞이므로 다른 날과 기준이 어긋난다.
    # 그래서 빠진 날은 저절로 메우지 않는다. 굳이 받아야 하면 -Start 와 -Backfill 을 함께 준다.
    [datetime]$Start,
    [datetime]$End = (Get-Date).Date.AddDays(-1),

    # 지난 날짜를 받겠다고 분명히 밝히는 스위치. -Start 를 과거로 줄 때 필요하다.
    [switch]$Backfill,

    # 클릭을 전혀 하지 않고, 컨트롤을 제대로 찾는지만 확인한다.
    [switch]$Diagnose,

    # 이미 받아둔 날짜도 다시 받는다.
    [switch]$Force,

    # 등록일 체크박스를 누르지 않는다. 이미 켜 둔 상태로 시작할 때 쓴다.
    [switch]$SkipDateCheck,

    # 목록 탭의 전체선택을 누르지 않는다.
    [switch]$SkipSelectAll,

    # 로그를 남길 파일. 주지 않으면 스크립트 옆 logs 폴더에 만든다.
    [string]$LogFile,

    # 등록일 체크박스 상태를 기억해 두는 파일. 창을 계속 열어 두고 반복 실행할 때 쓴다.
    [string]$StateFile,

    # 단계 사이 기본 대기(초). 서버가 느리면 늘린다.
    [int]$Wait = 2
)

# 탭 순서는 창 구성이라 바뀌지 않는다.
$TabIndex = @{ Filter = 0; List = 4; Process = 5 }

Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class C4 {
    public delegate bool EnumProc(IntPtr h, IntPtr p);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr p);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, IntPtr e);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    public static string ClassOf(IntPtr h){ StringBuilder s=new StringBuilder(256); GetClassName(h,s,256); return s.ToString(); }
    public static string TextOf(IntPtr h){ StringBuilder s=new StringBuilder(1024); GetWindowText(h,s,1024); return s.ToString(); }
}
"@

$script:Log = @()
function Write-Log([string]$msg, [string]$color = "Gray") {
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg
    Write-Host $line -ForegroundColor $color
    $script:Log += $line
}

# 델리게이트 콜백 안에서는 함수의 지역 변수가 보이지 않을 수 있다.
# 콜백은 담기만 하고, 판단은 바깥에서 한다.

function Get-TopWindows {
    $script:tops = @()
    $cb = [C4+EnumProc]{
        param($h, $l)
        $o = 0
        [void][C4]::GetWindowThreadProcessId($h, [ref]$o)
        $script:tops += [pscustomobject]@{
            Handle = $h; OwnerPid = [int]$o
            Title = [C4]::TextOf($h); Class = [C4]::ClassOf($h)
            Visible = [C4]::IsWindowVisible($h)
        }
        return $true
    }
    [void][C4]::EnumWindows($cb, [IntPtr]::Zero)
    return $script:tops
}

function Find-MainWindow([int]$processId) {
    $m = @(Get-TopWindows | Where-Object {
        $_.OwnerPid -eq $processId -and $_.Visible -and $_.Title -like "*통합고객목록*"
    })
    if ($m.Count -gt 0) { return $m[0].Handle }
    return [IntPtr]::Zero
}

function Get-Descendants([IntPtr]$root) {
    $script:kids = @()
    $cb = [C4+EnumProc]{
        param($h, $l)
        $r = New-Object C4+RECT
        [void][C4]::GetWindowRect($h, [ref]$r)
        $script:kids += [pscustomobject]@{
            H = $h
            P = [C4]::GetParent($h)
            Class = [C4]::ClassOf($h)
            Text  = [C4]::TextOf($h)
            L = $r.Left; T = $r.Top; R = $r.Right; B = $r.Bottom
            W = $r.Right - $r.Left; Ht = $r.Bottom - $r.Top
            Visible = [C4]::IsWindowVisible($h)
        }
        return $true
    }
    [void][C4]::EnumChildWindows($root, $cb, [IntPtr]::Zero)
    return $script:kids
}

<#
글자와 위치로 컨트롤을 집는다.

  탭          클래스가 SysTabControl32 인 것
  각 탭 페이지 글자가 '고객조건' / '목록' / '처리' 인 컨테이너
  등록일 체크  글자가 '등록일' 인 BUTTON
  날짜 두 칸   SysDateTimePick32 중 등록일 체크와 같은 줄에 있는 것, 왼쪽부터 시작/종료
  검색 버튼    탭 위쪽에 떠 있는 작은 버튼 중 가장 오른쪽
  건수 라벨    '조회고객' 이 적힌 STATIC
  전체선택     글자가 '전체선택' 인 BUTTON
  Excel       '기능' 상자 안의 버튼 두 개 중 오른쪽 (왼쪽은 SMS)
#>
# 탭 페이지들은 화면상 같은 자리에 겹쳐 있다. 위치만으로 고르면 다른 탭의 컨트롤이 걸린다.
# 부모를 거슬러 올라가 정말 그 안에 든 것인지 확인한다.
function Test-Under($ctrl, $ancestor) {
    $p = $ctrl.P
    $guard = 0
    while ($p -ne [IntPtr]::Zero -and $guard -lt 20) {
        if ($p -eq $ancestor.H) { return $true }
        $p = [C4]::GetParent($p)
        $guard++
    }
    return $false
}

function Resolve-Controls([IntPtr]$win) {
    $all = Get-Descendants $win
    $c = @{}

    $tab = $all | Where-Object { $_.Class -match "SysTabControl32" } | Select-Object -First 1
    $c.Tab = $tab

    foreach ($p in @(@("PageFilter","고객조건"), @("PageList","목록"), @("PageProcess","처리"))) {
        $c[$p[0]] = $all | Where-Object { $_.Text -eq $p[1] -and $_.W -gt 400 } | Select-Object -First 1
    }

    $chk = $all | Where-Object { $_.Class -match "BUTTON" -and $_.Text -eq "등록일" } | Select-Object -First 1
    $c.DateCheck = $chk

    if ($chk) {
        $picks = @($all | Where-Object {
            $_.Class -match "SysDateTimePick32" -and $_.Visible -and $_.W -gt 20 -and
            [Math]::Abs($_.T - $chk.T) -le 8
        } | Sort-Object L)

        # 같은 자리에 날짜 칸이 겹쳐 있다. 그대로 두면 앞의 둘이 모두 '시작일' 이 되어
        # 종료일이 영영 입력되지 않는다. 가로 위치가 겹치는 것은 하나로 본다.
        $spread = @()
        foreach ($p in $picks) {
            $dup = $false
            foreach ($q in $spread) { if ([Math]::Abs($q.L - $p.L) -le 4) { $dup = $true; break } }
            if (-not $dup) { $spread += $p }
        }
        $c.DatePicks = $spread
        if ($spread.Count -ge 2) { $c.DateFrom = $spread[0]; $c.DateTo = $spread[1] }
    }

    $c.SelectAll = $all | Where-Object { $_.Class -match "BUTTON" -and $_.Text -eq "전체선택" } | Select-Object -First 1

    # 건수 라벨은 검색 전에는 글자가 비어 있을 수 있다. 전체선택과 같은 줄의 오른쪽이라는 위치로 잡는다.
    if ($c.SelectAll) {
        $c.CountLabel = $all | Where-Object {
            $_.Class -match "STATIC" -and $_.L -gt $c.SelectAll.R -and
            [Math]::Abs($_.T - $c.SelectAll.T) -le 12
        } | Sort-Object L | Select-Object -First 1
    }

    if ($tab) {
        $c.Search = $all | Where-Object {
            $_.Visible -and $_.Text -eq "" -and $_.B -le $tab.T -and
            $_.W -ge 40 -and $_.W -le 140 -and $_.Ht -ge 14 -and $_.Ht -le 40
        } | Sort-Object R -Descending | Select-Object -First 1
    }

    # 처리 탭의 '기능' 상자 안에 SMS 와 Excel 이 나란히 있다. 오른쪽이 Excel.
    $grp = $all | Where-Object {
        $_.Text -eq "기능" -and ($null -eq $c.PageProcess -or (Test-Under $_ $c.PageProcess))
    } | Select-Object -First 1

    if ($grp) {
        $btns = @($all | Where-Object {
            $_.H -ne $grp.H -and $_.W -ge 30 -and $_.Ht -ge 14 -and (Test-Under $_ $grp)
        } | Sort-Object L)
        if ($btns.Count -ge 2) { $c.Excel = $btns[1] } elseif ($btns.Count -eq 1) { $c.Excel = $btns[0] }
    }

    $c.Total = $all.Count
    return $c
}

function Need($ctrls, [string]$name) {
    if (-not $ctrls[$name]) { throw "$name 을(를) 찾지 못했습니다." }
    return $ctrls[$name]
}

# 좌표로 누르는 방식이라 대상 창이 맨 앞이 아니면 엉뚱한 창을 클릭하게 된다.
# SetForegroundWindow 는 조용히 실패할 수 있으므로 매번 확인하고, 아니면 누르지 않고 멈춘다.
function Test-ForegroundIsCrm {
    $fg = [C4]::GetForegroundWindow()
    if ($fg -eq [IntPtr]::Zero) { return $false }
    $o = 0
    [void][C4]::GetWindowThreadProcessId($fg, [ref]$o)
    return ([int]$o -eq $script:TargetPid)      # 달력이나 안내창도 CRM4 것이므로 통과시킨다
}

# 창 전환 직후나 달력이 닫히는 순간처럼 포그라운드가 잠깐 비는 때가 있다.
# 몇 번 되찾아 보고, 그래도 다른 프로그램이 앞이면 그때 멈춘다.
function Confirm-Foreground {
    for ($i = 0; $i -lt 6; $i++) {
        if (Test-ForegroundIsCrm) { return $true }
        [void][C4]::SetForegroundWindow($script:TargetWindow)
        Start-Sleep -Milliseconds 400
    }
    return (Test-ForegroundIsCrm)
}

function Click-Ctrl($ctrl, [int]$offsetX = 0) {
    if (-not (Confirm-Foreground)) {
        throw "CRM4 창이 맨 앞이 아닙니다. 다른 프로그램을 클릭할 위험이 있어 중단합니다."
    }
    $r = New-Object C4+RECT
    [void][C4]::GetWindowRect($ctrl.H, [ref]$r)      # 창을 옮겼을 수 있으니 지금 좌표를 다시 읽는다
    $x = if ($offsetX -gt 0) { $r.Left + $offsetX } else { [int](($r.Left + $r.Right) / 2) }
    $y = [int](($r.Top + $r.Bottom) / 2)
    [void][C4]::SetCursorPos($x, $y)
    Start-Sleep -Milliseconds 150
    [C4]::mouse_event(0x02, 0, 0, 0, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 60
    [C4]::mouse_event(0x04, 0, 0, 0, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 400
}

# 한글 입력기가 켜져 있으면 타이핑한 영문이 그대로 한글로 바뀐다.
#   test\customers_2026-08-22.csv  ->  ㅅㄷㄴㅅ\쳔새ㅡㄷㄱㄴ_2026-08-22.ㅊㄴㅍ
# 폴더 이름까지 바뀌어 엉뚱한 곳에 저장되므로, 글자는 치지 않고 클립보드로 넣는다.
# 붙여넣기는 입력기를 거치지 않아 어떤 상태에서도 같은 결과가 나온다.
function Send-Text([string]$text, [IntPtr]$expect) {
    $prev = $null
    try { $prev = Get-Clipboard -Raw -ErrorAction SilentlyContinue } catch {}
    try {
        Set-Clipboard -Value $text
        Start-Sleep -Milliseconds 150
        Send-Keys "^a" $expect          # 미리 채워진 이름을 지운다
        Start-Sleep -Milliseconds 120
        Send-Keys "^v" $expect
        Start-Sleep -Milliseconds 250
    } catch {
        # 클립보드를 못 쓰면 예전처럼 친다. 입력기가 영문이면 그래도 된다.
        Write-Log "  클립보드를 쓰지 못해 직접 입력합니다: $($_.Exception.Message)" "Yellow"
        Send-Keys ($text.Replace("+","{+}").Replace("^","{^}").Replace("%","{%}").Replace("~","{~}")) $expect
    } finally {
        if ($null -ne $prev) { try { Set-Clipboard -Value $prev } catch {} }
    }
}

function Send-Keys([string]$keys, [IntPtr]$expect) {
    if (-not (Confirm-Foreground)) {
        throw "CRM4 창이 맨 앞이 아닙니다. 엉뚱한 곳에 입력될 위험이 있어 중단합니다."
    }
    [System.Windows.Forms.SendKeys]::SendWait($keys)
}

# 탭을 바꾸면 그 페이지에서만 만들어지는 컨트롤이 있으므로 매번 다시 찾는다.
function Switch-Tab([IntPtr]$win, [string]$name) {
    $c = Resolve-Controls $win
    $tab = Need $c "Tab"
    [void][C4]::SendMessage($tab.H, 0x1330, [IntPtr]$TabIndex[$name], [IntPtr]::Zero)   # TCM_SETCURFOCUS
    Start-Sleep -Milliseconds 700

    $c = Resolve-Controls $win
    $page = Need $c "Page$name"
    if (-not [C4]::IsWindowVisible($page.H)) { throw "$name 탭으로 전환하지 못했습니다." }
    return $c
}

# 메시지로 넣어 보고, 되읽어서 확인하고, 안 들어갔으면 타이핑으로 다시 시도한다.
# 넣었다고 믿고 넘어가면 엉뚱한 날짜로 검색해도 알 수가 없다.
# 날짜는 사람이 하듯 클릭하고 타이핑해서만 넣는다.
# 값을 되읽는 메시지는 쓰지 않는다. 그 메시지가 CRM4 를 죽게 만들었다.
# 제대로 들어갔는지는 내려받은 파일의 등록일로 확인한다.
function Set-DatePicker($ctrl, [datetime]$value, [string]$label, [int]$mode = 1) {
    Click-Ctrl $ctrl 12                       # 왼쪽 끝 = 연도 칸
    Send-Keys "{ESC}" $script:TargetWindow    # 달력이 펼쳐졌으면 닫는다
    Start-Sleep -Milliseconds 250

    if ($mode -eq 1) {
        # 칸이 다 차면 다음 칸으로 저절로 넘어가는 것이 표준 동작이다.
        Send-Keys $value.ToString("yyyyMMdd") $script:TargetWindow
    } else {
        # 자동으로 안 넘어가는 경우를 위해 칸 이동을 직접 지정한다.
        Send-Keys ($value.ToString("yyyy") + "{RIGHT}" + $value.ToString("MM") + "{RIGHT}" + $value.ToString("dd")) $script:TargetWindow
    }
    Start-Sleep -Milliseconds 600
    Write-Log "  $label 입력: $($value.ToString('yyyy-MM-dd')) (방식 $mode)"
}

# 이 프로그램의 체크박스는 BM_GETCHECK 에 늘 0 을 돌려주어 켜졌는지 알 수 없다.
# 그래서 상태를 읽는 대신 결과로 판정한다. 받아온 파일의 등록일이 요청한 날짜면 조건이 걸린 것이다.
function Test-FileDate([string]$path, [datetime]$day) {
    $lines = @(Get-Content $path -Encoding Default -ErrorAction SilentlyContinue)
    if ($lines.Count -lt 2) { return $true }
    $body = @($lines | Select-Object -Skip 1 | Where-Object { $_.Trim() })
    if ($body.Count -eq 0) { return $true }

    # '2026-08-20 오전 9:34:27' 형태는 등록일 칸에만 나온다. 다른 날짜 칸에는 오전/오후가 없다.
    $pat = "*" + $day.ToString("yyyy-MM-dd") + " 오*"
    $hit = @($body | Where-Object { $_ -like $pat }).Count
    return (($hit / $body.Count) -ge 0.8)
}

# 프로그램은 내보내기가 끝났는지 알려주지 않는다. 예전에는 정해진 시간만 기다렸다가
# 파일이 있기만 하면 넘어갔고, 내보내기가 느린 날에는 앞부분만 써진 파일을 그대로 집계했다.
# 그래서 파일이 생길 때까지 기다린 뒤, 크기가 더 늘지 않을 때까지 한 번 더 기다린다.
function Wait-ExportFile([string]$dir, [string]$dest, [datetime]$since, [int]$timeoutSec = 120) {
    $deadline = (Get-Date).AddSeconds($timeoutSec)

    $path = $null
    while ((Get-Date) -lt $deadline -and -not $path) {
        if (Test-Path $dest) {
            $path = $dest
        } else {
            # 프로그램이 이름을 제 방식대로 붙일 수 있다. 방금 생긴 CSV 를 찾아 맞춘다.
            $fresh = Get-ChildItem $dir -File -Filter *.csv -ErrorAction SilentlyContinue |
                     Where-Object { $_.LastWriteTime -gt $since } |
                     Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($fresh) { $path = $fresh.FullName }
        }
        if (-not $path) { Start-Sleep -Milliseconds 500 }
    }
    if (-not $path) { return $null }

    # 크기가 세 번 연속 그대로면 다 써진 것으로 본다.
    $prev = -1; $stable = 0
    while ((Get-Date) -lt $deadline) {
        $item = Get-Item $path -ErrorAction SilentlyContinue
        if ($item) {
            if ($item.Length -eq $prev -and $item.Length -gt 0) { $stable++ } else { $stable = 0 }
            $prev = $item.Length
            if ($stable -ge 3) { break }
        }
        Start-Sleep -Milliseconds 500
    }
    return $path
}

# 받아온 파일에 몇 건이 담겼는지 센다. 상담메모 칸에 줄바꿈이 들어 있어서 줄 수로는 못 센다.
# Import-Csv 는 따옴표 안의 줄바꿈을 제대로 묶어 준다.
function Measure-CsvRows([string]$path) {
    try {
        $n = @(Import-Csv $path -Encoding Default -ErrorAction Stop).Count
        if ($n -gt 0) { return $n }
    } catch { }
    return [math]::Max(0, @(Get-Content $path -Encoding Default | Where-Object { $_.Trim() }).Count - 1)
}

# 검색은 서버 왕복이라 걸리는 시간이 일정하지 않다. 고정 대기 대신 건수 라벨이 멈출 때까지 본다.
function Wait-Search([IntPtr]$win, [int]$timeoutSec = 180) {
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    $stable = 0; $prev = $null

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 1
        $c = Resolve-Controls $win
        $now = if ($c.CountLabel) { [C4]::TextOf($c.CountLabel.H) } else { "" }
        if ($now -eq $prev -and $now -match "조회고객") { $stable++ } else { $stable = 0 }
        $prev = $now
        if ($stable -ge 3) { break }
    }
    if ($prev -match "조회고객\s*:\s*([\d,]+)") { return [int]($Matches[1] -replace ",", "") }
    return -1
}

# ─────────── 실행 ───────────

$proc = Get-Process CRM4enterprise -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $proc) { Write-Host "CRM4enterprise 가 실행 중이 아닙니다." -ForegroundColor Red; exit 1 }

$win = Find-MainWindow $proc.Id
if ($win -eq [IntPtr]::Zero) {
    Write-Host "통합고객목록 창을 열어주세요. (고객관리 > 통합고객목록)" -ForegroundColor Red; exit 1
}
$script:TargetWindow = $win
$script:TargetPid = $proc.Id

if (-not (Test-Path $RawRoot)) {
    Write-Host "원본 저장 폴더가 없습니다: $RawRoot" -ForegroundColor Red
    Write-Host "폴더를 직접 만드시거나, -RawRoot 로 쓸 경로를 지정하세요." -ForegroundColor Yellow
    exit 1
}

# 저장 대화상자는 상대 경로를 자기가 마지막에 쓰던 폴더 기준으로 푼다. 우리 작업 폴더가 아니다.
# '.\test' 를 그대로 넘기면 엉뚱한 데 저장되고, 우리는 여기서 찾다가 없다고 한다.
$RawRoot = (Resolve-Path -LiteralPath $RawRoot).ProviderPath

# 탭을 한 번도 열지 않으면 그 페이지의 컨트롤은 아직 만들어지지 않는다.
# 점검할 때는 탭을 차례로 넘겨보며 모은다. 탭 전환은 메시지로 하므로 마우스는 움직이지 않는다.
$ctrls = Resolve-Controls $win
$tab0 = $ctrls.Tab
if (-not $tab0) { Write-Host "탭 컨트롤을 찾지 못했습니다. 통합고객목록 창이 맞는지 확인하세요." -ForegroundColor Red; exit 1 }

foreach ($name in @("List","Process","Filter")) {
    [void][C4]::SendMessage($tab0.H, 0x1330, [IntPtr]$TabIndex[$name], [IntPtr]::Zero)
    Start-Sleep -Milliseconds 600
    $found = Resolve-Controls $win
    foreach ($k in @($found.Keys)) { if ($k -ne "Total" -and $found[$k] -and -not $ctrls[$k]) { $ctrls[$k] = $found[$k] } }
}

Write-Host "찾은 결과:" -ForegroundColor Gray
foreach ($k in @("Tab","PageFilter","PageList","PageProcess","DateCheck","DateFrom","DateTo","Search","CountLabel","SelectAll","Excel")) {
    if ($ctrls[$k]) {
        Write-Host ("  OK   {0,-12} ({1},{2}) {3}" -f $k, $ctrls[$k].L, $ctrls[$k].T, $ctrls[$k].Class.Split('.')[1]) -ForegroundColor Green
    } else {
        Write-Host ("  실패 {0,-12}" -f $k) -ForegroundColor Red
    }
}

if ($Diagnose) {
    $dp = @($ctrls.DatePicks)
    $where = ($dp | ForEach-Object { "($($_.L),$($_.T)) w$($_.W)" }) -join "  "
    Write-Host ""
    Write-Host ("날짜 칸 {0}개: {1}" -f $dp.Count, $where) -ForegroundColor Gray
    Write-Host ""
    Write-Host "점검 모드였습니다. 클릭하지 않았습니다." -ForegroundColor Cyan
    exit 0
}

foreach ($k in @("Tab","PageFilter","PageList","PageProcess","DateCheck","DateFrom","DateTo","Search","CountLabel","Excel")) {
    if (-not $ctrls[$k]) { Write-Host "`n$k 을(를) 찾지 못해 중단합니다." -ForegroundColor Red; exit 1 }
}

# 기본은 어제 하루뿐이다. 밀린 날을 알아서 메우지 않는다.
$yesterday = (Get-Date).Date.AddDays(-1)
if (-not $PSBoundParameters.ContainsKey('Start')) { $Start = $End }

if (($Start -lt $yesterday) -and -not $Backfill) {
    Write-Host ""
    Write-Host "$($Start.ToString('yyyy-MM-dd')) 은(는) 어제보다 이전입니다." -ForegroundColor Red
    Write-Host "지금 뽑으면 그날 뽑았을 때가 아니라 오늘 기준으로 갱신된 값이 나옵니다." -ForegroundColor Yellow
    Write-Host "다른 날과 기준이 달라지므로 받지 않습니다." -ForegroundColor Yellow
    Write-Host "그래도 받아야 하면 -Backfill 을 함께 주세요." -ForegroundColor Gray
    exit 2
}
if ($Backfill -and $Start -lt $yesterday) {
    Write-Log "경고: 지난 날짜를 받습니다. 이 값은 다음날 받은 다른 날과 기준이 다릅니다." "Yellow"
}
if ($Start -gt $End) { Write-Host "받을 날짜가 없습니다." -ForegroundColor Green; exit 0 }

Write-Log "수집 시작: $($Start.ToString('yyyy-MM-dd')) ~ $($End.ToString('yyyy-MM-dd'))" "Cyan"
Write-Log "저장 위치: $RawRoot" "Cyan"

[void][C4]::SetForegroundWindow($win)
Start-Sleep -Seconds 2
if ([C4]::GetForegroundWindow() -ne $win) {
    Write-Host "통합고객목록 창을 맨 앞으로 가져오지 못했습니다." -ForegroundColor Red
    Write-Host "창을 직접 한 번 클릭해 활성화한 뒤 다시 실행하세요." -ForegroundColor Yellow
    exit 1
}

Write-Host "5 초 뒤 시작합니다. 마우스와 키보드를 건드리지 마세요. (Ctrl+C 로 취소)" -ForegroundColor Yellow
Start-Sleep -Seconds 5

$ctrls = Switch-Tab $win "Filter"
# 체크박스 상태는 읽을 수 없다. 대신 지난 실행이 어떻게 끝났는지를 적어 두고 그것을 믿는다.
# 창을 계속 열어 둔 채 다시 실행하면 이미 켜져 있으므로, 누르면 오히려 꺼진다.
if (-not $StateFile) {
    $StateFile = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "state.txt"
}
$wasOn = (Test-Path $StateFile) -and ((Get-Content $StateFile -ErrorAction SilentlyContinue) -eq "datecheck=on")

if ($SkipDateCheck) {
    Write-Log "등록일 체크박스는 건드리지 않습니다."
} elseif ($wasOn) {
    Write-Log "등록일 조건이 지난 실행에서 켜진 채 끝났으므로 그대로 둡니다."
} else {
    Click-Ctrl (Need $ctrls "DateCheck")
    Write-Log "등록일 조건을 켰습니다." "Yellow"
}

$ok = 0; $empty = 0; $failed = 0
$script:DateMode = 2      # 이 프로그램은 칸이 차도 다음 칸으로 넘어가지 않는다. 칸 이동을 직접 준다.
$script:Flipped = $false
$script:Truncated = $false   # 내보내기가 덜 끝나서 실패했는지

$day = $Start.Date

while ($day -le $End.Date) {
    $stamp = $day.ToString("yyyy-MM-dd")
    $dest = Join-Path $RawRoot "customers_$stamp.csv"

    # 일요일은 영업하지 않아 받을 것이 없다. 괜히 CRM4 를 한 번 더 돌리지 않는다.
    if ($day.DayOfWeek -eq [System.DayOfWeek]::Sunday) {
        Write-Log "$stamp 일요일, 건너뜀"
        $day = $day.AddDays(1); continue
    }

    if ((Test-Path $dest) -and -not $Force) {
        Write-Log "$stamp 이미 있음, 건너뜀"
        $day = $day.AddDays(1); continue
    }
    if (Test-Path $dest) { Remove-Item $dest -Force }

    # 한 날짜를 최대 세 번까지 시도한다. 실패할 때마다 조건을 하나씩 바꿔 본다.
    $done = $false
    for ($try = 1; $try -le 3 -and -not $done; $try++) {
        try {
            if (-not (Get-Process -Id $proc.Id -ErrorAction SilentlyContinue)) {
                throw "CRM4enterprise 가 종료되었습니다."
            }
            [void][C4]::SetForegroundWindow($win)
            Start-Sleep -Milliseconds 400

            $c = Switch-Tab $win "Filter"
            Set-DatePicker (Need $c "DateFrom") $day "시작일" $script:DateMode
            Set-DatePicker (Need $c "DateTo") $day "종료일" $script:DateMode

            Click-Ctrl (Need $c "Search")
            $count = Wait-Search $win

            if ($count -eq 0) {
                Write-Log "$stamp 0 건, 저장 생략"
                $empty++; $done = $true; break
            }
            Write-Log "$stamp 조회 $count 건" "White"

            if (-not $SkipSelectAll) {
                $c = Switch-Tab $win "List"
                Click-Ctrl (Need $c "SelectAll")
            }

            $c = Switch-Tab $win "Process"
            $started = Get-Date
            Click-Ctrl (Need $c "Excel")
            Start-Sleep -Seconds $Wait

            # 저장 대화상자를 기다린다.
            $dlg = [IntPtr]::Zero
            $deadline = (Get-Date).AddSeconds(30)
            while ((Get-Date) -lt $deadline -and $dlg -eq [IntPtr]::Zero) {
                $cand = @(Get-TopWindows | Where-Object {
                    $_.OwnerPid -eq $proc.Id -and $_.Visible -and $_.Class -eq "#32770"
                })
                if ($cand.Count -gt 0) { $dlg = $cand[0].Handle } else { Start-Sleep -Milliseconds 500 }
            }
            if ($dlg -eq [IntPtr]::Zero) { throw "저장 대화상자가 뜨지 않았습니다." }

            [void][C4]::SetForegroundWindow($dlg)
            Start-Sleep -Milliseconds 500
            Send-Text $dest $dlg
            Start-Sleep -Milliseconds 300
            Send-Keys "{ENTER}" $dlg
            Start-Sleep -Seconds ($Wait * 2)

            # 저장이 끝나면 '엑셀 다운로드 성공' 안내창이 뜬다. 닫아야 다음으로 넘어간다.
            foreach ($n in @(Get-TopWindows | Where-Object {
                $_.OwnerPid -eq $proc.Id -and $_.Visible -and $_.Handle -ne $win -and $_.Title -notlike "*통합고객목록*"
            })) {
                [void][C4]::SetForegroundWindow($n.Handle)
                Start-Sleep -Milliseconds 300
                Send-Keys "{ENTER}" $n.Handle
                Start-Sleep -Milliseconds 400
            }

            $src = Wait-ExportFile $RawRoot $dest $started
            if (-not $src) { throw "파일이 만들어지지 않았습니다." }
            if ($src -ne $dest) { Move-Item $src $dest -Force }

            # 프로그램이 말한 건수와 파일에 실제로 담긴 건수를 맞춰 본다. 이 대조가 없던 동안
            # 2026-09-15 은 조회 229 건에 파일 69 행이 들어왔는데도 조용히 넘어갔다.
            $rows = Measure-CsvRows $dest
            if ($count -gt 0 -and $rows -lt [math]::Floor($count * 0.9)) {
                Remove-Item $dest -Force -ErrorAction SilentlyContinue
                $script:Truncated = $true
                throw "받아온 파일이 $rows 건뿐입니다. 조회는 $count 건이었습니다."
            }

            if (Test-FileDate $dest $day) {
                Write-Log "$stamp 저장 완료 (조회 $count 건 · 파일 $rows 건)" "Green"
                $ok++; $done = $true
                "datecheck=on" | Out-File $StateFile -Encoding ascii
            } else {
                Remove-Item $dest -Force
                throw "받아온 자료의 등록일이 $stamp 이 아닙니다."
            }
        }
        catch {
            Write-Log "$stamp 시도 $try 실패: $($_.Exception.Message)" "Yellow"
            if (Test-Path $dest) { Remove-Item $dest -Force -ErrorAction SilentlyContinue }

            # 파일이 잘려서 실패한 것이면 조회 조건은 멀쩡하다. 체크박스를 건드리면 오히려
            # 날짜 조건이 풀려서 엉뚱한 자료를 받는다. 기다리는 시간만 늘려서 다시 한다.
            if ($script:Truncated) {
                $script:Truncated = $false
                $Wait = $Wait * 2
                Write-Log "  내보내기가 느립니다. 대기를 $Wait 초로 늘려 다시 해봅니다." "Yellow"
            }
            # 날짜 입력 방식은 바꾸지 않는다. 이 프로그램에서 칸 이동을 주는 방식만 제대로 들어간다.
            elseif ($try -eq 1) {
                Write-Log "  등록일 체크박스를 뒤집고 다시 해봅니다." "Yellow"
                try {
                    $c = Switch-Tab $win "Filter"
                    Click-Ctrl (Need $c "DateCheck")
                    $script:Flipped = $true
                } catch { Write-Log "  체크박스를 누르지 못했습니다: $($_.Exception.Message)" "Red" }
            }
            elseif ($try -eq 2) {
                # 뒤집어도 안 됐으면 원래대로 돌려놓고, 일시적인 문제였을 수 있으니 한 번 더 한다.
                if ($script:Flipped) {
                    Write-Log "  체크박스를 원래대로 돌리고 다시 해봅니다." "Yellow"
                    try {
                        $c = Switch-Tab $win "Filter"
                        Click-Ctrl (Need $c "DateCheck")
                        $script:Flipped = $false
                    } catch { Write-Log "  체크박스를 누르지 못했습니다: $($_.Exception.Message)" "Red" }
                } else {
                    Write-Log "  잠시 뒤 다시 해봅니다." "Yellow"
                    Start-Sleep -Seconds 3
                }
            }
        }
    }

    if (-not $done) { Write-Log "$stamp 세 번 모두 실패했습니다." "Red"; $failed++ }
    $day = $day.AddDays(1)
}

Write-Log "끝. 성공 $ok / 0건 $empty / 실패 $failed" "Cyan"

# 원본 폴더에는 CSV 만 둔다. 로그가 섞이면 집계 대상과 구분이 안 된다.
if ($LogFile) {
    $logPath = $LogFile
} else {
    $logDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "logs"
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    $logPath = Join-Path $logDir ("collect_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt")
}
New-Item -ItemType Directory -Path (Split-Path -Parent $logPath) -Force | Out-Null
$script:Log | Out-File $logPath -Encoding utf8 -Append
Write-Host "로그: $logPath"
