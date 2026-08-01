# ------------------------------------------------------------------------------
# デスクトップ操作モジュール (OSネイティブ/UIA)
# ------------------------------------------------------------------------------

# .NET UIAutomation アセンブリのロード
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

# 外部ウィンドウ制御用Win32 APIの宣言
$win32Signature = @"
using System;
using System.Runtime.InteropServices;

namespace Win32Api {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public class Win32Utils {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
    }
}
"@

if (-not ([System.Management.Automation.PSTypeName]"Win32Api.Win32Utils").Type) {
    Add-Type -TypeDefinition $win32Signature
}

# --- 外部アプリケーションウィンドウのアクティブ化 ---
function Switch-AppWindow {
    param ([string]$Name)

    $func = $MyInvocation.MyCommand.Name

    # 完全一致および部分一致によるウィンドウ検索
    $hWnd = [Win32Api.Win32Utils]::FindWindow($null, $Name)
    if ($hWnd -eq [IntPtr]::Zero -or $hWnd -eq $null) {
        $proc = Get-Process | Where-Object { $_.MainWindowTitle -like "*$Name*" } | Select-Object -First 1
        if ($proc) { $hWnd = $proc.MainWindowHandle }
    }

    # ウィンドウの前面化
    if ($hWnd -and $hWnd -ne [IntPtr]::Zero) {
        [Win32Api.Win32Utils]::ShowWindow($hWnd, 9) | Out-Null
        [Win32Api.Win32Utils]::SetForegroundWindow($hWnd) | Out-Null
        return "Focused: $Name"
    }

    throw (New-EngineException -Func $func -Type "未発見" -Message "指定されたウィンドウが見つかりません" -Details $Name)
}

# --- UIAを利用したデスクトップアプリの要素操作 ---
function Invoke-UiaAction {
    param (
        [Parameter(Mandatory=$true)][string]$Action,
        [Parameter(Mandatory=$true)][string]$Name = "",
        [string]$AutomationId = "",
        [string]$Value = "",
        [string]$Mode = "Pattern",
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec,
        [int]$PollIntervalMs = $global:CONFIG.PollIntervalMs
    )

    $func = $MyInvocation.MyCommand.Name

    try {
        # 検索条件の構築（AutomationId優先）
        $condition = if (![string]::IsNullOrEmpty($AutomationId)) {
            New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::AutomationIdProperty,
                $AutomationId
            )
        } else {
            New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::NameProperty,
                $Name
            )
        }

        # タイムアウト付きUIA要素探索
        $root = [System.Windows.Automation.AutomationElement]::RootElement
        $element = $null
        $endTime = (Get-Date).AddSeconds($TimeoutSec)

        while ((Get-Date) -lt $endTime) {
            $element = $root.FindFirst([System.Windows.Automation.TreeScope]::Subtree, $condition)
            if ($element) { break }
            Start-Sleep -Milliseconds $PollIntervalMs
        }

        if (-not $element) { throw (New-EngineException -Func $func -Type "未発見" -Message "指定されたUIA要素が見つかりません" -Details "Timeout: ${TimeoutSec}s") }

        # アクションの実行
        switch ($Action) {
            "Click" {
                try {
                    if ($Mode -eq "Safe") {
                        # --- 安全モード（物理クリック） ---
                        $element.SetFocus()
                        Start-Sleep -Milliseconds 100
                        $wshell = New-Object -ComObject WScript.Shell
                        $wshell.SendKeys(" ")
                    } else {
                        # --- パターンモード（バックグラウンドクリック） ---
                        $invokePattern = $null
                        $legacyPattern = $null

                        if ($element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$invokePattern)) {
                            $invokePattern.Invoke()
                        } elseif ($element.TryGetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern, [ref]$legacyPattern)) {
                            $legacyPattern.DoDefaultAction()
                        } else {
                            throw "InvokePattern および LegacyIAccessiblePattern をサポートしていません"
                        }
                    }
                        Start-Sleep -Milliseconds 300
                } catch {
                    throw (New-EngineException -Func $func -Type "UIAエラー" -Message "対象要素のクリックに失敗しました" -Details $_.Exception.Message)
                }
            }
            "Input" {
                try {
                    if ($Mode -eq "Safe") {
                        # --- 安全モード（物理入力） ---
                        $element.SetFocus()
                        Start-Sleep -Milliseconds 100
                        $wshell = New-Object -ComObject WScript.Shell
                        $wshell.SendKeys($Value)
                    } else {
                        # --- パターンモード（バックグラウンド入力） ---
                        $valuePattern = $null

                        if ($element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
                            $valuePattern.SetValue($Value)
                        } else {
                            throw "ValuePattern をサポートしていません"
                        }
                    }
                    Start-Sleep -Milliseconds 300
                } catch {
                    throw (New-EngineException -Func $func -Type "UIAエラー" -Message "対象要素への入力に失敗しました" -Details $_.Exception.Message)
                }
            }
            "GetText" {
                return $element.Current.Name
            }
            "GetValue" {
                $valuePattern = $null
                if ($element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
                    return $valuePattern.Current.Value
                }
                return ""
            }
            default { throw (New-EngineException -Func $func -Type "引数エラー" -Message "未対応のUIAアクションが指定されました" -Details $Action) }
        }

        return "Action Completed: $Action"

    } catch {
        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "UIA操作中に予期せぬエラーが発生しました" -Details $_.Exception.Message)
    }
}

# --- 「名前を付けて保存」ダイアログの捕捉および保存の実行 ---
function Invoke-UiaSafeSaveAs {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    $dialogNames = @("名前を付けて保存", "Save As", "保存", "Save")
    $hwnd = [IntPtr]::Zero
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    # --- 1. Win32APIで安全にハンドルを取得 ---
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        foreach ($name in $dialogNames) {
            # [重要] PowerShellの $null はWin32API(C#)に渡る際 "" (空文字) に変換されてしまうため、
            # [NullString]::Value を使用するか、標準ダイアログクラス "#32770" を明示指定する
            
            # パターンA: 標準のダイアログクラス名(#32770)で検索
            $hwnd = [Win32Api.Win32Utils]::FindWindow("#32770", $name)
            
            if ($hwnd -eq [IntPtr]::Zero -or $hwnd -eq $null) {
                # パターンB: クラス名問わず、完全な null ポインタとして検索（特殊定数を使用）
                $hwnd = [Win32Api.Win32Utils]::FindWindow([NullString]::Value, $name)
            }

            if ($hwnd -ne [IntPtr]::Zero -and $hwnd -ne $null) { break }
        }
        if ($hwnd -ne [IntPtr]::Zero -and $hwnd -ne $null) { break }
        Start-Sleep -Milliseconds 200
    }

    if ($hwnd -eq [IntPtr]::Zero -or $hwnd -eq $null) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "保存ダイアログのウィンドウが見つかりませんでした")
    }

    # --- 2. ウィンドウを最前面に強制引き上げ ---
    try {
        [Win32Api.Win32Utils]::ShowWindow($hwnd, 9) | Out-Null
        [Win32Api.Win32Utils]::SetForegroundWindow($hwnd) | Out-Null
        Start-Sleep -Milliseconds 300
    } catch {}

    # --- 3. 取得した安全なハンドルからのみUIA要素を生成 ---
    $dialog = $null
    try {
        $dialog = [System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
    } catch {
        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "ハンドルからUIA要素への変換に失敗しました")
    }

    # --- ファイル名入力欄（AutomationId=1001）の取得 ---
    $cndEdit = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::AutomationIdProperty, "1001"
    )
    $editBox = $dialog.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cndEdit)

    if ($null -eq $editBox) {
        Write-DebugLog -Message "[$func] 警告: 入力欄(1001)が未発見。デフォルトフォーカスを利用します。" -Level Warning
    } else {
        try {
            $editBox.SetFocus()
            Start-Sleep -Milliseconds 200
        } catch {}
    }

    # --- 4. クリップボードと物理キーによる絶対安全な操作 ---
    try {
        Write-DebugLog -Message "[$func] 情報: 物理キー(SendKeys)で保存を実行します" -Level Info

        $wshell = New-Object -ComObject WScript.Shell

        # 既存のテキストを全選択して削除 (Ctrl+A -> Delete)
        $wshell.SendKeys("^a")
        Start-Sleep -Milliseconds 100
        $wshell.SendKeys("{DELETE}")
        Start-Sleep -Milliseconds 100

        # クリップボード経由でパスを貼り付け
        Set-Clipboard -Value $FilePath
        Start-Sleep -Milliseconds 100
        $wshell.SendKeys("^v")
        Start-Sleep -Milliseconds 300

        # Enterキーで保存実行
        $wshell.SendKeys("{ENTER}")

    } catch {
        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "パスの入力または保存の実行に失敗しました" -Details $_.Exception.Message)
    }

    # --- 保存完了待機（ファイルロック解除まで） ---
    $waitSw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($waitSw.Elapsed.TotalSeconds -lt 30) {
        if (Test-Path $FilePath) {
            try {
                $stream = [System.IO.File]::Open($FilePath, 'Open', 'Read', 'None')
                $stream.Close()

                return $FilePath
            } catch {}
        }
        Start-Sleep -Milliseconds 300
    }

    throw (New-EngineException -Func $func -Type "Timeout" -Message "ファイル保存の完了確認ができませんでした" -Details $FilePath)
}

# --- アクティブウィンドウへの物理キー送信 ---
function Invoke-DesktopSendKeys {
    param (
        [Parameter(Mandatory=$true)][string]$Keys,
        [int]$WaitMs = 300
    )

    $func = $MyInvocation.MyCommand.Name

    try {
        # 画面のフォーカスが安定するまで少し待機
        Start-Sleep -Milliseconds $WaitMs
        
        $wshell = New-Object -ComObject WScript.Shell
        $wshell.SendKeys($Keys)
        
        Start-Sleep -Milliseconds 300
        return "SendKeys Completed: $Keys"

    } catch {
        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "キーボード送信に失敗しました" -Details $_.Exception.Message)
    }
}

# --- アクティブウィンドウの中央を物理クリックする ---
function Invoke-DesktopCenterClick {
    param ([int]$WaitMs = 500)

    $func = $MyInvocation.MyCommand.Name

    try {
        # 画面のフォーカスが安定するまで少し待機
        Start-Sleep -Milliseconds $WaitMs
        
        # 最前面のウィンドウハンドルを取得
        $hwnd = [Win32Api.Win32Utils]::GetForegroundWindow()
        if ($hwnd -eq [IntPtr]::Zero) {
            throw "最前面のウィンドウが取得できません"
        }
        
        # ウィンドウの矩形領域を取得
        $rect = New-Object Win32Api.RECT
        $res = [Win32Api.Win32Utils]::GetWindowRect($hwnd, [ref]$rect)
        if (-not $res) {
            throw "ウィンドウ領域の取得に失敗しました"
        }
        
        # 画面中央の座標を計算
        $centerX = [int]($rect.Left + (($rect.Right - $rect.Left) / 2))
        $centerY = [int]($rect.Top + (($rect.Bottom - $rect.Top) / 2))
        
        # マウスカーソルを中央へ移動
        [Win32Api.Win32Utils]::SetCursorPos($centerX, $centerY) | Out-Null
        
        # 物理クリックの発火 (LeftDown: 0x0002, LeftUp: 0x0004)
        [Win32Api.Win32Utils]::mouse_event(0x0002, 0, 0, 0, 0)
        [Win32Api.Win32Utils]::mouse_event(0x0004, 0, 0, 0, 0)
        
        return "Clicked Center: X=$centerX, Y=$centerY"

    } catch {
        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "画面中央の物理クリックに失敗しました" -Details $_.Exception.Message)
    }
}
