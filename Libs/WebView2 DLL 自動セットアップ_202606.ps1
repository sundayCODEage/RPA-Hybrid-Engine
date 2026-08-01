# 定期的（たまには）に、AIに照会（確認）すること（2026/06に更新しました。）
<#
AIプロントで、このコードを添付して
WebView2 ポータブル環境用 DLL を下記コードで取得していたが、現在の安定最新に更新したい。
などで、修正してください。
#>

# ==============================================================================
# WebView2 ポータブル環境用 DLL 自動セットアップスクリプト
# ==============================================================================
# 定期的に WebView2 SDK の最新安定版バージョンを確認し、必要に応じて変数を更新してください。

# 1. 実行場所の安全性チェック
$currentDir = $PSScriptRoot
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " WebView2 DLL セットアップツール" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "現在の実行ディレクトリ: $currentDir"

# RPAエンジンのルートフォルダかどうかの簡易判定（sandboxフォルダの有無）
if (-not (Test-Path (Join-Path $currentDir "sandbox") -PathType Container) -and -not (Test-Path (Join-Path $currentDir "Sandbox") -PathType Container)) {
    Write-Host "`n【注意】ここはRPAプロジェクトのルートフォルダではない可能性があります。" -ForegroundColor Yellow
    Write-Host "　本来は 'sandbox' や 'Libs' フォルダが存在する階層で実行する必要があります。" -ForegroundColor Yellow
    $ans = Read-Host "このまま現在の場所に Libs フォルダを作成して処理を続行しますか？ (Y/N)"
    if ($ans -notmatch "^[Yy]$") {
        Write-Host "処理を中断しました。正しいフォルダに移動して再度実行してください。" -ForegroundColor Red
        Start-Sleep -Seconds 3
        exit
    }
}

# 2. 設定：バージョンとディレクトリ
$libsDir = Join-Path $currentDir "Libs"
$tempDir = Join-Path $currentDir "Temp_WebView2_Setup"

# WebView2 SDK のバージョン (最新の安定版を指定)
$wv2Version = "1.0.4022.49"
$packageUrl = "https://www.nuget.org/api/v2/package/Microsoft.Web.WebView2/$wv2Version"

# 抽出対象ファイルのマッピング (PowerShell 5.1 / VBA連携を想定し net462 用を抽出)
$fileMap = @(
    @{ Src = "lib/net462/Microsoft.Web.WebView2.Core.dll";     Dest = "Microsoft.Web.WebView2.Core.dll" }
    @{ Src = "lib/net462/Microsoft.Web.WebView2.WinForms.dll"; Dest = "Microsoft.Web.WebView2.WinForms.dll" }
    @{ Src = "build/native/x64/WebView2Loader.dll";            Dest = "WebView2Loader.dll" } 
)

# 3. 準備
if (-not (Test-Path $libsDir)) { New-Item -ItemType Directory -Path $libsDir | Out-Null }
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
New-Item -ItemType Directory -Path $tempDir | Out-Null

Write-Host "`n--- WebView2 コンポーネント取得開始 ---" -ForegroundColor Cyan
Write-Host "Target Version: $wv2Version"

# 4. パッケージのダウンロードと展開
$zipPath = Join-Path $tempDir "wv2.zip"
$extractPath = Join-Path $tempDir "extract"

Write-Host "1. Downloading WebView2 SDK from NuGet..." -NoNewline
try {
    Invoke-WebRequest -Uri $packageUrl -OutFile $zipPath -ErrorAction Stop
    Write-Host " [Success]" -ForegroundColor Green

    Write-Host "2. Extracting package..." -NoNewline
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
    Write-Host " [Success]" -ForegroundColor Green

    # 5. 必要ファイルの選別とコピー
    Write-Host "3. Installing DLLs to Libs folder..."
    foreach ($item in $fileMap) {
        $sourceFull = Join-Path $extractPath $item.Src
        $destFull = Join-Path $libsDir $item.Dest
        
        Write-Host "  -> $($item.Dest)..." -NoNewline
        if (Test-Path $sourceFull) {
            Copy-Item $sourceFull -Destination $destFull -Force
            Write-Host " [Done]" -ForegroundColor Green
        } else {
            Write-Host " [Error: Not Found in Package]" -ForegroundColor Red
        }
    }
} catch {
    Write-Host " [Failed: $($_.Exception.Message)]" -ForegroundColor Red
}

# 6. 後片付け
Write-Host "4. Cleaning up temporary files..." -NoNewline
Remove-Item $tempDir -Recurse -Force
Write-Host " [Done]" -ForegroundColor Cyan

# 終了メッセージ
Write-Host "`nセットアップ完了！" -ForegroundColor White
Write-Host "以下のファイルが $libsDir に配置されました:"
Get-ChildItem $libsDir | Select-Object Name, Length | Out-String

Write-Host ""
# 実行環境(ISE等)に依存しない安全な待機方法に変更
Read-Host "Enterキーを押して終了してください..."
