# NSIPS Converter パッチ公開スクリプト (PowerShell)
# 使い方:  .\publish_patch.ps1 -Version 1.0.1
# 手順:  VERSION更新 → build_exe.py でexeビルド → GitHub Release作成(exe添付)
# 事前に CHANGELOG.md に新バージョンの項目を追記しておくこと。

param(
    [Parameter(Mandatory = $true)]
    [string]$Version
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

$repo = "kngmincorp-stack/NSIPSConverter"
$tag  = "v$Version"

Write-Host "==== NSIPS Converter publish $tag ====" -ForegroundColor Cyan

# CHANGELOG に該当バージョンがあるか確認
$changelog = Get-Content -Path "CHANGELOG.md" -Raw -Encoding UTF8
if ($changelog -notmatch [regex]::Escape("v$Version")) {
    Write-Host "⚠ CHANGELOG.md に v$Version の記載がありません。先に追記してください。" -ForegroundColor Yellow
    exit 1
}

# VERSION 退避 & 更新
$oldVersion = (Get-Content -Path "VERSION" -Raw).Trim()
[System.IO.File]::WriteAllText((Join-Path $PSScriptRoot "VERSION"), $Version)
Write-Host "VERSION: $oldVersion -> $Version"

try {
    # ビルド
    Write-Host "building exe..." -ForegroundColor Cyan
    python build_exe.py
    if ($LASTEXITCODE -ne 0) { throw "build failed" }

    $exe = Join-Path $PSScriptRoot "dist\NSIPSConverter.exe"
    if (-not (Test-Path $exe)) { throw "exe not found: $exe" }

    # リリース作成 (既存タグがあれば資産差し替え)
    Write-Host "creating GitHub release $tag..." -ForegroundColor Cyan
    $exists = (gh release view $tag --repo $repo 2>$null)
    if ($LASTEXITCODE -eq 0) {
        gh release upload $tag $exe --repo $repo --clobber
    } else {
        gh release create $tag $exe --repo $repo --title $tag --notes "See CHANGELOG.md"
    }
    if ($LASTEXITCODE -ne 0) { throw "gh release failed" }

    Write-Host "==== 公開完了: $tag ====" -ForegroundColor Green
}
catch {
    # 失敗時は VERSION を戻す
    [System.IO.File]::WriteAllText((Join-Path $PSScriptRoot "VERSION"), $oldVersion)
    Write-Host "✖ 失敗のため VERSION を $oldVersion に戻しました: $_" -ForegroundColor Red
    exit 1
}
