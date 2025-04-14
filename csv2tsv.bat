@echo off
setlocal

:: 入力ファイルが指定されていない場合
if "%~1"=="" (
    echo Please drag and drop a CSV file onto this batch file.
    pause
    exit /b
)

:: 入力ファイルのパスを取得
set "inputFile=%~1"
set "outputFile=output.txt"

:: PowerShellでインライン実行
powershell -NoProfile -Command ^
" ^
$data = Import-Csv -Path '%inputFile%'; ^
$headers = $data[0].PSObject.Properties.Name; ^
$transposed = foreach ($h in $headers) { ^
    $row = @($h); foreach ($item in $data) { $row += $item.$h }; $row -join \"`t\" ^
}; ^
$transposed ^| Set-Content -Path '%outputFile%' -Encoding UTF8 ^
"

pause

