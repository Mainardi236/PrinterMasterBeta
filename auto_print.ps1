$folder = "C:\Users\leao-\Downloads\resto"
$printer = "Lexmark MX710 Series XL"
$sumatra = "C:\Users\leao-\AppData\Local\SumatraPDF\SumatraPDF.exe"

$logFile = "$folder\print_log.txt"

# limpa/cria log
"INICIO: $(Get-Date)" | Out-File $logFile

# pega PDFs em ordem numérica (evita 1,10,11,2)
$files = Get-ChildItem $folder -Filter *.pdf | 
Sort-Object { [regex]::Replace($_.Name, '\D', '') -as [int] }

foreach ($file in $files) {

    try {
        Write-Host "Imprimindo: $($file.Name)"

        & $sumatra -print-to "$printer" $file.FullName

        # pequeno delay para fila de impressão respirar
        Start-Sleep -Milliseconds 1200

        "OK: $($file.Name) - $(Get-Date)" | Out-File $logFile -Append
    }
    catch {
        "ERRO: $($file.Name) - $(Get-Date)" | Out-File $logFile -Append
        Write-Host "Erro ao imprimir $($file.Name)"
        
        # pequena pausa em caso de erro para não “explodir” a fila
        Start-Sleep -Seconds 3
    }
}

"FIM: $(Get-Date)" | Out-File $logFile -Append