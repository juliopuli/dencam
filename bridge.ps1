# bridge.ps1
# Versión 12: Corrección de sintaxis y Escritura Shadow Estable
# Esta versión soluciona el fallo de conexión de la v11.

$port = 8080
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://localhost:$port/")

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {
    $listener.Start()
} catch {
    Write-Host "ERROR: No se pudo iniciar el servidor. Cierra otras ventanas abiertas." -ForegroundColor Red
    pause
    exit
}

Write-Host "--- Puente Local ExcelAtenciones (v12: Shadow Write Estable) activo ---"
Write-Host "Esperando peticiones..."

function Send-Response($context, $content, $contentType = "application/json") {
    if ($null -eq $content) { $content = "[]" }
    if ($content -isnot [string]) { $content = $content | ConvertTo-Json -Depth 10 }
    
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($content)
    $context.Response.ContentType = "$contentType; charset=utf-8"
    $context.Response.ContentLength64 = $buffer.Length
    $context.Response.AddHeader("Access-Control-Allow-Origin", "*")
    $context.Response.AddHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    $context.Response.AddHeader("Access-Control-Allow-Headers", "Content-Type")
    $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
    $context.Response.Close()
}

function Get-DecodedParam($context, $paramName) {
    if ($context.Request.Url.Query -match "$paramName=([^&]+)") {
        return [uri]::UnescapeDataString($matches[1])
    }
    return $null
}

while ($listener.IsListening) {
    $excel = $null
    $wb = $null
    try {
        $context = $listener.GetContext()
        $request = $context.Request
        
        if ($request.HttpMethod -eq "OPTIONS") {
            Send-Response $context ""
            continue
        }

        # --- GET: Cargar Personas o Historial ---
        if ($request.HttpMethod -eq "GET" -and ($request.Url.LocalPath -eq "/people" -or $request.Url.LocalPath -eq "/history")) {
            $isHistory = $request.Url.LocalPath -eq "/history"
            $filePath = Get-DecodedParam $context "path"
            $typeLabel = if($isHistory){"HISTORIAL"}else{"PERSONAS"}
            
            Write-Host "[$typeLabel] Lectura: $filePath"

            if (-not (Test-Path $filePath)) {
                $err = "Archivo no encontrado: $filePath"
                Write-Host $err -ForegroundColor Red
                Send-Response $context (@{error = $err} | ConvertTo-Json)
                continue
            }

            try {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                $excel.AutomationSecurity = 1 
                
                $wb = $excel.Workbooks.Open($filePath, 0, $true)
                
                # Seleccionar la hoja correcta
                $ws = $null
                if ($isHistory) {
                    foreach ($s in $wb.Worksheets) {
                        if ($s.Name -like "*CONSULTA ENFERMERIA*") { $ws = $s; break }
                    }
                } else {
                    foreach ($s in $wb.Worksheets) {
                        if ($s.Name -like "*Mediaci*") { $ws = $s; break }
                    }
                }
                if ($null -eq $ws) { $ws = $wb.ActiveSheet }
                
                Write-Host "  -> Leyendo hoja: $($ws.Name)" -ForegroundColor Cyan
                
                $data = @()
                $rows = $ws.UsedRange.Rows.Count
                
                if ($rows -gt 1) {
                    # La hoja 'Mediación' tiene encabezados agrupados en la fila 1 y nombres en la 2. Empezamos en la 3.
                    # La hoja 'CONSULTA ENFERMERIA' tiene encabezados en la 1. Empezamos en la 2.
                    $startRow = if($ws.Name -like "*Mediaci*"){ 3 } else { 2 }
                    
                    for ($i = $startRow; $i -le $rows; $i++) {
                        if ($isHistory) {
                            $fechaRaw = $ws.Cells.Item($i, 1).Value2
                            $fechaStr = if ($fechaRaw -is [double]) { [DateTime]::FromOADate($fechaRaw).ToString("dd/MM/yyyy HH:mm") } else { $ws.Cells.Item($i, 1).Text }
                            
                            $row = @{
                                fecha = $fechaStr;
                                dni = $ws.Cells.Item($i, 4).Text; # NIE
                                nombre = $ws.Cells.Item($i, 5).Text;
                                apellidos = $ws.Cells.Item($i, 6).Text;
                                tipo = "Enfermería";
                                observaciones = $ws.Cells.Item($i, 15).Text # Motivo de consulta
                            }
                            if ($row.nombre -and $row.nombre -ne "" -and $row.nombre -notlike "*NOMBRE*") { $data += $row }
                        } else {
                            $row = @{
                                hab = $ws.Cells.Item($i, 3).Text;
                                siria = $ws.Cells.Item($i, 4).Text;
                                dni = $ws.Cells.Item($i, 5).Text; # NIE mapped to 'dni' for frontend compatibility
                                nombre = $ws.Cells.Item($i, 6).Text;
                                apellidos = $ws.Cells.Item($i, 7).Text;
                                pais = $ws.Cells.Item($i, 8).Text;
                                fn = $ws.Cells.Item($i, 9).Text;
                                pi = $ws.Cells.Item($i, 13).Text;
                                conf_salida = $ws.Cells.Item($i, 21).Text
                            }
                            if ($row.nombre -and $row.nombre -ne "" -and $row.nombre -notlike "*NOMBRE*") { $data += $row }
                        }
                    }
                }
                
                if ($isHistory -and $data.Count -gt 0) { [array]::Reverse($data) }
                Send-Response $context $data
                Write-Host "  -> Exito. Encontradas $($data.Count) filas." -ForegroundColor Green
            } catch {
                $ex = $_.Exception.Message
                Write-Host "  -> FALLO LECTURA: $ex" -ForegroundColor Red
                Send-Response $context (@{error = $ex} | ConvertTo-Json)
            } finally {
                if ($wb) { $wb.Close($false); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null }
                if ($excel) { $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
                [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
            }
        }

        # --- POST: Escritura Shadow ---
        elseif ($request.HttpMethod -eq "POST") {
            $isAlta = $request.Url.LocalPath -eq "/add-person"
            $isSave = $request.Url.LocalPath -eq "/save"
            
            if ($isAlta -or $isSave) {
                $reader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)
                $body = $reader.ReadToEnd()
                $data = $body | ConvertFrom-Json
                $originalPath = $data.path
                
                $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.IO.Path]::GetRandomFileName() + ".xlsx")
                
                $label = if($isAlta){"ALTA"}else{"REGISTRO"}
                Write-Host "[$label] Escritura Shadow iniciada..."
                
                $success = $false
                try {
                    Write-Host "  -> Clonando original en temporal..." -NoNewline
                    Copy-Item $originalPath $tempPath -Force
                    Write-Host " OK." -ForegroundColor Green
                    
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Visible = $false
                    $excel.DisplayAlerts = $false
                    $excel.AutomationSecurity = 1
                    $wb = $excel.Workbooks.Open($tempPath)
                    
                    $ws = $null
                    if ($isSave) {
                         foreach ($s in $wb.Worksheets) {
                            if ($s.Name -like "*CONSULTA ENFERMERIA*") { $ws = $s; break }
                        }
                    } else {
                        foreach ($s in $wb.Worksheets) {
                           if ($s.Name -like "*Mediaci*") { $ws = $s; break }
                       }
                    }
                    if ($null -eq $ws) { $ws = $wb.ActiveSheet }

                    Write-Host "  -> Escribiendo en hoja: $($ws.Name)" -ForegroundColor Cyan
                    
                    $lastRow = $ws.UsedRange.Rows.Count + 1
                    if ($ws.Cells.Item($lastRow-1, 1).Value2 -eq $null -and $lastRow -gt 1) { $lastRow-- }
                    
                    if ($isAlta) {
                        # No implementamos alta para la nueva estructura compleja por ahora,
                        # pero si se requiere, seguiría el patrón de columnas 3 a 9.
                        $ws.Cells.Item($lastRow, 5) = $data.dni
                        $ws.Cells.Item($lastRow, 6) = $data.nombre
                        $ws.Cells.Item($lastRow, 7) = $data.apellidos
                    } else {
                        # Fecha (1) | Hab. (2) | SIRIA (3) | NIE (4) | NOMBRE (5) | APELLIDOS (6) | NACIONALIDAD (7) | F.N (8) ...
                        # General (11) | Social (12) | Entrega Bienes (13)
                        $ws.Cells.Item($lastRow, 1) = (Get-Date).ToString("dd/MM/yyyy HH:mm")
                        $ws.Cells.Item($lastRow, 2) = $data.hab
                        $ws.Cells.Item($lastRow, 3) = $data.siria
                        $ws.Cells.Item($lastRow, 4) = $data.dni # NIE
                        $ws.Cells.Item($lastRow, 5) = $data.nombre
                        $ws.Cells.Item($lastRow, 6) = $data.apellidos
                        $ws.Cells.Item($lastRow, 7) = $data.pais
                        $ws.Cells.Item($lastRow, 8) = $data.fn
                        
                        # P.I (9) | Conf. salida (10)
                        $ws.Cells.Item($lastRow, 9) = $data.pi
                        $ws.Cells.Item($lastRow, 10) = $data.conf_salida
                        
                        # Mapeo de intervención a columna específica
                        $tipo = $data.tipo.ToLower()
                        if ($tipo -like "*general*") {
                            $ws.Cells.Item($lastRow, 11) = $data.observaciones
                        } elseif ($tipo -like "*social*") {
                            $ws.Cells.Item($lastRow, 12) = $data.observaciones
                        } elseif ($tipo -like "*bienes*" -or $tipo -like "*entrega*") {
                            $ws.Cells.Item($lastRow, 13) = $data.observaciones
                        } else {
                            $ws.Cells.Item($lastRow, 15) = $data.observaciones # Fallback a Col 15
                        }
                    }

                    $wb.Save()
                    $wb.Close($true)
                    $excel.Quit()
                    
                    Write-Host "  -> Reemplazando original con temporal..." -NoNewline
                    Move-Item $tempPath $originalPath -Force
                    Write-Host " OK." -ForegroundColor Green
                    
                    $success = $true
                    Send-Response $context '{"status": "ok"}'
                } catch {
                    $ex = $_.Exception.Message
                    Write-Host "  -> FALLO: $ex" -ForegroundColor Red
                    Send-Response $context (@{error = $ex} | ConvertTo-Json)
                } finally {
                    if ($wb) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null }
                    if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
                    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
                    if (Test-Path $tempPath) { Remove-Item $tempPath -Force -ErrorAction SilentlyContinue }
                }
            }
        }
    } catch {
        Write-Host "Error inesperado: $_"
    }
}
