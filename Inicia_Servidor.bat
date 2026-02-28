@echo off
echo ==========================================
echo Iniciando Servidor Local para ExcelAtenciones
echo ==========================================
echo.
echo Para que la autenticacion de Microsoft funcione, 
echo la app debe ejecutarse en http://localhost:8000
echo.
echo NO CIERRES ESTA VENTANA mientras uses la app.
echo.
powershell -ExecutionPolicy Bypass -Command "$p=8000;$l=[System.Net.HttpListener]::new();$l.Prefixes.Add(\"http://localhost:$p/\");$l.Start();Write-Host \"Servidor corriendo en http://localhost:$p\";while($l.IsListening){$c=$l.GetContext();$r=$c.Request;$s=$c.Response;$f=Join-Path $pwd $r.Url.LocalPath.TrimStart('/');if(!(Test-Path $f -PathType Leaf)){$f=Join-Path $pwd 'index.html'};$b=[System.IO.File]::ReadAllBytes($f);$s.ContentLength64=$b.Length;$s.OutputStream.Write($b,0,$b.Length);$s.Close()}"
pause
