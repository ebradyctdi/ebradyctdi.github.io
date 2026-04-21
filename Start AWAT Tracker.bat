@echo off
title AWAT Production Tracker
echo ============================================
echo   AWAT Production Tracker - Local Server
echo   CTDI - Alburtis
echo ============================================
echo.
echo Starting server on http://localhost:8080
echo.
echo DO NOT CLOSE THIS WINDOW while using the app.
echo Press Ctrl+C to stop the server.
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$dir = Split-Path -Parent '%~f0'; ^
   $listener = [System.Net.HttpListener]::new(); ^
   $listener.Prefixes.Add('http://localhost:8080/'); ^
   $listener.Start(); ^
   Write-Host 'Server running — opening browser...' -ForegroundColor Green; ^
   Start-Process 'http://localhost:8080/overview.html'; ^
   Write-Host ''; ^
   Write-Host 'Ready. Leave this window open.' -ForegroundColor Green; ^
   Write-Host ''; ^
   while($listener.IsListening) { ^
     $ctx = $listener.GetContext(); ^
     $file = $ctx.Request.Url.LocalPath.TrimStart('/'); ^
     if([string]::IsNullOrEmpty($file) -or $file -eq '/') { $file = 'overview.html' }; ^
     $path = Join-Path $dir $file; ^
     if(Test-Path $path) { ^
       $bytes = [System.IO.File]::ReadAllBytes($path); ^
       $ext = [System.IO.Path]::GetExtension($path).ToLower(); ^
       $ct = switch($ext) { '.html'{'text/html;charset=utf-8'} '.js'{'application/javascript'} '.css'{'text/css'} '.json'{'application/json'} '.png'{'image/png'} '.jpg'{'image/jpeg'} '.ico'{'image/x-icon'} default{'application/octet-stream'} }; ^
       $ctx.Response.ContentType = $ct; ^
       $ctx.Response.ContentLength64 = $bytes.Length; ^
       $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length) ^
     } else { ^
       $ctx.Response.StatusCode = 404; ^
       $bytes = [System.Text.Encoding]::UTF8.GetBytes('File not found'); ^
       $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length) ^
     }; ^
     $ctx.Response.Close() ^
   }"

pause
