@echo off
echo A abrir porta 8050 no Windows Firewall...
netsh advfirewall firewall add rule name="LUGEST Impulse API 8050" dir=in action=allow protocol=TCP localport=8050
echo.
echo Se nao deu erro, a porta 8050 ficou autorizada.
pause
