@echo off
rem Libera a porta 8081 (painel web no WAMP) so para a rede local 192.168.1.x. Clique com o botao direito > Executar como administrador.
netsh advfirewall firewall delete rule name="Projeto invest - painel web" >/dev/null 2>&1
netsh advfirewall firewall add rule name="Projeto invest - painel web" dir=in action=allow protocol=TCP localport=8081 remoteip=192.168.1.0/24 profile=any
pause
