@echo off

:: Desinstala os pacotes necessarios via pip
echo Iniciando a desinstalacao dos pacotes...
pip uninstall -y requests beautifulsoup4 pandas winshell pywin32 pyinstaller || (
    echo Ocorreu um erro ao desinstalar os pacotes.
    pause
    exit /b 1
)

echo Todos os pacotes foram desinstalados com sucesso. Pressione ENTER para continuar.

pause

