o@echo off
setlocal

:: Instala os pacotes necessarios via pip
echo Iniciando a instalacao dos pacotes necessarios...
pip install requests beautifulsoup4 pandas winshell pywin32 pyinstaller || (
    echo Ocorreu um erro ao instalar os pacotes.
    pause
    exit /b 1
)

echo Todos os pacotes foram instalados com sucesso. Pressione ENTER para continuar

pause

