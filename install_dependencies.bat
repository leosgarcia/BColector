@echo off
echo Verificando se o Python está instalado...

python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo Python não está instalado. Por favor, instale Python 3.x primeiro.
    exit /b 1
)

echo Instalando o pip...
python -m ensurepip --default-pip

echo Instalando pacotes necessários...
python -m pip install --upgrade pip
pip install pandas psutil tldextract openpyxl

echo Todas as dependências foram instaladas com sucesso!
pause
