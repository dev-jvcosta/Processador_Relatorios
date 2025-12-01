@echo off
REM Script para criar o ambiente virtual no Windows
REM Tenta diferentes métodos para encontrar o Python

echo ========================================
echo  Configuracao do Ambiente Virtual
echo ========================================
echo.

REM Verifica se o venv já existe
if exist "venv" (
    echo AVISO: O ambiente virtual 'venv' ja existe!
    echo Deseja recriar? (S/N)
    set /p resposta=
    if /i not "%resposta%"=="S" (
        echo Operacao cancelada.
        pause
        exit /b 0
    )
    echo Removendo ambiente virtual antigo...
    rmdir /s /q venv 2>nul
    echo.
)

echo Criando ambiente virtual...
echo.

REM Tenta usar python diretamente
echo [1/3] Tentando usar 'python'...
python -m venv venv 2>nul
if %errorlevel% == 0 (
    echo [OK] Ambiente virtual criado com sucesso usando 'python'!
    goto :install
)

REM Tenta usar caminho completo comum do Python 3.14
echo [2/3] Tentando usar caminho completo do Python 3.14...
"C:\Program Files\Python314\python.exe" -m venv venv 2>nul
if %errorlevel% == 0 (
    echo [OK] Ambiente virtual criado com sucesso usando caminho completo!
    goto :install
)

REM Tenta usar py launcher (última tentativa, pois pode ter problemas)
echo [3/3] Tentando usar 'py' launcher...
py -m venv venv 2>nul
if %errorlevel% == 0 (
    echo [OK] Ambiente virtual criado com sucesso usando 'py'!
    goto :install
)

echo.
echo [ERRO] Nao foi possivel criar o ambiente virtual.
echo.
echo Possiveis solucoes:
echo 1. Verifique se o Python esta instalado: python --version
echo 2. Tente instalar o Python de: https://www.python.org/downloads/
echo 3. Certifique-se de que o Python esta no PATH do sistema
echo.
pause
exit /b 1

:install
echo.
echo ========================================
echo  Instalando dependencias
echo ========================================
echo.

REM Verifica se requirements.txt existe
if not exist "requirements.txt" (
    echo [AVISO] Arquivo requirements.txt nao encontrado!
    echo Criando requirements.txt com dependencias basicas...
    (
        echo pandas
        echo openpyxl
        echo reportlab
        echo numpy
    ) > requirements.txt
    echo [OK] requirements.txt criado!
    echo.
)

echo Instalando pacotes do requirements.txt...
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo [ERRO] Nao foi possivel ativar o ambiente virtual.
    pause
    exit /b 1
)

pip install --upgrade pip
pip install -r requirements.txt

if %errorlevel% == 0 (
    echo.
    echo ========================================
    echo  Configuracao concluida com sucesso!
    echo ========================================
    echo.
    echo Para ativar o ambiente virtual, execute:
    echo   venv\Scripts\activate
    echo.
    echo Para executar o programa:
    echo   python main.py
    echo.
) else (
    echo.
    echo [ERRO] Falha ao instalar dependencias.
    echo Verifique o arquivo requirements.txt
    echo.
)

pause

