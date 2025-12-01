@echo off
REM Script para fazer push do projeto para o GitHub
REM Execute este script APOS criar o repositório no GitHub

echo ========================================
echo  Publicar no GitHub
echo ========================================
echo.

REM Verifica se já existe remote
git remote get-url origin >nul 2>&1
if %errorlevel% == 0 (
    echo Repositorio remoto ja configurado:
    git remote get-url origin
    echo.
    echo Deseja atualizar? (S/N)
    set /p resposta=
    if /i not "%resposta%"=="S" (
        echo Operacao cancelada.
        pause
        exit /b 0
    )
    echo.
    echo Digite a URL do seu repositorio GitHub:
    echo Exemplo: https://github.com/SEU_USUARIO/Processador_Relatorios.git
    set /p repo_url=
    git remote set-url origin %repo_url%
) else (
    echo Digite a URL do seu repositorio GitHub:
    echo Exemplo: https://github.com/SEU_USUARIO/Processador_Relatorios.git
    set /p repo_url=
    git remote add origin %repo_url%
)

echo.
echo Verificando branch...
git branch -M main

echo.
echo Fazendo push para o GitHub...
echo.
git push -u origin main

if %errorlevel% == 0 (
    echo.
    echo ========================================
    echo  Sucesso! Projeto publicado no GitHub!
    echo ========================================
    echo.
) else (
    echo.
    echo ========================================
    echo  Erro ao fazer push
    echo ========================================
    echo.
    echo Possiveis causas:
    echo 1. Repositorio nao foi criado no GitHub ainda
    echo 2. Problema de autenticacao
    echo 3. URL do repositorio incorreta
    echo.
    echo Consulte o arquivo GITHUB_SETUP.md para mais detalhes.
    echo.
)

pause

