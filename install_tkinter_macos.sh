#!/bin/bash
# Script para instalar suporte ao tkinter no macOS com pyenv

echo "=========================================="
echo "Instalação do suporte ao tkinter no macOS"
echo "=========================================="
echo ""

# Verifica se está no macOS
if [[ "$OSTYPE" != "darwin"* ]]; then
    echo "ERRO: Este script é apenas para macOS."
    exit 1
fi

# Verifica se o Homebrew está instalado
if ! command -v brew &> /dev/null; then
    echo "ERRO: Homebrew não está instalado."
    echo "Instale o Homebrew em: https://brew.sh"
    exit 1
fi

# Verifica se o pyenv está instalado
if ! command -v pyenv &> /dev/null; then
    echo "ERRO: pyenv não está instalado."
    echo "Instale o pyenv: brew install pyenv"
    exit 1
fi

echo "1. Instalando Tcl/Tk via Homebrew..."
brew install tcl-tk

if [ $? -ne 0 ]; then
    echo "ERRO: Falha ao instalar tcl-tk."
    exit 1
fi

echo ""
echo "2. Verificando versão atual do Python..."
CURRENT_PYTHON=$(pyenv version-name)
echo "Versão atual: $CURRENT_PYTHON"

if [ -z "$CURRENT_PYTHON" ] || [ "$CURRENT_PYTHON" = "system" ]; then
    echo "ERRO: Nenhuma versão do Python do pyenv está ativa."
    echo "Configure uma versão do Python com: pyenv global 3.14.0"
    exit 1
fi

echo ""
echo "3. Reinstalando Python com suporte ao tkinter..."
echo "Isso pode levar alguns minutos..."

# Obtém o caminho do tcl-tk
TCLTK_PATH=$(brew --prefix tcl-tk)

# Reinstala o Python com suporte ao tkinter
env PATH="$TCLTK_PATH/bin:$PATH" \
    LDFLAGS="-L$TCLTK_PATH/lib" \
    CPPFLAGS="-I$TCLTK_PATH/include" \
    pyenv install --force "$CURRENT_PYTHON"

if [ $? -eq 0 ]; then
    echo ""
    echo "=========================================="
    echo "SUCESSO! tkinter foi instalado corretamente."
    echo "=========================================="
    echo ""
    echo "Teste a instalação executando:"
    echo "python -c \"import tkinter; print('tkinter OK!')\""
else
    echo ""
    echo "ERRO: Falha ao reinstalar o Python."
    echo "Tente manualmente:"
    echo "env PATH=\"\$(brew --prefix tcl-tk)/bin:\$PATH\" pyenv install --force $CURRENT_PYTHON"
    exit 1
fi

