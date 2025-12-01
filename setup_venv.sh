#!/bin/bash
# Script para criar o ambiente virtual no Linux/Mac

echo "Criando ambiente virtual..."

# Tenta usar python3
if command -v python3 &> /dev/null; then
    python3 -m venv venv
    if [ $? -eq 0 ]; then
        echo "Ambiente virtual criado com sucesso usando 'python3'!"
        source venv/bin/activate
        pip install -r requirements.txt
        echo ""
        echo "Ambiente virtual configurado com sucesso!"
        echo "Para ativar, execute: source venv/bin/activate"
        exit 0
    fi
fi

# Tenta usar python
if command -v python &> /dev/null; then
    python -m venv venv
    if [ $? -eq 0 ]; then
        echo "Ambiente virtual criado com sucesso usando 'python'!"
        source venv/bin/activate
        pip install -r requirements.txt
        echo ""
        echo "Ambiente virtual configurado com sucesso!"
        echo "Para ativar, execute: source venv/bin/activate"
        exit 0
    fi
fi

echo "ERRO: Não foi possível criar o ambiente virtual."
echo "Por favor, verifique se o Python está instalado."
exit 1

