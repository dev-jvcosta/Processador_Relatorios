# Como Publicar no GitHub

## Passo 1: Criar o Repositório no GitHub

1. Acesse https://github.com e faça login
2. Clique no botão **"+"** no canto superior direito e selecione **"New repository"**
3. Preencha:
   - **Repository name:** `Processador_Relatorios` (ou outro nome de sua preferência)
   - **Description:** "Processador Unificado de Relatórios - Ferramenta de automação para processamento e geração de relatórios gerenciais"
   - **Visibility:** Escolha **Public** ou **Private**
   - **NÃO marque** "Initialize this repository with a README" (já temos um)
   - **NÃO adicione** .gitignore ou license (já temos)
4. Clique em **"Create repository"**

## Passo 2: Conectar e Fazer Push

Após criar o repositório, você pode usar o script automatizado ou fazer manualmente:

### Método Rápido (Recomendado):

Execute o script:
```bash
push_to_github.bat
```

O script irá pedir a URL do seu repositório e fazer o push automaticamente.

### Método Manual:

#### Opção A: Usando HTTPS (Recomendado para iniciantes)

```bash
# Adicionar o repositório remoto (substitua SEU_USUARIO pelo seu username do GitHub)
git remote add origin https://github.com/SEU_USUARIO/Processador_Relatorios.git

# Renomear branch para main (se necessário)
git branch -M main

# Fazer push do código
git push -u origin main
```

#### Opção B: Usando SSH (Se você já configurou SSH keys)

```bash
# Adicionar o repositório remoto (substitua SEU_USUARIO pelo seu username do GitHub)
git remote add origin git@github.com:SEU_USUARIO/Processador_Relatorios.git

# Renomear branch para main (se necessário)
git branch -M main

# Fazer push do código
git push -u origin main
```

## Passo 3: Verificar

Após o push, acesse seu repositório no GitHub e verifique se todos os arquivos foram enviados corretamente.

## Comandos Úteis para o Futuro

```bash
# Ver status das alterações
git status

# Adicionar arquivos modificados
git add .

# Fazer commit
git commit -m "Descrição das alterações"

# Enviar para o GitHub
git push

# Ver histórico de commits
git log

# Ver repositórios remotos configurados
git remote -v
```

## Solução de Problemas

### Erro de autenticação no push
Se aparecer erro de autenticação, você pode:
1. Usar **Personal Access Token** (recomendado):
   - Vá em GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)
   - Gere um novo token com permissão `repo`
   - Use o token como senha quando solicitado

2. Ou configurar SSH keys:
   - Siga o guia: https://docs.github.com/en/authentication/connecting-to-github-with-ssh

### Erro "remote origin already exists"
Se já existe um remote, remova primeiro:
```bash
git remote remove origin
git remote add origin https://github.com/SEU_USUARIO/Processador_Relatorios.git
```
