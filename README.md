# Compêndio de 86/89 obras — Claude via GitHub Copilot

Esta versão substitui o motor DeepSeek do projeto original pelo **GitHub Copilot SDK**, usando a autenticação do GitHub/Copilot e o modelo `claude-opus-5.5`.

O projeto mantém a lógica principal do original:

- lista de obras em `data/lista_obras.txt`;
- pesquisa web com cache;
- 11 eixos de análise;
- aproximadamente 129 tópicos por obra;
- redação paralela;
- salvamento tópico a tópico;
- retomada após interrupção;
- DOCX individual por obra;
- índice geral;
- artefato ZIP no GitHub Actions.

## Importante sobre as 89 obras

O arquivo fornecido no projeto original contém **86 obras**, das quais 3 já aparecem concluídas no estado enviado. Portanto, a base atual tem 83 obras restantes. O programa aceita qualquer quantidade de linhas em `data/lista_obras.txt`; basta acrescentar as 3 obras que faltam para chegar às 89.

## Como usar no GitHub

1. Crie um repositório privado no GitHub.
2. Suba estes arquivos.
3. Confirme que sua conta/repositório possui GitHub Copilot com créditos de IA disponíveis e acesso ao Claude Opus 5.5.
4. Em **Actions**, execute `Gerar compêndio com Claude`.
5. Escolha:
   - `limite_obras`: quantidade a processar nessa rodada;
   - `pular`: quantidade inicial a ignorar;
   - `palavras_por_topico`: padrão 750;
   - `paralelismo`: padrão 6;
   - `modelo`: padrão `claude-opus-5.5`.
6. Ao final, baixe o artefato `compendio-docx`.

Para repositório pessoal, o workflow usa o segredo `COPILOT_GITHUB_TOKEN`, que deve conter um fine-grained personal access token da sua conta pessoal com a permissão de conta `Copilot Requests`. A permissão `copilot-requests: write` no workflow é necessária para o mecanismo de Copilot, mas o token pessoal é o que autentica as solicitações no repositório pessoal.

## Rodar localmente

```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt

# estando autenticado no GitHub/Copilot:
python scripts/gerar_compendio.py --modelo claude-opus-5.5 --limite-obras 1
```

## Segurança

Não coloque `ANTHROPIC_API_KEY`, senha ou token pessoal no código. Esta versão foi desenhada para usar a autenticação do GitHub/Copilot.

## Retomada

O estado fica em `saida/_estado.json` e os tópicos parciais ficam em `saida/_parciais/`. Se uma execução parar, rode novamente. Tópicos já concluídos não são reescritos.
