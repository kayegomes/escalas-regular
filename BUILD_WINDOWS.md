# Compilação e distribuição do executável Windows

O executável é compilado em um ambiente Windows automatizado pelo GitHub Actions. O ambiente Linux local não gera um `.exe` nativo confiável para distribuição Windows, por isso o workflow usa `windows-latest`.

## Gerar uma compilação

No GitHub, abra a aba **Actions**, selecione **Build Windows Executable** e clique em **Run workflow**. O workflow instala as dependências, executa os testes, compila o aplicativo em modo pasta e publica o artefato `GeradorEscalas-Windows.zip`.

Também é possível iniciar a compilação criando e enviando uma tag no padrão `v*`, por exemplo:

```bash
git tag v2.6.0
git push origin v2.6.0
```

## Distribuição interna

Baixe o artefato ZIP na execução concluída, extraia a pasta `GeradorEscalas` no computador Windows e execute `GeradorEscalas.exe`. A pasta inteira deve ser mantida, pois o executável depende dos arquivos internos gerados pelo PyInstaller.

A planilha `app_config.json` deve permanecer junto ao executável para permitir ajustes de nomes de arquivos e saídas sem recompilar. Os arquivos Excel operacionais podem ficar em qualquer pasta acessível pelo usuário e são selecionados pela interface.

O aplicativo não instala serviço, não exige administrador e mantém a criação de rascunhos no Outlook sob conferência humana. O Windows Defender ou o antivírus corporativo pode exibir um alerta para executáveis internos recém-compilados; nesse caso, a equipe de TI deve validar o artefato e liberar a pasta conforme a política interna.
