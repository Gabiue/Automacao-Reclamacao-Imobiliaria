# Sistema de Reclamações Imobiliárias

Automação em **Google Apps Script** que transforma reclamações de clientes de uma imobiliária em um **Termo de Manutenção de Imóvel em PDF**, salvo no Drive e enviado por e-mail para a equipe de manutenção, sem ninguém precisar abrir a planilha.

Está **em produção** numa empresa do setor imobiliário. Antes, cada reclamação chegava pelo WhatsApp e era passada à mão para um documento e impressa.

## Como funciona

```
Cliente preenche o Google Forms
        │
        ▼
Aba "Respostas do formulário 1"  (caixa de entrada)
        │  trigger "ao enviar formulário"
        ▼
Move a linha para a aba do mês (ex: 09-2026), criando a aba se não existir
        │
        ▼
Gera o PDF do termo  →  salva no Drive  →  envia por e-mail  →  marca como PROCESSADO
        │
        └── se algo falhar: a linha fica pendente e o dono do script recebe um aviso
```

Cada aba mensal tem duas colunas de controle:

| Coluna | Uso |
|---|---|
| K · STATUS | Controlada pelo sistema: em branco = pendente, `PROCESSADO` = PDF já enviado |
| L · RESOLVIDO | Checkbox que a construtora marca quando conclui a manutenção |

## Decisões que valem destacar

- **E-mail antes de marcar como processado.** Se a execução morrer no meio, o pior caso é um e-mail duplicado, nunca uma reclamação marcada como enviada sem ter saído.
- **Só processa o que acabou de chegar.** O processamento automático recebe a lista exata de linhas criadas pelo trigger, então o histórico antigo nunca é reenviado por engano.
- **Lock na organização das abas** (`LockService`), para duas respostas quase simultâneas não mexerem na caixa de entrada ao mesmo tempo.
- **Proteção contra zero à esquerda.** O Sheets converte CPF, contrato e telefone em número e apaga os zeros do começo. As colunas são fixadas como Texto Simples, e um CPF que já chegou sem o zero é recusado com uma mensagem explicando como corrigir.
- **Imagens em paralelo.** As fotos anexadas no Forms são baixadas numa única chamada (`UrlFetchApp.fetchAll`) e reduzidas pela miniatura do Drive, o que deixou cada PDF bem mais leve.

## Formulário

O Google Forms tem 8 perguntas, que precisam ficar nesta ordem porque o script lê a planilha por posição:

| Coluna | Pergunta | Tipo |
|---|---|---|
| A | Carimbo de data/hora | preenchido pelo Forms |
| B | Nome completo | resposta curta |
| C | CPF (somente números) | resposta curta |
| D | Nº do contrato (somente números, 13 dígitos) | resposta curta |
| E | Endereço | resposta curta |
| F | Telefone para contato | resposta curta |
| G | Condomínio e apartamento | resposta curta |
| H | Motivo da reclamação | resposta curta |
| I | Fotos do problema | upload de arquivo |
| J | *(reservada, sem pergunta)* | fica vazia |
| K | STATUS | escrita pelo script |
| L | RESOLVIDO | checkbox criado pelo script |

A coluna J é lida como "fotos do contrato", que hoje não existe no formulário. Ela fica vazia e mantém o STATUS na coluna K, onde as planilhas em produção já guardam o histórico.

## Configuração

1. Crie o formulário com as perguntas da tabela acima e vincule a uma planilha.
2. Na planilha, abra **Extensões → Apps Script** e cole o conteúdo de [`Code.gs`](Code.gs).
3. Em **Configurações do projeto (engrenagem) → Propriedades do script**, cadastre:

   | Propriedade | Obrigatória | O que é |
   |---|---|---|
   | `ID_PASTA_RECLAMACOES` | sim | ID da pasta do Drive onde os PDFs são salvos (o trecho da URL depois de `/folders/`) |
   | `EMAIL_EQUIPE_MANUTENCAO` | sim | E-mail que recebe o termo em PDF |
   | `EMAIL_AVISO_FALHA` | não | Quem recebe o aviso de falha. Em branco, vai para o dono do script |

4. Recarregue a planilha. No menu **🏠 Reclamações**:
   1. **Instalar trigger automático** (autorize as permissões)
   2. **Organizar respostas por mês**, para mover o histórico que já existe
   3. **Testar Sistema**, que confere as propriedades, a aba do mês e a pasta do Drive

Nenhum dado sensível fica no código: IDs e e-mails moram nas Propriedades do script, e este repositório não contém dados de clientes.

## Stack

Google Apps Script · Google Sheets · Google Forms · Google Drive · Gmail
