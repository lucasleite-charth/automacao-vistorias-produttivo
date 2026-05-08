Passo a Passo – Como Utilizar o Sistema de Vistorias
Objetivo do sistema

O sistema automatiza:
controle de vistorias
geração de pendências
agendamento de revisões
envio automático de lembretes por e-mail

1. Abrir a planilha
Abra a planilha: Controle de Vistorias

2. Liberar permissões (primeira vez)
No Apps Script:
Clique em:
Extensões → Apps Script
Execute a função: Codigo.gs

3. Clique em:
Revisar o acesso
Aceite todas as permissões do Google.

4. Configurar a aba Config na planilha 
Abra a aba: Config

Preencha:
Campo	                    Exemplo
email_para	              cliente@empresa.com

email_cc	                gestor@empresa.com

pastaDriveId	            ID da pasta Drive(não altera)

calendarId                cliente@empresa.com

5. Colocar os arquivos na pasta Drive
Controle Vistorias -> File (obs.solicitar compartilhamento)

adicione os arquivos exportados do sistema de vistoria:
- atividades
- relatório consolidado


6.Executar o pipeline completo

No Apps Script:
Execute manualmente: Code.gs selecionar executar:
*executarPipelineCompleto* 


O que o sistema faz automaticamente
1. Importa os arquivos:
Atualiza:
- Base_Atividades
- Base_Formularios

 2. Processa as vistorias
Atualiza:
- Vistorias_Tratadas

Pendencias
- Regras automáticas
- Vistoria conforme

O sistema agenda:
Nova revisão em 120 dias
Vistoria não conforme

O sistema cria:
lembrete em 15 dias
lembrete em 28 dias

3. Verificar resultados
Aba principal:
Vistorias_Tratadas

Contém:
pontuação
criticidade
revisão futura
responsáveis
vistorias não conformes
lembretes pendentes
4. Envio de e-mails
Enviar relatório consolidado

 8. Automação semanal

O sistema pode rodar automaticamente:
toda segunda-feira
sem necessidade de abrir a planilha


⚠️ Importante
Não alterar:
❌ nomes das abas
❌ nomes das colunas
❌ fórmulas internas

📝 Logs do sistema

A aba da planilha:
Logs

mostra:
importações
erros
processamento
envio de e-mails

🆘 Em caso de erro
Verificar aba:
Logs
Conferir:
permissões
pasta Drive
e-mails
arquivos importados

✅ Fluxo resumido
1. Colocar arquivos na pasta
2. Rodar executarPipelineCompleto
3. Conferir Vistorias_Tratadas
4. Enviar relatório
5. Criar eventos

🚀 Resultado final

O cliente terá:

✔ controle automatizado
✔ acompanhamento de pendências
✔ lembretes automáticos
✔ histórico organizado
✔ redução de trabalho manual
