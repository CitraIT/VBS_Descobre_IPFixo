# VBS_Descobre_IPFixo
Script para identificar máquinas com IP fixo na rede


1) Crie uma pasta compartilhada com duas subpastas: scripts e arquivos.
## Cole os dois scripts dentro da pasta scripts, e dê permissão de somente leitura para todos os usuários.
## Edite os scripts e cole os respectivos caminhos de rede destas pastas, como no exemplo (yt video).
Crie uma política de grupo, nas configurações de usuário, preferências, adicione uma tarefa agendada do tipo Imediata (at least windows 7).
Na ação da tarefa agendada, no campo programa insira cscript.exe, no argumento cole o caminho do script descobre_ip_fixo.vbs (exemplo: \\servidor\pasta\scripts\descobre_ip_fixo.vbs).
Salve a tarega agendada.
Aguarde as máquinas atualizarem a política de grupo.
Edite o script relatorio_Maquinas_ip_fixo.ps1, altere o caminho da rede também.
Execute pelo powershell o script relatorio_maquinas_ip_fixo.ps1 e ele vai mostrar numa tabela a lista das máquinas e sua configuração de rede.
Use o filtro do grid para exibir somente as máquinas com IP fixo (campo DHCPEnabled é igual a 0).

Ficou com dúvida? Assista o vídeo com o passo-a-passo.
