# DIMSCALE Mail AI — Resumo de Estado e Próximos Passos (2026-10-26)

## 1. Estado Atual
- **Add-in no Outlook:**
  - O separador **"Responder"** está operacional.
  - A adição manual de contactos funciona perfeitamente (o contacto `carla@dimscale.com` já está selecionado).
  - O bloqueio ao clicar em *"Autorizar acesso"* deve-se à política de segurança do tenant da DIMSCALE (exige autorização do administrador para a permissão `Mail.ReadBasic`).
- **Testes de Código:**
  - Suite de testes com **200/200 testes aprovados** (17 ficheiros de teste verdes).

---

## 2. Porque ainda não foram criadas respostas automáticas?
1. **Interruptor desligado no ecrã:** No topo do painel, o botão *"Respostas automáticas"* estava na posição desligada (`Estado: Desligado`).
2. **Apenas para novos e-mails futuros:** O sistema foi desenhado para atuar quando chega um **novo** e-mail vindo daquele remetente (não cria rascunhos sobre e-mails que já estavam na caixa de correio no passado).
3. **Dependência do Servidor (Backend):** Para o servidor processar e-mails em segundo plano (mesmo com o Outlook fechado no teu computador), o administrador precisa de configurar a autorização da aplicação no Microsoft 365.

---

## 3. Mensagem Pronta a Enviar ao Administrador / MMSI
*(Podes copiar e colar diretamente para enviar ao suporte/administrador de TI)*

> **Assunto:** DIMSCALE Mail AI: Consentimento de Administrador para a aplicação (Entra ID)
> 
> Olá,
> 
> Não preciso de permissões de administrador na minha conta de utilizador. O que é necessário é apenas conceder autorização à própria aplicação a nível do tenant da DIMSCALE, para que os utilizadores não fiquem bloqueados com pedidos de aprovação individuais.
> 
> Podem, por favor, efetuar os seguintes passos no **Microsoft Entra admin center**?
> 
> 1. Aceder a **Entra ID > App registrations > All applications >** procurar por **`DIMSCALE AI Email Assistant`** (Client ID: `d157a20c-8820-43ad-8dbf-174a5cade5ea`).
> 2. Entrar em **API permissions**:
>    - Confirmar/Adicionar a permissão delegada: **Microsoft Graph > `Mail.ReadBasic`** (usada apenas para ler a lista de remetentes recentes, sem ler corpo nem anexos dos e-mails).
> 3. Na mesma página de **API permissions**, clicar no botão:  
>    👉 **"Grant admin consent for DIMSCALE"** (Conceder consentimento de administrador para DIMSCALE).
> 4. *(Opcional / Recomendado)* Em **Enterprise apps > "DIMSCALE AI Email Assistant" > Properties**, definir **Assignment required? = Yes** e associar o grupo de utilizadores da DIMSCALE que terão acesso.
> 
> *(O checklist técnico completo com as configurações de Exchange para o servidor encontra-se arquivado em `docs/admin/2026-10-23-mmsi-request.md`)*.
> 
> Mais tarde, quando o add-in estiver finalizado, enviarei o ficheiro `manifest.xml` final para fazerem a publicação centralizada (deploy) no Microsoft 365 Admin Center para todos os utilizadores.
> 
> Podem confirmar assim que o consentimento estiver concedido?
> 
> Obrigado!

---

## 4. O que fazer quando retomaares o trabalho
1. **No Add-in do Outlook:**
   - Clicar no interruptor **"Respostas automáticas"** para ficar **Ligado** (`Estado: Ativo`);
   - Clicar em **Guardar**.
2. **Enviar o e-mail acima ao administrador:**
   - Assim que o admin clicar em *"Grant admin consent"*, o botão *"Autorizar acesso"* no Outlook passa a carregar os remetentes recentes automaticamente sem abrir janelas de erro.
3. **Testar com a Carla:**
   - Pedir à Carla (`carla@dimscale.com`) para enviar um e-mail novo de teste;
   - Verificar a criação do rascunho de resposta automática na pasta **Rascunhos (Drafts)**.
4. **Deploy final:**
   - Gerar e enviar o `manifest.xml` para a TI publicar centralmente.
