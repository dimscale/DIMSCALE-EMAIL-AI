# Pedido ao administrador do tenant (MMSI), versão final de 2026-11-02

- **Estado:** pronto a enviar pelo utilizador quando disser que o add-in está pronto. Atualizado em 2026-11-02 (ficheiro original de 2026-10-23). Substitui todos os pedidos anteriores.
- **Fontes Microsoft:** no fim de cada secção. Nomes dos menus em inglês, como aparecem nos portais.
- **Nada aqui é secreto.** Nenhuma chave, segredo ou chave privada passa por este documento ou por e-mail. Só o ficheiro público `drafts.crt` é enviado.
- **Estado do add-in:** manifesto **1.0.2.0**, nome **"Assistente IA"**, **um** botão no friso "DIMSCALE Mail" (grupo "DIMSCALE MAIL"), permissão `ReadItem` apenas. Não há botão "Resumo" nem segunda página.

---

**Assunto:** DIMSCALE Mail AI: checklist final (Entra, Exchange, implementação centralizada)

Olá. Segue a lista final para o DIMSCALE Mail AI. Os itens **[devolver]** são valores que precisamos de volta (nenhum é secreto). Regra fixa do projeto: **nunca** `Mail.Send`, nunca envio, eliminação ou movimentação automática de e-mail; o sistema só cria rascunhos.

> **Confirmar o que já foi feito (nota interna, apagar antes de enviar).** A MMSI respondeu em 2026-10-29 ("Permissões atualizadas") sobre o consentimento de `Mail.ReadBasic`, mas não ficou confirmado o que foi feito exatamente. Marcar o que já está feito e enviar só o resto:
> - [ ] A1 `Mail.ReadBasic` delegada adicionada à app
> - [ ] A1 `Mail.ReadWrite`, `User.Read`, `offline_access` removidas da lista
> - [ ] A2 consentimento de administrador concedido
> - [ ] A3 separador Admin consent só com `access_as_user`, `openid`, `profile`, `Mail.ReadBasic`
> - [ ] A4 Assignment required = Yes, com o grupo atribuído (e P1/P2 confirmado)
> - [ ] B grupo com e-mail criado, utilizadores piloto como membros diretos
> - [ ] C app "DIMSCALE Mail AI Drafts" criada com o certificado
> - [ ] D papéis Exchange atribuídos e verificados
> - [ ] E/F implementação centralizada do manifesto

### A. App existente "DIMSCALE AI Email Assistant" (client ID `d157a20c-8820-43ad-8dbf-174a5cade5ea`)
1. **Entra admin center > Entra ID > App registrations > All applications >** "DIMSCALE AI Email Assistant" **> API permissions**. A lista deve ficar exatamente:
   - a API da própria app: `access_as_user` (delegada);
   - Microsoft Graph, delegadas: `openid`, `profile` e **`Mail.ReadBasic`** (lista os remetentes recentes da Caixa de Entrada, sem corpo nem pré-visualização).
   - **Remover** (… > **Remove permission**): `Mail.ReadWrite`, `User.Read`, `offline_access`.
   - Para adicionar: **Add a permission > Microsoft Graph > Delegated permissions > Mail.ReadBasic**.
2. Na mesma página: **Grant admin consent for <tenant>**. É obrigatório porque a app terá *Assignment required* (ponto 4).
3. **Entra ID > Enterprise apps > All applications >** "DIMSCALE AI Email Assistant" **> Security > Permissions >** separador **Admin consent**: confirmar que só restam `access_as_user`, `openid`, `profile`, `Mail.ReadBasic`. Se aparecerem `Mail.ReadWrite`, `User.Read` ou `offline_access`: **… > Revoke permission**. **Não** usar os scripts "remove all" dessa página (apagam também `access_as_user`).
4. **Enterprise apps >** "DIMSCALE AI Email Assistant" **> Properties > Assignment required? = Yes > Save**. Depois **Users and groups > Add user/group >** o grupo dos utilizadores DIMSCALE **> Assign**. Requer Entra ID **P1 ou P2**; grupos aninhados não contam.
   - **[devolver]** nome do grupo, e se o tenant tem P1/P2.
5. **[devolver]** confirmação dos pontos 1 a 4 (captura do separador Admin consent sem dados pessoais, ou a lista em texto).

Fontes: https://learn.microsoft.com/entra/identity/enterprise-apps/grant-admin-consent, https://learn.microsoft.com/entra/identity/enterprise-apps/manage-application-permissions, https://learn.microsoft.com/entra/identity/enterprise-apps/assign-user-or-group-access-portal, https://learn.microsoft.com/entra/identity/enterprise-apps/application-properties, https://learn.microsoft.com/graph/permissions-reference

### B. Grupo para os rascunhos automáticos (Exchange)
6. Um grupo **com e-mail** (Microsoft 365, segurança com e-mail ou lista de distribuição), por exemplo `DIMSCALE Mail AI Auto-Drafts`, com os utilizadores piloto como **membros diretos** (grupos aninhados não contam no filtro `MemberOfGroup`). Pode ser o grupo do ponto 4 se tiver e-mail.
   - **[devolver]** nome do grupo e utilizadores piloto.

### C. App "DIMSCALE Mail AI Drafts" (só servidor, sem utilizador)
7. **Entra ID > App registrations > New registration**: nome `DIMSCALE Mail AI Drafts`; **Accounts in this organizational directory only**; **sem Redirect URI**.
8. **API permissions**: remover o `User.Read` por omissão. **Não adicionar nem consentir nenhuma permissão de aplicação do Microsoft Graph** (nem `Mail.Read`, nem `Mail.ReadWrite`, nem `Mail.Send`). O acesso ao correio é dado só no Exchange (secção D), limitado ao grupo; um consentimento Entra valeria para todas as caixas e anularia o limite.
9. **Certificates & secrets > Certificates > Upload certificate**: o ficheiro público **`drafts.crt`** que enviaremos (a chave privada nunca sai do nosso lado). **Não** criar client secret.
10. **[devolver]** Application (client) ID (será o nosso `GRAPH_APP_CLIENT_ID`); thumbprint e data de expiração do certificado; e, em **Enterprise apps >** `DIMSCALE Mail AI Drafts`, o **Object ID** do service principal (não é o Object ID de App registrations).

Fontes: https://learn.microsoft.com/entra/identity-platform/certificate-credentials, https://learn.microsoft.com/entra/identity-platform/v2-oauth2-client-creds-grant-flow

### D. Exchange Online PowerShell (Organization Management)
```powershell
Connect-ExchangeOnline
New-ServicePrincipal -AppId <Application (client) ID> -ObjectId <Object ID do service principal> -DisplayName "DIMSCALE Mail AI Drafts"
$dn = (Get-Group "DIMSCALE Mail AI Auto-Drafts").DistinguishedName
New-ManagementScope -Name "DIMSCALE Mail AI Auto-Drafts users" -RecipientRestrictionFilter "MemberOfGroup -eq '$dn'"
New-ManagementRoleAssignment -App <Object ID do service principal> -Role "Application Mail.Read" -CustomResourceScope "DIMSCALE Mail AI Auto-Drafts users"
New-ManagementRoleAssignment -App <Object ID do service principal> -Role "Application Mail.ReadWrite" -CustomResourceScope "DIMSCALE Mail AI Auto-Drafts users"
```
11. **Nunca** atribuir `Application Mail.Send`, `Application Mail Full Access`, `Application Exchange Full Access`, `Application EWS.AccessAsApp` ou `Application SMTP.SendAsApp`. (`Application Mail.ReadWrite` não inclui envio.)
12. Verificação, **[devolver]** o resultado dos três:
```powershell
Test-ServicePrincipalAuthorization -Identity <Object ID do service principal> -Resource <caixa de um utilizador piloto>   # esperado InScope = True nos dois papéis
Test-ServicePrincipalAuthorization -Identity <Object ID do service principal> -Resource <caixa fora do grupo>              # esperado InScope = False
Get-ManagementRoleAssignment -RoleAssignee <Object ID do service principal> | Format-Table Name,Role,CustomResourceScope
```
13. As alterações demoram **30 minutos a 2 horas** a aplicar (cache do Exchange).

Fonte: https://learn.microsoft.com/exchange/permissions-exo/application-rbac

### E. Implementação centralizada do add-in (manifesto final)
14. O manifesto final **1.0.2.0** está publicado em `https://dimscale.github.io/DIMSCALE-EMAIL-AI/manifest.xml` (gerado pelo CI). Nome do add-in: **"Assistente IA"**. Traz **um** botão "DIMSCALE Mail" no grupo "DIMSCALE MAIL" do friso do Outlook, painel fixável, permissão `ReadItem`.
15. **Microsoft 365 admin center > Settings > Integrated apps > Add-ins** (ligação perto do topo) **> Deploy Add-in > Next >** carregar o manifesto (**Upload custom apps**, ficheiro ou URL acima) **> Upload >** escolher **Specific users/groups** (recomendado: o grupo piloto, depois alargar; a Microsoft recomenda atribuir por grupos e implementar por fases) **> Deploy**.
16. Pode demorar: a Microsoft indica que os add-ins **podem demorar 24 a 72 horas a aparecer no friso**; os utilizadores podem ter de reiniciar o Outlook.
17. Futuras alterações de código chegam aos utilizadores sem intervenção. Só uma alteração do manifesto pede novo upload (selecionar o add-in > **Update**).
18. Funções que podem aceder ao portal: Global admin, Exchange admin (se tiver também Application Administrator ou a propriedade "App Registration" ativa), entre outras. Atribuições só a grupos de topo (sem grupos-pai).

Fontes: https://learn.microsoft.com/microsoft-365/admin/manage/manage-deployment-of-add-ins, https://learn.microsoft.com/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps

### F. Paragem de emergência (para referência)
- Do nosso lado (imediato): desligamos a funcionalidade no backend; cada utilizador também a pode desligar no separador "Automático".
- Do vosso lado: apagar o certificado da app `DIMSCALE Mail AI Drafts` (os tokens existentes expiram em cerca de 1 hora) e/ou `Remove-ManagementRoleAssignment` nas duas atribuições (30 min a 2 h).

### Perguntas
- O tenant aceita certificados para apps (validade e política de rotação)?
- Há Conditional Access para workload identities, ou outra política que bloqueie um token só de aplicação?
- Contacto para emergências?

Obrigado.

---

**Valores a devolver (resumo, para o utilizador):** nome do grupo e P1/P2 (A4); grupo e utilizadores piloto (B); `GRAPH_APP_CLIENT_ID` (client ID da app Drafts), thumbprint, validade do certificado, Object ID do service principal (C); saídas dos três comandos (D); confirmação do upload do manifesto (E).
