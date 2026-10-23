# Pedido ao administrador do tenant (MMSI), 2026-10-23

- **Estado:** rascunho pronto a enviar pelo utilizador. Substitui os pedidos anteriores (status.md "Pending with the tenant admin", plano F-013 P1, rascunho em F-015).
- **Fontes Microsoft:** citadas no fim de cada secção. Nomes dos menus em inglês, como aparecem nos portais.
- **Nada aqui é secreto.** Nenhuma chave, segredo ou chave privada passa por este documento ou por e-mail.

---

**Assunto:** DIMSCALE Mail AI: checklist técnica (Entra + Exchange), 2026-10-23

Olá. Segue a lista consolidada para o DIMSCALE Mail AI. Os itens marcados **[devolver]** são valores que precisamos de volta (nenhum é secreto). Regra fixa do projeto: **nunca** `Mail.Send`, nunca envio, eliminação ou movimentação automática de e-mail; o sistema só cria rascunhos.

### A. App existente "DIMSCALE AI Email Assistant" (client ID `d157a20c-8820-43ad-8dbf-174a5cade5ea`)
1. **Entra admin center > Entra ID > App registrations > All applications >** "DIMSCALE AI Email Assistant" **> API permissions**. A lista configurada deve ficar exatamente:
   - a API da própria app: `access_as_user` (delegada);
   - Microsoft Graph, delegadas: `openid` e `profile` (início de sessão), e **`Mail.ReadBasic`** (nova: lista os remetentes recentes da Caixa de Entrada, sem corpo nem pré-visualização; serve para o utilizador escolher os contactos a responder).
   - **Remover** da lista (… > **Remove permission**): `Mail.ReadWrite`, `User.Read`, `offline_access`.
   - Para adicionar: **Add a permission > Microsoft Graph > Delegated permissions > Mail.ReadBasic**.
2. Na mesma página: **Grant admin consent for <tenant>**. É necessário consentimento de administrador porque a app terá *Assignment required* (ponto 4). Nota da Microsoft: conceder consentimento ao nível do tenant pode revogar permissões já concedidas ao nível do tenant para esta app.
3. **Entra ID > Enterprise apps > All applications >** "DIMSCALE AI Email Assistant" **> Security > Permissions >** separador **Admin consent**: confirmar que só restam `access_as_user`, `openid`, `profile`, `Mail.ReadBasic`. Se ainda aparecer `Mail.ReadWrite`, `User.Read` ou `offline_access`: **… > Revoke permission**. (A Microsoft indica que `offline_access` fica implicitamente concedido sempre que há uma permissão delegada; a remoção é só para a lista ficar limpa.) **Não** usar os scripts de PowerShell "remove all" dessa página: apagam também `access_as_user`.
4. **Enterprise apps >** "DIMSCALE AI Email Assistant" **> Properties > Assignment required? = Yes > Save**. Depois **Users and groups > Add user/group >** o grupo dos utilizadores DIMSCALE **> Assign**. Requer Microsoft Entra ID **P1 ou P2**; grupos aninhados não contam.
   - **[devolver]** nome do grupo, e se o tenant tem P1/P2.
5. **[devolver]** confirmação dos pontos 1 a 4 (captura do separador Admin consent sem dados pessoais, ou a lista em texto).

Fontes: https://learn.microsoft.com/entra/identity/enterprise-apps/grant-admin-consent, https://learn.microsoft.com/entra/identity/enterprise-apps/manage-application-permissions, https://learn.microsoft.com/entra/identity/enterprise-apps/assign-user-or-group-access-portal, https://learn.microsoft.com/entra/identity/enterprise-apps/application-properties, https://learn.microsoft.com/entra/identity-platform/scopes-oidc, https://learn.microsoft.com/graph/permissions-reference

### B. Grupo para os rascunhos automáticos (Exchange)
6. Um grupo **com e-mail** (grupo Microsoft 365, grupo de segurança com e-mail ou lista de distribuição), por exemplo `DIMSCALE Mail AI Auto-Drafts`, com os utilizadores piloto como **membros diretos** (membros de grupos aninhados não contam no filtro `MemberOfGroup`). Pode ser o mesmo grupo do ponto 4 se já tiver e-mail.
   - **[devolver]** nome do grupo e utilizadores piloto.

### C. Nova app "DIMSCALE Mail AI Drafts" (só servidor, sem utilizador)
7. **Entra ID > App registrations > New registration**: nome `DIMSCALE Mail AI Drafts`; **Accounts in this organizational directory only**; **sem Redirect URI**.
8. **API permissions**: remover o `User.Read` por omissão. **Não adicionar nem consentir nenhuma permissão de aplicação do Microsoft Graph** (nem `Mail.Read`, nem `Mail.ReadWrite`, nem `Mail.Send`). O acesso ao correio é dado só no Exchange (secção D), limitado ao grupo; um consentimento Entra seria para todas as caixas do tenant e anularia o limite.
9. **Certificates & secrets > Certificates > Upload certificate**: o ficheiro `drafts.cer` que enviaremos (só a chave pública; a chave privada nunca sai do nosso lado). **Não** criar client secret.
10. **[devolver]** Application (client) ID; thumbprint e data de expiração do certificado; e, em **Enterprise apps >** `DIMSCALE Mail AI Drafts`, o **Object ID** do service principal (não é o Object ID da página App registrations).

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

### E. Paragem de emergência (para referência)
- Do nosso lado (imediato): desligamos a funcionalidade no backend; cada utilizador também a pode desligar no separador "Responder".
- Do vosso lado: apagar o certificado da app `DIMSCALE Mail AI Drafts` (deixam de ser emitidos tokens; os existentes expiram em cerca de 1 hora) e/ou `Remove-ManagementRoleAssignment` nas duas atribuições (30 min a 2 h).

### F. Mais tarde (não é preciso agora)
14. A implementação centralizada do add-in (**Microsoft 365 admin center > Settings > Integrated apps > Add-ins > Deploy Add-in**) será pedida quando o manifesto final estiver publicado: `https://dimscale.github.io/DIMSCALE-EMAIL-AI/manifest.xml`, gerado pelo CI depois do novo design (F-012) estar em produção. Traz a entrada "DIMSCALE MAIL" no friso e o painel fixável. Depois disso, alterações de código chegam aos utilizadores sem intervenção; só uma alteração do manifesto pede novo upload (**Update**).

Fonte: https://learn.microsoft.com/microsoft-365/admin/manage/manage-deployment-of-add-ins

### Perguntas
- O tenant aceita certificados para apps (validade e política de rotação)?
- Há Conditional Access para workload identities, ou outra política que bloqueie um token só de aplicação?
- Contacto para emergências?

Obrigado.
