# Azure app registration — the one step you do by hand

This is the only manual part. You are producing a single GUID: an application (client) ID. Once, then never again.

When you have it, hand [SETUP-FOR-CLAUDE.md](SETUP-FOR-CLAUDE.md) to Claude Code and the rest is automatic.

It is free. No Azure subscription is required.

---

## Why you do this yourself

It involves browser sign-in, a consent screen, and picking a tenant. That is not work to delegate to an agent, and it should not be — **you are issuing access to your own mailbox**. Read what you approve.

---

## Getting a tenant

**App registrations can only live inside a tenant (directory).** A personal Microsoft account — Hotmail, Outlook.com, or one you made from your own address — does not belong to a tenant by default, so opening the [Microsoft Entra admin center](https://entra.microsoft.com/) greets you with:

```
The user account you are signing in with does not exist in tenant 'Microsoft Services',
so it cannot access the application '...' in that tenant.
```

**This does not mean personal accounts are unsupported.** Sign-in and mailbox access both work fine with a personal account. The only thing missing is *somewhere to mint a GUID*.

**If you have a work or school account**, signing in with that is the shortest path. **The tenant you register in does not have to match the mailbox you read.** As long as step 3 below includes personal accounts, an ID issued in your employer's tenant will happily sign you into your personal Hotmail.

**If you do not**, create a [free Azure account](https://azure.microsoft.com/free/). Signing up creates a directory alongside it and gets you into the admin center.

- Identity verification asks for a credit card, but nothing moves to pay-as-you-go unless you explicitly upgrade
- App registration and Microsoft Graph calls are not billable in the first place
- Once registered you can ignore Azure entirely; the token is tied to your personal account

> If your browser is signed into several Microsoft accounts, the wrong one may be picked and you will
> see the same error. An InPrivate window rules that out.

---

## Register

1. [Microsoft Entra admin center](https://entra.microsoft.com/) → **App registrations** → **New registration**

2. Any name (e.g. `outlook-mcp`)

3. **Supported account types**:
   - For Hotmail/Outlook.com → **Accounts in any organizational directory and personal Microsoft accounts**

4. Leave the redirect URI **empty** on this screen — you add it in step 7

5. From the overview page, copy the **Application (client) ID**
   → this is `OUTLOOK_CLIENT_ID` in `.env`. Keep it handy

6. Left menu **Authentication** → **Advanced settings** at the bottom →
   set **Allow public client flows** to **Yes** and save
   (required for device code flow; without it `login.py` fails)

7. Still under **Authentication**, choose **Add a platform** →
   **Mobile and desktop applications** → tick both of these and save
   ```
   https://login.microsoftonline.com/common/oauth2/nativeclient
   https://login.live.com/oauth20_desktop.srf     ← required for personal accounts
   ```

8. Left menu **API permissions** → **Add a permission** →
   **Microsoft Graph** → **Delegated permissions** → add **both**

   | Permission | Used for |
   |---|---|
   | `Mail.ReadWrite` | searching, reading, moving, marking read; creating/renaming/moving/deleting folders |
   | `MailboxSettings.ReadWrite` | inbox rules (`list_rules` / `create_rule` / `delete_rule`) |

   Do **not** add `Mail.Send`. This server has no send capability and never asks for one.

**Do not create a client secret.** The server runs as a public client and does not need one.

---

## The two traps

Both cost real time and neither is obvious. If you only read one section of this document, read this one.

### Skipping step 7 makes personal-account login fail every time

Right after you enter the code in the browser:

```
invalid_request: The provided request must include a 'redirect_uri' input parameter.
```

"Allow public client flows" in step 6 is **necessary but not sufficient**. It only sets `allowPublicClient`. The consumer authentication server (`login.live.com`) separately insists that a redirect URI actually exists on the registration. Work and school accounts pass without one, which is exactly why this is easy to miss.

Device code flow never navigates to that URI. It only requires that it is registered.

### Forgetting `MailboxSettings.ReadWrite` in step 8 breaks only the rule tools

You get a confusing partial failure: search and move work perfectly, while `list_rules` and `create_rule` return 403. `check_config` names it directly:

```
OK: Inbox 3,412 messages (87 unread) reachable.
Inbox rules: unavailable — MailboxSettings.ReadWrite has not been granted.
```

**Adding the permission in Azure is not enough by itself.** An already-issued token still carries the old set, so you have to consent again:

```
python login.py
```

`login.py` compares the scopes it needs against the cached token and starts the consent flow automatically when something is missing. Earlier versions reported "already signed in" and exited without doing anything — which made this failure very hard to diagnose.

---

## When the labels do not match

Azure changes its wording often. Searching for the English strings usually finds them:

- **Allow public client flows**
- **Delegated permissions**
- **Add a platform**

If the Authentication screen is the newer **Authentication (Preview)** UI and you cannot find "Add a platform", the link at the top — *"To switch to the old experience, please click here."* — takes you back.

---

## Next

With the client ID in hand, clone the repository, start Claude Code in it, and say:

```
Read docs/SETUP-FOR-CLAUDE.md and set this up
```

The virtual environment, dependencies, `.env`, MCP registration and connectivity check all happen from there. You will be asked to pause once, to sign in.

> That runbook is written in Japanese because its reader is an agent, and Claude follows it without
> trouble. The English [README](../README.md) covers the manual equivalent if you prefer to do it yourself.
