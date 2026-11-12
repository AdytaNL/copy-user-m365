# Copy a Microsoft 365 user

**Create a new employee from an existing colleague, with the same access.** A PowerShell
tool with a small Windows window: pick a colleague as the template, enter the new name and
address, and start. The tool creates the user and copies group memberships, licenses and
shared mailbox permissions, and can transfer the licenses and Teams phone number of a
colleague who left.

Background: [A new colleague, set up like the last](https://adyta.nl/en/engineering/a-new-colleague-set-up-like-the-last/).

## Requirements

- Windows (the window uses WinForms) with Windows PowerShell 5.1 or PowerShell 7.
- Modules:

  ```powershell
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  Install-Module Microsoft.Graph.Users -Scope CurrentUser
  Install-Module Microsoft.Graph.Users.Actions -Scope CurrentUser
  Install-Module Microsoft.Graph.Groups -Scope CurrentUser
  Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
  Install-Module ExchangeOnlineManagement -MinimumVersion 3.0.0 -Scope CurrentUser
  Install-Module MicrosoftTeams -Scope CurrentUser   # only for the phone number transfer; use a current version
  ```

- An admin account with the roles below. The tool signs in interactively (delegated).
- A test tenant for your first runs.

## Quick start

```powershell
.\CopyUser-M365.ps1
```

1. Search for the colleague to copy (display name, UPN, first name or surname, starting with
   the text you type; up to 25 results) and select them. The Graph sign-in starts at the first
   search.
2. Click **Fill from source**, adjust first name, last name, display name and the new UPN
   (suggested as `first.last@<source domain>`), and set or generate a password.
3. Choose the options and click **Start copy**. Every step is shown in the log pane and written
   to a log file.

Sign-in: Microsoft Graph opens the Windows sign-in window; Exchange Online signs in through the
browser; Microsoft Teams (only for the phone number transfer) uses a device code, shown in the
PowerShell window behind the tool. Start the tool from a visible PowerShell window and keep it
open during a copy.

## Options

| Option | What it does |
|---|---|
| **Copy mailbox permissions** | FullAccess and SendAs on shared mailboxes |
| **Copy licenses** | Directly assigned licenses, including disabled service plans |
| **Transfer licenses from source** | Moves the direct licenses instead of using new seats; also ticks **Copy licenses** |
| **Transfer phone number from source (Teams)** | Moves the source's Teams number and turns on Enterprise Voice |

Group memberships are always copied.

## Permissions

Graph scopes (requested at first sign-in):

| Scope | For |
|---|---|
| `User.Read.All` | Find and read the source user |
| `User.ReadWrite.All` | Create the new user |
| `Group.Read.All` | Read the source user's group memberships |
| `GroupMember.ReadWrite.All` | Add the new user to groups |
| `Directory.Read.All` | Read tenant and domain data |
| `Directory.ReadWrite.All` | Write directory objects |
| `Organization.Read.All` | Read verified domains for UPN validation |

Admin roles of the account that runs the tool (at least):

| Role | For |
|---|---|
| User Administrator | Create the user, block sign-in of the source, revoke sessions |
| License Administrator (or User Administrator) | Assign, remove and reprocess licenses |
| Groups Administrator (or User Administrator) | Add and remove group memberships |
| Exchange Administrator or Recipient Management | Convert a mailbox to shared, mailbox permissions, distribution groups |
| Teams Telephony Administrator (or Teams Administrator) | Phone number transfer, Enterprise Voice |
| Privileged Authentication Administrator | Only if the source user holds an admin role (block sign-in, revoke sessions) |

To change distribution groups that the admin does not own, the tool uses
`-BypassSecurityGroupManagerCheck`; that needs enough Exchange rights, for example
Organization Management or the Security Group Creation and Membership role.

## What is copied

### Profile

`UsageLocation`, `JobTitle`, `Department`, `OfficeLocation`, `City`, `Country` and
`CompanyName` come from the source user; empty values are skipped. Name, display name, UPN and
password are entered by hand. The mail nickname is first name plus last name, lower case,
other characters removed. The new user must change the password at first sign-in.

### Groups

- Microsoft 365 groups and security groups: through Microsoft Graph.
- Distribution groups and mail-enabled security groups: read-only in Graph, so the tool
  recognises them in the same scan and adds the user through Exchange Online
  (`Add-DistributionGroupMember`). This can start an Exchange sign-in even if mailbox
  permissions are not selected.
- Dynamic groups and dynamic distribution groups are skipped: the rule decides membership.
- Groups synced from on-premises and role-assignable groups are not changed and are reported.

### Licenses (option)

Only licenses assigned **directly** to the source are assigned directly to the new user,
including disabled service plans. Not copied: a direct license the source also gets through a
group, and a direct license in an error state. Licenses that come through a group follow from
the copied group membership, so not from dynamic, on-premises or role-assignable groups.
Before assigning, the tool checks the tenant's free seats (`Get-MgSubscribedSku`); without a
free seat that license is skipped and the run ends as `Partial`.

### License transfer (option)

For a source user who has left, in a tenant without free seats: the licenses move instead of
using a new seat. **Before anything is created**, the tool shows an overview to confirm:

- which direct licenses move from the source to the new user;
- licenses the source has both directly and through a group (the direct one is only removed
  once the source has left that group; otherwise no seat comes free);
- licenses in an error state (not moved);
- license groups without a free seat, with how each seat comes free: a fixed group (remove the
  source), a dynamic group whose rule requires `user.accountEnabled -eq true` (block sign-in),
  or by hand, with the reason;
- the source mailbox: without a license it is deleted after 30 days, so the tool offers to
  convert it to a shared mailbox first, and warns when it is over 50 GB or has an archive or
  litigation hold (a shared mailbox then needs a license too). Other holds are not checked.

If seats can be freed through groups, a second question follows: release the source (block
sign-in and revoke sessions, and/or remove it from fixed license groups).

Execution order:

1. Convert the source mailbox to shared (optional) and wait until Exchange reports it.
2. Remove the direct licenses from the source and assign them to the new user, with retries.
   If assigning fails, the licenses go back to the source and releasing the source is skipped.
3. If you chose to release the source: block sign-in, revoke sessions and/or remove it from
   fixed license groups, and wait until it is no longer a member (up to 10 minutes).
4. Only when every group through which the source gets a license is released, remove a direct
   assignment of that same license from the source.
5. Reprocess the new user's licenses and check that each group license that was freed is
   **active** on the new user. Directly transferred licenses count as done once Graph accepts
   the assignment.

### Teams phone number transfer (option)

The tool looks up the source's number before anything is created and asks for confirmation.
The transfer runs last, because the new user first needs Teams Phone, which only comes if a
license for it reaches the new user through the copied licenses or groups:

1. Wait until Teams knows the new user with Phone System, up to 30 minutes. The wait stops
   early if the new user's license is in error, or after repeated other errors from Teams.
   On a time-out nothing is changed, and the final report gives the commands to finish by hand.
2. Remove the number from the source and assign it to the new user with the same number type
   and emergency location, with retries.
3. If assigning fails, the number goes back to the source. If that is no longer possible (for
   example because the source lost Teams Phone through the license transfer), the final report
   says the number is free, with the command to assign it.
4. Turn Enterprise Voice on explicitly and check that it is on.

Calling policies and call queue memberships are not copied separately; if they come through
groups, they follow from the group copy.

### Shared mailbox permissions (option)

The tool checks all shared mailboxes for **FullAccess** and **SendAs** of the source and
gives the new user the same, with AutoMapping. Permissions on personal mailboxes are not
copied.

Exchange waits in two steps:

1. Up to 3 minutes until the new user appears in Exchange. On a time-out, both the
   distribution groups and the mailbox permissions are skipped, with the commands in the report.
2. Up to 15 minutes until Exchange reports a real `UserMailbox`. Exchange accepts
   `Add-MailboxPermission -AutoMapping:$true` before the mailbox exists, without an error, but
   does not record the AutoMapping link: the user has access, but the mailbox does not appear
   in Outlook. So the tool waits. A mailbox only appears if the new user gets an Exchange
   license; the wait stops early if that license is in error.

If the mailbox does not appear in time, the permissions are **not** set and the final report
gives the commands:

```powershell
Add-MailboxPermission -Identity <shared@contoso.example> -User <new@contoso.example> -AccessRights FullAccess -InheritanceType All -AutoMapping $true
Add-RecipientPermission -Identity <shared@contoso.example> -Trustee <new@contoso.example> -AccessRights SendAs -Confirm:$false
```

Before any Exchange or Teams action, the tool checks that the session belongs to the same
tenant as the Graph session.

## Checks before creating the user

- All required fields filled in; password of at least 12 characters with upper case, lower
  case, digit and special character.
- Valid UPN format, and the domain is verified in the tenant.
- No user with that UPN, and the address is not in use as a `proxyAddress` of another user,
  group or contact.
- With **Copy licenses**: the source has a `UsageLocation` if it has active licenses;
  otherwise the run stops before anything is created.
- A warning (not a block) when the display name or mail nickname already exists.

## Result and log

The window ends with "Copy fully completed" or "Copy completed; some items need attention",
with the concrete actions per license, group, shared mailbox, phone number or release step.
The log records the final status as `Success` or `Partial`.

Log file: `%LOCALAPPDATA%\Microsoft\IntuneManagementExtension\Logs\CopyUser-M365.log`
(`C:\ProgramData\...` when run as SYSTEM). Format:
`yyyy-MM-dd HH:mm:ss [LEVEL] [USER] message`.

Most wait times and retry counts are variables at the top of the script.

## Limitations

- One tenant at a time; no cross-tenant copy.
- Not copied: manager, calendar permissions, mailbox language, group ownership, Send on
  Behalf, enterprise app assignments, directory roles, administrative unit memberships, and
  profile fields other than those listed above.
- AutoMapping only works in classic Outlook for Windows; new Outlook and Outlook for Mac do
  not add shared mailboxes automatically.
- All Graph, Exchange and Teams calls run on the main thread: the sign-in sessions are bound
  to the PowerShell runspace you signed in from. The window stays responsive while waiting,
  with its controls disabled, and cannot be closed during a copy.

## License

MIT, see [LICENSE](LICENSE).

## Author

Built by Adyta (Lambert van der Plas) – https://adyta.nl
