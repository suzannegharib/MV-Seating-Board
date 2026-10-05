# MV Seating Board — hosted entirely in SharePoint

This replaces the earlier externally hosted Microsoft 365 approach. Both the web part assets and shared data live in Microsoft 365. Users open a modern page on https://darmvegypt.sharepoint.com/sites/DMGHRKnowledgeBase and use their existing SharePoint sign-in.

## Status

This is a source handoff for a SharePoint developer, not an installable .sppkg package. The original board JavaScript has passed syntax checks and the starting records have been validated for the list storage size limit. The SPFx wrapper has not been compiled or tested in a tenant. The available local runtime is Node 24, whereas the verified SPFx compatibility table specifies Node 22 for SPFx 1.23.2. No deployment or changes to your SharePoint site have been made.

## For your IT helpdesk

Please route this to the administrator/developer who manages our SharePoint App Catalog. We need a seating board hosted directly on the DMGHRKnowledgeBase site. The attached source adapts the existing HTML as a SharePoint Framework web part, uses the signed-in user's SharePoint permissions, and saves shared state in a dedicated SharePoint list. Please build, review, test and deploy it, then add it to a modern page. No separate web host or Azure Static Web Apps is required for this approach.

## Developer packaging

1. In an approved development environment with Node 22 LTS, create a standard SPFx 1.23.2 project using Microsoft's SharePoint generator. Choose a client-side web part, no JavaScript framework, component name `MvSeatingBoard`, and solution name `mv-seating-board`.
2. Replace the generated files under `src/webparts/mvSeatingBoard/` with the three source files in this handoff. Remove unused generated style/localization imports if your generator retains them elsewhere. Ensure `config/config.json` points the bundle to `./lib/webparts/mvSeatingBoard/MvSeatingBoardWebPart.js` and the supplied manifest. Retain the generator's package, TypeScript, Heft rig and dependency configuration.
3. In `config/package-solution.json`, use `includeClientSideAssets: true`. Keep deployment scoped to the required site rather than automatically enabling every site. No `webApiPermissionRequests` or custom Entra application registration is needed by this implementation.
4. Build and package using the generated Heft toolchain:

   ```text
   heft build --production
   heft package-solution --production
   ```

5. Review the resulting `.sppkg`, upload it to the tenant App Catalog or an enabled site collection App Catalog, approve the deployment, and install it on DMGHRKnowledgeBase as required by the chosen deployment scope. SharePoint hosts included client-side assets automatically.
6. Create a modern site page, add **MV Seating Board**, and publish the page. Use a wide layout for the floor maps.

## One-time data setup

A site owner with Manage Lists permission opens the new page and uploads `Starting-Data.json` using the setup control inside the web part. The web part creates **MV Seating Board State**, enables version history, creates a plain multiline `Payload` field and a unique `Title` field, and imports nine documents covering six floors and 570 seats. The employee data is kept in this separate initialization file; do not bundle it in the client-side assets. If the original offline board contains newer edits, initialize using its current full JSON backup instead.

Grant board editors item edit permission and viewers read permission on this dedicated list. Site owners should review list permissions and retention. The source always uses the site where the web part is placed, so install it on DMGHRKnowledgeBase. On a new site it would have separate data.

## Validation before rollout

- Compile and address any project-specific TypeScript/lint configuration errors.
- Test the srcdoc iframe, same-origin messaging, downloads and print popups under the tenant's CSP and browser policies.
- Confirm the initial setup control works for a site owner and cannot initialize data as an ordinary reader/editor.
- Open the page as two different users. Edit and save an entry as user A; confirm user B sees the saved change within about 15 seconds while not typing or in a dialog. Reload to confirm persistence.
- Simultaneously edit the same floor. Verify the stale list item update receives a conflict and does not overwrite newer data. State is saved per floor/document, so conflicts require redoing the affected edit rather than merging individual seats.
- Verify response ETags, rapid edits, network interruptions, missing list items, and read-only users. Writes must not show as saved after a failed request. If SharePoint omits ETag on update, this version retains the old version to force a safe refresh/conflict on the next write; improve that behavior if needed after tenant testing.
- Check Excel/PDF export, imports, full backups and undo with disposable test records.

## Practical limits

Updates are polled every 15 seconds. Existing UI refresh deferral during typing/dialogs remains. The storage adapter rejects document payloads over 60,000 characters; current starting documents are at most 35,029 characters. If the board outgrows this limit, migrate to one list item per seat/request rather than raising the limit blindly. Activity logging, cross-floor moves and bulk operations span multiple state items and are not atomic transactions. Treat this as a prototype pending tenant acceptance testing.

The iframe isolates the original board's styles and script libraries from the SharePoint page. This is trusted app code, not a sandbox for arbitrary user HTML. Do not add arbitrary HTML execution features. Backups and original floor data may contain employee information and should follow company handling rules.

## Microsoft documentation

- Hosting packaged client assets: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/enterprise-guidance
- App Catalog deployment: https://learn.microsoft.com/en-us/sharepoint/dev/general-development/site-collection-app-catalog
- Toolchain compatibility: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/compatibility
- Conditional list updates: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints
- Heft packaging: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/customize-heft-toolchain-overview
