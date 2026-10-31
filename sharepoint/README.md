# Tool Manager SharePoint web part

SharePoint Framework 1.24.0 package for the tool-calibration manager. The web part runs the complete offline photo-recognition program inside the company SharePoint page, using that page's existing signed-in session and same-origin SharePoint REST. It requests no Microsoft Graph permissions and does not require users to enter a Client ID, Tenant ID, password, or client secret.

The page author chooses the workbook once in the web-part property pane. After the page is published, engineers open that page on a phone or PC and upload tool-label/SMAX photos. The application still needs the current user's normal file read/edit permissions. The package must be added through the site's or tenant's SharePoint app catalog; uploading it to an ordinary document library does not install a web part. Deployment availability and company SharePoint CSP/worker restrictions are evaluated by the actual site when tested.

Installation instructions: [SharePoint setup](../docs/SharePoint-setup.md).

## Rebuild

Use a supported Node.js version (Node 24 is used here). From the repository root, build the self-contained app and point this project's embed step to it:

```bash
python web/build.py --output /absolute/output/Tool-Manager-SharePoint.html
cd sharepoint
npm ci
TOOL_MANAGER_APP=/absolute/output/Tool-Manager-SharePoint.html npm run build
```

The environment variable is optional in the cloud workspace; the default input is `/workspace/outputs/Tool-Manager-SharePoint.html`. Set it explicitly for another machine. `npm run build` embeds the exact app bytes, runs host lifecycle checks, compiles and lints TypeScript, builds a production bundle, packages the official `.sppkg`, then verifies its contents and CRCs.

Output: `sharepoint/solution/tool-manager-sharepoint.sppkg` relative to this directory. `app-build-info.json` records the embedded application SHA256. The package includes the app, WASM, and Korean/English OCR models; it does not load the program from GitHub at runtime. Framework host libraries are supplied by SharePoint itself. Original workbooks, images, company links, and authentication material are not included in the public package.

`config/package-solution.json` has `includeClientSideAssets: true`, `skipFeatureDeployment: false`, and `isDomainIsolated: false`. There are no `webApiPermissionRequests`. The package's internal SPFx component/solution GUIDs identify this deployable web part; they are not Microsoft Entra Client/Tenant IDs and require no user input.

## Shared-file updates

The SharePoint adapter does not rely on a file-stream `If-Match` header alone for concurrency. It refuses an existing checkout, acquires a new SharePoint `CheckOut`, downloads and compares the locked file with the expected original SHA, writes and verifies the draft, then performs `CheckIn` and verifies the final file. It does not take over a pre-existing checkout, publish someone else's draft, or retry an uncertain upload automatically. A save that cannot be confirmed is reported for review with its local backup retained. The exact library rules and concurrent-write behavior are still subject to validation on the actual company SharePoint tenant.

Microsoft's REST file guide documents `CheckOut`/`CheckIn` as the way to keep others from changing a file during an update: [Working with folders and files with REST](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest).

## Verification scope

`npm test` exercises the actual TypeScript host helpers and web-part render lifecycle with controlled SPFx/DOM fixtures: context escaping/freezing, source/origin validation, resize bounds, same-property renders, pending-photo/work target guards, startup failure reporting, and disposal. Browser OCR/workbook tests live in the repository's `tests/` directory. The build and these checks do not prove deployment or REST behavior on a company tenant; the actual company page, file permissions, mobile browser, CSP and concurrent-edit behavior must be checked there.
