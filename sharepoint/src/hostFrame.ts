/** Pure framing helpers shared by the SPFx host and browser integration tests. */
export interface ISharePointAppContext {
  siteUrl: string;
  fileUrl: string;
  applicationSha256: string;
}

export function hostedDocument(html: string, context: ISharePointAppContext): string {
  const encoded = JSON.stringify(context).replace(/</g, '\\u003c');
  const script = '<script data-sharepoint-shell>Object.defineProperty(window,"ToolSharePointContext",{value:Object.freeze(' + encoded + '),writable:false,configurable:false});</script>';
  if (html.indexOf('<head>') < 0) throw new Error('The app has no head element');
  return html.replace('<head>', '<head>' + script);
}

export function safeFrameHeight(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? Math.round(Math.max(650, Math.min(30000, value))) : undefined;
}

interface IFrameAppWindow extends Window {
  ToolSharePointFile?: { inspect: () => { pendingCheckout?: boolean; uncertain?: boolean; busy?: boolean } };
  ToolApp?: { state?: {
    busy?: boolean; preparing?: boolean; restoring?: boolean;
    source?: { saving?: boolean };
    photos?: { saved?: unknown }[];
  } };
}

/** Do not discard a running job or any photo whose save is unconfirmed. */
export function frameHasPendingWork(frame: HTMLIFrameElement): boolean {
  try {
    const appWindow = frame.contentWindow as IFrameAppWindow | null;
    const remote = appWindow?.ToolSharePointFile?.inspect();
    if (remote?.pendingCheckout || remote?.uncertain || remote?.busy) return true;
    const state = appWindow?.ToolApp?.state;
    return !!(state && (state.busy || state.preparing || state.restoring || state.source?.saving || state.photos?.some(photo => !photo.saved)));
  } catch {
    // A frame whose state cannot be inspected cannot be safely replaced.
    return true;
  }
}
