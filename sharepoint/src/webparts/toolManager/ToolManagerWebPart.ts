import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { APP_HTML, APP_SHA256 } from '../../generated/appPayload';
import { hostedDocument, safeFrameHeight, frameHasPendingWork } from '../../hostFrame';

export interface IToolManagerWebPartProps {
  fileUrl: string;
}

/** Hosts the fixed application on the authenticated SharePoint origin.
 * No Graph permission requests, bearer tokens or passwords are supplied.
 */
export default class ToolManagerWebPart extends BaseClientSideWebPart<IToolManagerWebPartProps> {
  private _frame: HTMLIFrameElement | undefined;
  private _status: HTMLParagraphElement | undefined;
  private _readyTimer: number | undefined;
  private _siteOrigin: string = '';
  private _listening: boolean = false;
  private _frameKey: string = '';
  private _renderedFileUrl: string = '';

  private readonly _onMessage = (event: MessageEvent): void => {
    if (!this._frame || event.source !== this._frame.contentWindow || event.origin !== this._siteOrigin || !event.data || typeof event.data !== 'object') return;
    const message = event.data as { type?: string; height?: number };
    if (message.type === 'TOOL_SHAREPOINT_READY') {
      if (this._readyTimer !== undefined) window.clearTimeout(this._readyTimer);
      this._readyTimer = undefined;
      if (this._status) { this._status.textContent = ''; this._status.hidden = true; }
    }
    const height = safeFrameHeight(message.height);
    if (message.type === 'TOOL_SHAREPOINT_HEIGHT' && height !== undefined) this._frame.style.height = height + 'px';
  };

  public render(): void {
    const fileUrl = typeof this.properties.fileUrl === 'string' ? this.properties.fileUrl.trim() : '';
    const key = JSON.stringify([this.context.pageContext.web.absoluteUrl, fileUrl]);
    if (this._frame && this._frameKey === key) return;
    if (this._frame && frameHasPendingWork(this._frame)) {
      this.properties.fileUrl = this._renderedFileUrl;
      if (this._status) {
        this._status.hidden = false;
        this._status.textContent = '사진 작업을 완료한 뒤 저장 대상을 변경해 주세요. 현재 사진과 연결된 엑셀을 유지했습니다.';
      }
      return;
    }
    this._removeFrame();
    this.domElement.replaceChildren();
    const status = document.createElement('p');
    status.setAttribute('role', 'status');
    status.style.cssText = 'padding:12px;border-radius:8px;background:#eef6f3;color:#123e37;font:14px/1.7 sans-serif';
    this.domElement.appendChild(status);
    this._status = status;
    let siteUrl: URL;
    try {
      siteUrl = new URL(this.context.pageContext.web.absoluteUrl);
      if (siteUrl.protocol !== 'https:' || siteUrl.origin !== window.location.origin || siteUrl.username || siteUrl.password) throw new Error('origin');
    } catch {
      status.textContent = '회사 SharePoint 사이트의 페이지에서 이 앱을 열어 주세요. 다른 사이트의 로그인 권한을 대신 사용하지 않습니다.';
      return;
    }
    this._siteOrigin = siteUrl.origin;
    if (!fileUrl) {
      status.textContent = '페이지 편집에서 이 웹 파트의 연필 버튼을 누르고, 관리 엑셀의 SharePoint 링크를 한 번 입력한 뒤 페이지를 게시하세요. 직원들은 게시된 페이지에서 사진만 올리면 됩니다.';
      return;
    }
    const context = { siteUrl: siteUrl.href.replace(/\/$/, ''), fileUrl, applicationSha256: APP_SHA256 };
    const frame = document.createElement('iframe');
    frame.title = '툴 검교정 매니저';
    frame.setAttribute('sandbox', 'allow-scripts allow-same-origin allow-forms allow-downloads allow-modals');
    frame.setAttribute('allow', 'camera; clipboard-write; web-share');
    frame.style.cssText = 'display:block;width:100%;height:1000px;border:0;min-height:650px;background:#f3f6f5';
    this._frame = frame;
    this._frameKey = key;
    this._renderedFileUrl = fileUrl;
    if (!this._listening) { window.addEventListener('message', this._onMessage); this._listening = true; }
    status.textContent = '툴 사진 인식 앱을 준비하고 있습니다…';
    this._readyTimer = window.setTimeout((): void => {
      if (this._status) {
        this._status.hidden = false;
        this._status.textContent = '앱의 시작 응답을 받지 못했습니다. 이 SharePoint 페이지에서 스크립트 실행이 제한됐는지 확인하고 페이지를 다시 여세요. 엑셀 저장은 시작되지 않았습니다.';
      }
    }, 45000);
    try {
      frame.srcdoc = hostedDocument(APP_HTML, context);
      this.domElement.appendChild(frame);
    } catch {
      if (this._readyTimer !== undefined) window.clearTimeout(this._readyTimer);
      this._readyTimer = undefined;
      status.textContent = '이 SharePoint 페이지에서 앱 화면을 실행하지 못했습니다. 페이지의 스크립트 실행 설정을 확인해 주세요. 엑셀은 변경되지 않았습니다.';
    }
  }

  private _removeFrame(): void {
    if (this._readyTimer !== undefined) window.clearTimeout(this._readyTimer);
    this._readyTimer = undefined;
    this._frame = undefined;
    this._frameKey = '';
    this._renderedFileUrl = '';
  }

  protected onDispose(): void {
    this._removeFrame();
    if (this._listening) window.removeEventListener('message', this._onMessage);
    this._listening = false;
    this.domElement.replaceChildren();
    super.onDispose();
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected get disableReactivePropertyChanges(): boolean { return true; }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: '담당자가 한 번 지정한 공유 엑셀을 모든 직원이 사용합니다.' },
        groups: [{
          groupName: '관리 엑셀',
          groupFields: [PropertyPaneTextField('fileUrl', {
            label: 'SharePoint 엑셀 링크 또는 사이트 안의 파일 경로',
            description: '이 사이트의 .xlsx 파일을 지정하세요. 회사 계정·앱 ID·비밀번호를 입력하지 않습니다.',
            multiline: true
          })]
        }]
      }]
    };
  }
}
