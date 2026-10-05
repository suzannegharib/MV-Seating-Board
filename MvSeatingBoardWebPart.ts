import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-core-library';
import { boardDocument } from './boardDocument';

const KEYS = ['floors/F1','floors/F2','floors/F3','floors/F4','floors/F5','floors/SP5','plan/main','changes/main','settings/main'];
const LIST = 'MV Seating Board State';
interface StoredDocument { id: number; etag: string; data: object; }
interface DocumentMap { [key: string]: StoredDocument; }

export default class MvSeatingBoardWebPart extends BaseClientSideWebPart<{}> {
  private frame: HTMLIFrameElement | undefined;
  private listener: ((event: MessageEvent) => void) | undefined;
  private channel: string = '';
  private get api(): string { return this.context.pageContext.web.absoluteUrl + '/_api/web'; }
  private get listApi(): string { return this.api + "/lists/getbytitle('" + LIST + "')"; }

  public render(): void {
    this.cleanup();
    this.channel = this.instanceId + ':' + Math.random().toString(36).slice(2);
    this.frame = document.createElement('iframe');
    this.frame.title = 'MV Seating Board';
    this.frame.name = this.channel;
    this.frame.setAttribute('sandbox','allow-scripts allow-same-origin allow-downloads allow-popups allow-popups-to-escape-sandbox allow-modals');
    this.frame.style.cssText = 'width:100%;height:1000px;border:0;display:block';
    this.listener = (event: MessageEvent): void => {
      if (!this.frame || event.source !== this.frame.contentWindow || event.origin !== window.location.origin || !event.data || event.data.channel !== this.channel) return;
      const { id, operation, args } = event.data;
      const frame = this.frame;
      this.dispatch(operation,args || {}).then(result => {
        if (frame === this.frame && frame.contentWindow) frame.contentWindow.postMessage({channel:this.channel,id,result},window.location.origin);
      }).catch(error => {
        if (frame === this.frame && frame.contentWindow) frame.contentWindow.postMessage({channel:this.channel,id,error:{code:error.code || 'unavailable',message:error.message || 'SharePoint operation failed'}},window.location.origin);
      });
    };
    window.addEventListener('message',this.listener);
    this.frame.srcdoc = boardDocument;
    this.domElement.replaceChildren(this.frame);
  }

  private async read(): Promise<{ missing: boolean; docs: DocumentMap; user: string }> {
    const response = await this.context.spHttpClient.get(this.listApi + '/items?$select=Id,Title,Payload&$top=100',SPHttpClient.configurations.v1);
    if (response.status === 404) return {missing:true,docs:{},user:this.context.pageContext.user.displayName};
    if (response.status === 400) {
      const field = await this.context.spHttpClient.get(this.listApi + "/fields/getbyinternalnameortitle('Payload')",SPHttpClient.configurations.v1);
      if (field.status === 404) return {missing:true,docs:{},user:this.context.pageContext.user.displayName};
    }
    await this.check(response);
    const json = await response.json();
    const docs: DocumentMap = {};
    for (const item of json.value) {
      if (KEYS.indexOf(item.Title) < 0) continue;
      const etag = item['@odata.etag'] || item['odata.etag'];
      if (!etag) throw new Error('SharePoint did not return a document version.');
      docs[item.Title] = {id:item.Id,etag,data:JSON.parse(item.Payload)};
    }
    return {missing:KEYS.some(key => !docs[key]),docs,user:this.context.pageContext.user.displayName};
  }

  private async check(response: SPHttpClientResponse): Promise<void> {
    if (response.ok) return;
    const code = response.status === 412 || response.status === 409 ? 'conflict' : response.status === 401 || response.status === 403 ? 'invalid_argument' : 'unavailable';
    throw Object.assign(new Error('SharePoint returned ' + response.status),{code});
  }

  private post(url: string, body: object, extra: { [key: string]: string } = {}): Promise<SPHttpClientResponse> {
    return this.context.spHttpClient.post(url,SPHttpClient.configurations.v1,{
      headers:{'Accept':'application/json;odata=nometadata','Content-Type':'application/json;odata=nometadata',...extra},body:JSON.stringify(body)
    });
  }

  private async dispatch(operation: string, args: any): Promise<any> {
    if (operation === 'getAll') return this.read();
    if (operation === 'put') {
      if (KEYS.indexOf(args.key) < 0 || !args.body || typeof args.body !== 'object' || typeof args.etag !== 'string') throw new Error('Invalid board update');
      const payload = JSON.stringify(args.body);
      if (payload.length > 60000) throw Object.assign(new Error('This board document exceeds the storage limit. Contact the administrator.'),{code:'invalid_argument'});
      const current = await this.read();
      const doc = current.docs[args.key];
      if (!doc || doc.etag !== args.etag) throw Object.assign(new Error('Someone else has changed this document.'),{code:'conflict'});
      const response = await this.post(this.listApi + '/items(' + doc.id + ')',{Payload:payload},{'IF-MATCH':args.etag,'X-HTTP-Method':'MERGE'});
      await this.check(response);
      // Return the version belonging to this write, not a later user's version.
      const etag = response.headers.get('ETag');
      if (etag) return {etag};
      // If the service omits the response version, the old version forces a safe refresh/conflict next time.
      return {etag:args.etag};
    }
    if (operation === 'initialize') {
      if (!this.context.pageContext.web.permissions.hasPermission(SPPermission.manageLists)) throw new Error('A site owner with Manage Lists permission must initialize the board.');
      const seed = args.seed;
      if (!seed || seed.app !== 'MV Seating Board' || KEYS.some(key => !seed[key] || JSON.stringify(seed[key]).length > 60000)) throw new Error('The backup is incomplete or too large.');
      const current = await this.read();
      if (!current.missing) throw new Error('The board is already initialized.');
      const probe = await this.context.spHttpClient.get(this.listApi,SPHttpClient.configurations.v1);
      if (probe.status === 404) {
        await this.check(await this.post(this.api + '/lists',{Title:LIST,BaseTemplate:100,Description:'Shared MV seating board state',EnableVersioning:true}));
      } else await this.check(probe);
      const field = await this.context.spHttpClient.get(this.listApi + "/fields/getbyinternalnameortitle('Payload')",SPHttpClient.configurations.v1);
      if (field.status === 404) await this.check(await this.post(this.listApi + '/fields/createfieldasxml',{parameters:{SchemaXml:'<Field Type="Note" DisplayName="Payload" Name="Payload" StaticName="Payload" NumLines="6" RichText="FALSE" />'}}));
      else await this.check(field);
      await this.check(await this.post(this.listApi + "/fields/getbyinternalnameortitle('Title')",{Indexed:true,EnforceUniqueValues:true},{'IF-MATCH':'*','X-HTTP-Method':'MERGE'}));
      for (const key of KEYS) {
        if (current.docs[key]) continue; // Resume setup without replacing previously stored data.
        await this.check(await this.post(this.listApi + '/items',{Title:key,Payload:JSON.stringify(seed[key])}));
      }
      return {ok:true};
    }
    throw new Error('Unknown board operation');
  }

  private cleanup(): void {
    if (this.listener) window.removeEventListener('message',this.listener);
    if (this.frame) this.frame.remove();
    this.frame = undefined; this.listener = undefined;
  }
  protected onDispose(): void { this.cleanup(); }
}
