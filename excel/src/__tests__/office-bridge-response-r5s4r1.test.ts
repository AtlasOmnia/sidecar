import { describe,it,expect,vi,afterEach } from 'vitest';
import { OfficeCompanionBridge } from '../officeBridge';
const L=10*1024*1024+64*1024;
const id='aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
const tid='cccccccc-cccc-4ccc-8ccc-cccccccccccc';
const session={sessionId:id,sessionKey:'k'.repeat(64),protocolVersion:'1.0',expiresInSeconds:60};
const json=(status:number,v:unknown)=>new Response(JSON.stringify(v),{status});
const flush=async()=>{for(let i=0;i<30;i++)await Promise.resolve();};
afterEach(()=>vi.useRealTimers());
function body(text:string,header?:string,opts:{throwRead?:boolean;hang?:boolean;cancelHang?:boolean;fallback?:boolean;chunk?:number}={}) {
 const bytes=new TextEncoder().encode(text); let n=0; let release!:()=>void;let reject!: (e:Error)=>void;
 const stats={reads:0,cancel:0,release:0,text:0,retained:0};
 const reader={read:async()=>{stats.reads++;if(opts.throwRead)throw new Error('reader exploded');if(opts.hang)return new Promise<any>((r,j)=>{release=()=>r({done:true});reject=j;});if(n>=bytes.length)return {done:true};const end=Math.min(bytes.length,n+(opts.chunk??65536));const value=bytes.slice(n,end);n=end;stats.retained+=value.byteLength;return {done:false,value};},cancel:async()=>{stats.cancel++;if(opts.cancelHang)await new Promise<void>(r=>release=r);},releaseLock:()=>stats.release++};
 return {response:{status:201,headers:{get:()=>header??null},body:opts.fallback?null:{getReader:()=>reader},text:async()=>{stats.text++;return text;}} as unknown as Response,stats,finish:()=>release?.(),abort:()=>reject?.(new DOMException('abort','AbortError'))};
}
function padded(n:number,multi=false){const prefix=JSON.stringify({...session,pad:''}).slice(0,-2);const suffix='"}';const budget=n-new TextEncoder().encode(prefix+suffix).length;const pad=multi?'😀'.repeat(Math.floor(budget/4))+'x'.repeat(budget%4):'x'.repeat(budget);return prefix+pad+suffix;}
function harness(target:string,response:Response){let next=response;const posts:string[]=[];const office:unknown[]=[];const bridge=new OfficeCompanionBridge({hostPrefix:'excel',documentId:async()=> 'doc',paneId:id,executeTool:async(n,a)=>{office.push([n,a]);return '{"ok":true,"data":"NORMAL"}';},fetchImpl:(async(input,init)=>{const route=String(input).split('/').pop();if(init?.method==='DELETE')return json(200,{});if(route==='session')return target==='connect'?next:json(201,session);if(route==='pending')return target==='pending'?next:json(200,{ticketId:tid,requestId:id,tool:'read_range',arguments:{sheet:'Sheet1',address:'A1'}});if(route==='result'){posts.push(String(init?.body));return target==='result'||target==='413'?next:json(200,{});}if(route==='heartbeat')return next;throw new Error('unexpected route');}) as typeof fetch});return {bridge,posts,office,set:(r:Response)=>next=r};}
async function invoke(target:string,b:OfficeCompanionBridge){if(target==='connect')return b.connect();await b.connect();return target==='heartbeat'?b.heartbeatOnce():b.pollOnce();}
describe('S4-r1 bounded reader repair',()=>{
 for(const n of [L-1,L,L+1])for(const multi of [false,true])it(`exact bytes ${n} multi=${multi}`,async()=>{const r=body(padded(n,multi),undefined,{chunk:65535});const h=harness('connect',r.response);if(n>L){await expect(h.bridge.connect()).rejects.toThrow('Companion response is too large');expect(r.stats.cancel).toBe(1);}else await expect(h.bridge.connect()).resolves.toMatchObject({sessionId:id});expect(r.stats.release).toBe(1);expect(r.stats.text).toBe(0);await h.bridge.stop();});
 it('SPEC: oversize rejects without waiting for unresolved stream cancellation',async()=>{vi.useFakeTimers();const r=body(padded(L+1),undefined,{chunk:L+1,cancelHang:true});const h=harness('connect',r.response);let settled=false;const p=h.bridge.connect().catch(()=>{settled=true;});await flush();await vi.advanceTimersByTimeAsync(30000);const didSettle=settled;r.finish();await p;await h.bridge.stop();expect(r.stats.cancel).toBe(1);expect(r.stats.release).toBe(1);expect(didSettle).toBe(true);});
 it('SPEC: accepted limit-size stream never holds more than limit bytes',async()=>{const r=body(padded(L),undefined,{chunk:65536});const h=harness('connect',r.response);const Original=globalThis.Uint8Array;let peak=0;const replacement=new Proxy(Original,{construct(target,args){if(typeof args[0]==='number')peak=Math.max(peak,r.stats.retained+args[0]);return Reflect.construct(target,args);}});vi.stubGlobal('Uint8Array',replacement);try{await h.bridge.connect();}finally{vi.unstubAllGlobals();await h.bridge.stop();}expect(peak).toBeLessThanOrEqual(L);});
});

const ID='aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';

function bridge(response:Response){return new OfficeCompanionBridge({hostPrefix:'excel',documentId:async()=> 'doc',paneId:ID,executeTool:async()=> '{}',fetchImpl:(async()=>response) as typeof fetch});}
it('real Response oversize cancel callback can keep connect and stop pending indefinitely',async()=>{
 vi.useFakeTimers();let release!:()=>void;let cancelled=0;
 const r=new Response(new ReadableStream<Uint8Array>({start(c){c.enqueue(new Uint8Array(L));c.enqueue(new Uint8Array(1));},cancel(){cancelled++;return new Promise<void>(r=>release=r);}}),{status:201});
 const b=bridge(r);let settled=false;const p=b.connect().catch(()=>{settled=true;});await flush();await vi.advanceTimersByTimeAsync(30000);const stop=b.stop();await flush();const before={settled,cancelled,locked:r.body!.locked};release();await p;await stop;expect(before.cancelled).toBe(1);expect(vi.getTimerCount()).toBe(0);vi.useRealTimers();console.log('cancel pending',JSON.stringify(before));expect(before.settled).toBe(true);
});
it('cancel rejection must not replace the oversize error',async()=>{
 const r=new Response(new ReadableStream<Uint8Array>({start(c){c.enqueue(new Uint8Array(L+1));},cancel(){throw new Error('cancel exploded');}}),{status:201});
 await expect(bridge(r).connect()).rejects.toThrow('Companion response is too large');
});
