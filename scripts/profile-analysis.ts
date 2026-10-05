// Concise CPU/heap conclusions. Raw captures remain disposable, even on failure.
export function analyseProfiles(cpu:any,heap:any,intervalMicros:number){
 const frameName=(value:string)=>value.length>180?value.slice(0,180)+'…':value;
 const nodes=new Map<number,any>((cpu.nodes??[]).map((n:any)=>[n.id,n]));
 const parents=new Map<number,number>();
 for(const n of cpu.nodes??[])for(const child of n.children??[])parents.set(child,n.id);
 const cumulative=new Map<string,number>();
 for(let i=0;i<(cpu.samples??[]).length;i++){
  let id:number|undefined=cpu.samples[i];const seen=new Set<string>(),visited=new Set<number>();
  while(id!==undefined&&!visited.has(id)){
   visited.add(id);const n=nodes.get(id);if(!n)break;
   const frame=n.callFrame,key=frameName(`${frame.functionName||'(anonymous)'} ${frame.url}:${frame.lineNumber+1}`);
   if(!seen.has(key)){cumulative.set(key,(cumulative.get(key)??0)+(cpu.timeDeltas?.[i]??intervalMicros));seen.add(key)}
   id=parents.get(id);
  }
 }
 const sums=new Map<string,{bytes:number;objects:number}>();let heapKind:string,totalHeapBytes=0;
 if(heap.head){
  heapKind='sampled-allocation-tree';
  function visit(n:any):number{const total=(n.selfSize??0)+(n.children??[]).reduce((sum:number,child:any)=>sum+visit(child),0);const f=n.callFrame,key=frameName(`${f?.functionName||'(anonymous)'} ${f?.url}:${(f?.lineNumber??-1)+1}`);const old=sums.get(key)??{bytes:0,objects:0};old.bytes+=total;sums.set(key,old);return total}
  totalHeapBytes=visit(heap.head);
 }else if(heap.snapshot?.meta?.node_fields){
  heapKind='end-of-process-live-heap-snapshot';
  const fields=heap.snapshot.meta.node_fields,step=fields.length,size=fields.indexOf('self_size'),name=fields.indexOf('name');
  if(!step||size<0||name<0||!Array.isArray(heap.nodes)||!Array.isArray(heap.strings))throw Error('Invalid heap snapshot');
  for(let i=0;i<heap.nodes.length;i+=step){const key=frameName(heap.strings[heap.nodes[i+name]]??'(unknown)'),row=sums.get(key)??{bytes:0,objects:0};row.bytes+=heap.nodes[i+size];row.objects++;totalHeapBytes+=heap.nodes[i+size];sums.set(key,row)}
 }else throw Error('Unsupported Bun heap profile shape');
 const samples=cpu.samples?.length??0;
 return {
  status:samples?'analysed-needs-engineering-review':'empty-cpu-samples',cpuSamples:samples,cpuIntervalMicros:intervalMicros,
  cumulativeCpuTop:[...cumulative].sort((a,b)=>b[1]-a[1]).slice(0,12).map(([frame,microseconds])=>({frame,microseconds})),
  heapKind,totalHeapBytes,heapTop:[...sums].sort((a,b)=>b[1].bytes-a[1].bytes).slice(0,12).map(([frame,value])=>({frame,...value})),
  limits:'Live heap is retained memory, not allocation churn or Go alloc_space/alloc_objects; parent capture excludes external subprocesses. Review hotspots and tune before pre-release acceptance; no performance comparison is inferred.'
 };
}
