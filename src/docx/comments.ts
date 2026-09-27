import { OoxmlError } from '../errors.ts';
import { OpcPackage, getContentType } from '../opc/index.ts';
import { parseXml, attribute, applyEdits, type XmlElement } from '../xml/index.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const COMMENTS_REL = OFFICE + 'comments';
const EXTENDED_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const COMMENTS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const EXTENDED_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';

export type CommentInfo = {
  id: string; author?: string; initials?: string; date?: string; text: string;
  paragraphId?: string; parentId?: string; resolved: boolean;
};
export type CommentFinding = { part: string; kind: string };
export type CommentInspection = { comments: CommentInfo[]; unsupported: CommentFinding[] };
export type CommentThread = Readonly<{root:Readonly<CommentInfo>;replies:readonly Readonly<CommentInfo>[]} >;
export type CommentThreadInspection = Readonly<{threads:readonly CommentThread[];unsupported:readonly Readonly<CommentFinding>[]} >;
export type CommentThreadReceipt = {rootId:string;commentIds:string[];changed:number;changedParts:string[]};
type CommentScan = CommentInspection & { extension?: string; xml?: string; entries: Map<string, XmlElement> };
function fail(code: string, message: string): never { throw new OoxmlError('docx-comments-' + code, message); }

/** Read existing comments without creating parts or touching Document caches.
 * Malformed graph metadata refuses; unsupported comment bodies are reported.
 */
export function inspectComments(pkg: OpcPackage): CommentInspection {
  const { comments, unsupported } = scan(pkg);
  return { comments, unsupported };
}

/** Detached immutable root groups; descendants are flat, in comment-part order. */
export function inspectCommentThreads(pkg:OpcPackage):CommentThreadInspection {
  const state=scan(pkg),roots=threadRoots(state.comments);
  const groups=new Map<string,{root:Readonly<CommentInfo>;replies:Readonly<CommentInfo>[]} >();
  for(const c of state.comments)if(c.parentId===undefined)groups.set(c.id,{root:Object.freeze({...c}),replies:[]});
  for(const c of state.comments)if(c.parentId!==undefined)groups.get(roots.get(c.id)!)!.replies.push(Object.freeze({...c}));
  return Object.freeze({threads:Object.freeze([...groups.values()].map(g=>Object.freeze({root:g.root,replies:Object.freeze(g.replies)}))),unsupported:Object.freeze(state.unsupported.map(f=>Object.freeze({...f})))});
}

/** Resolve all existing members in the selected root's thread; never creates metadata. */
export function setCommentThreadResolved(pkg:OpcPackage,id:string,resolved:boolean):CommentThreadReceipt {
  if(typeof id!=='string'||!/^\d+$/.test(id)||typeof resolved!=='boolean')fail('argument','Expected a decimal comment ID and boolean state');
  const key=decimalId(id),state=scan(pkg);
  if(state.unsupported.length)fail('unsupported','Unsupported comment content must be reviewed before thread resolution');
  const roots=threadRoots(state.comments),rootId=roots.get(key);
  if(rootId===undefined)fail('target','Unknown comment ID');
  const selected=state.comments.filter(c=>roots.get(c.id)===rootId),commentIds=selected.map(c=>c.id),selectedIds=new Set(commentIds);
  for(const c of selected)if(!state.extension||!state.xml||!c.paragraphId||!state.entries.has(c.paragraphId))fail('missing-extension','Every thread member requires an existing linked commentEx entry');
  assertUnprotected(pkg);
  const changed=selected.filter(c=>c.resolved!==resolved);
  if(!changed.length)return {rootId,commentIds,changed:0,changedParts:[]};
  const part=state.extension!,xml=state.xml!,next=applyEdits(xml,changed.map(c=>doneEdit(xml,state.entries.get(c.paragraphId!)!,resolved)));
  return pkg.transaction(()=>{
    writeCommentXml(pkg,part,next);
    const checked=scan(pkg),expected=state.comments.map(c=>selectedIds.has(c.id)?{...c,resolved}:c);
    if(JSON.stringify(checked.comments)!==JSON.stringify(expected)||checked.unsupported.length)fail('state','Thread update changed unexpected comment metadata');
    pkg.toBytes();
    return {rootId,commentIds,changed:changed.length,changedParts:[part]};
  });
}

function threadRoots(comments:CommentInfo[]):Map<string,string> {
  if(comments.length>10000)fail('limit','At most 10000 comments are supported by thread operations');
  const byId=new Map(comments.map(c=>[c.id,c])),roots=new Map<string,string>();
  // scan() has already rejected missing parents and cycles.
  for(const comment of comments){let c=comment;const chain:string[]=[];while(!roots.has(c.id)&&c.parentId!==undefined){chain.push(c.id);c=byId.get(c.parentId)!;}const root=roots.get(c.id)??c.id;roots.set(c.id,root);for(const id of chain)roots.set(id,root);}
  return roots;
}
function writeCommentXml(pkg:OpcPackage,part:string,xml:string):void {
  const bytes=pkg.get(part)!;pkg.set(part,bytes[0]===239&&bytes[1]===187&&bytes[2]===191?'\ufeff'+xml:xml);
}

/** Set only one existing commentEx done flag; no cascading thread state. */
export function setCommentResolved(pkg: OpcPackage, id: string, resolved: boolean): { changed: number; changedParts: string[] } {
  if (typeof id !== 'string' || !/^[0-9]+$/.test(id) || typeof resolved !== 'boolean') fail('argument', 'Expected a decimal comment ID and boolean state');
  const state = scan(pkg);
  if (state.unsupported.length) fail('unsupported', 'Unsupported comment content must be reviewed before resolution');
  const comment = state.comments.find(c => c.id === decimalId(id));
  if (!comment) fail('target', 'Unknown comment ID');
  if (!state.extension || !state.xml || !comment.paragraphId || !state.entries.has(comment.paragraphId)) fail('missing-extension', 'Resolution requires an existing linked commentEx entry');
  assertUnprotected(pkg);
  if (comment.resolved === resolved) return { changed: 0, changedParts: [] };
  const part = state.extension!, xml = state.xml!, entry = state.entries.get(comment.paragraphId!)!;
  const next = editDone(xml, entry, resolved);
  return pkg.transaction(() => {
    writeCommentXml(pkg, part, next);
    const checked = scan(pkg).comments.find(c => c.id === comment.id);
    if (!checked || checked.resolved !== resolved) fail('state', 'Resolution did not preserve the selected comment identity');
    // Serialization validation belongs inside the rollback boundary.
    pkg.toBytes();
    return { changed: 1, changedParts: [part] };
  });
}

function editDone(xml: string, entry: XmlElement, resolved: boolean): string {
  return applyEdits(xml,[doneEdit(xml,entry,resolved)]);
}
function doneEdit(xml:string,entry:XmlElement,resolved:boolean):{start:number;end:number;value:string} {
  const opening = xml.slice(entry.start, entry.openEnd), value = resolved ? '1' : '0';
  const name = Object.keys(entry.attributes).find(key => key.split(':').at(-1) === 'done' && entry.attributeNamespaces[key] === W15);
  if (name) {
    // Match whole attributes, consuming quoted values so lookalike text within
    // another value cannot be mistaken for the target. The XML is already parsed.
    const attributes = /[ \t\r\n]+([^\s=<>/"']+)[ \t\r\n]*=[ \t\r\n]*(["'])([\s\S]*?)\2/g;
    for (const match of opening.matchAll(attributes)) {
      if (match[1] !== name) continue;
      const start = entry.start + match.index! + match[0].length - match[3]!.length - 1;
      return { start, end: start + match[3]!.length, value };
    }
    fail('state', 'Cannot locate the existing done attribute');
  }
  // paraId is required and namespace-qualified, so its prefix is already in scope
  // even when the commentEx element itself uses a default namespace.
  const identity = Object.keys(entry.attributes).find(key => key.split(':').at(-1) === 'paraId' && entry.attributeNamespaces[key] === W15);
  if (!identity?.includes(':')) fail('state', 'Missing namespace-qualified comment identity');
  const suffix = opening.match(/[ \t\r\n]*\/?>$/);
  if (!suffix || suffix.index === undefined) fail('state', 'Cannot locate commentEx opening boundary');
  const start = entry.start + suffix.index;
  return { start, end: start, value: ` ${identity.split(':')[0]}:done="${value}"` };
}

function linkedPart(pkg: OpcPackage, type: string, mime: string): string | undefined {
  const links = pkg.relationships(pkg.mainPart()).filter(r => r.type === type);
  if (links.length > 1 || links.some(r => r.external)) fail('relationship', 'Comment relationships must be unique and internal');
  const part = links[0]?.resolved;
  if (part && getContentType(pkg, part) !== mime) fail('content-type', `Unexpected content type for ${part}`);
  return part;
}

function scan(pkg: OpcPackage): CommentScan {
  const main = parseXml(pkg.text(pkg.mainPart()));
  if (main.root.namespaceURI !== W || main.root.localName !== 'document') fail('document', 'Expected a Word document');
  const part = linkedPart(pkg, COMMENTS_REL, COMMENTS_TYPE);
  const extension = linkedPart(pkg, EXTENDED_REL, EXTENDED_TYPE);
  const comments: CommentInfo[] = [], unsupported: CommentFinding[] = [], entries = new Map<string, XmlElement>();
  const finding = (part: string, kind: string) => { if (!unsupported.some(f => f.part === part && f.kind === kind)) unsupported.push({ part, kind }); };
  // Avoid adopting or overwriting unlinked parts. New parts are outside this API.
  for (const name of pkg.names()) {
    const type = getContentType(pkg, name);
    if ((type === COMMENTS_TYPE && name !== part) || (type === EXTENDED_TYPE && name !== extension)) finding(name, 'unlinked-comment-part');
  }
  if (!part) {
    if (extension) fail('graph', 'Extended comments have no comments owner');
    return { comments, unsupported, entries };
  }
  const source = pkg.text(part), doc = parseXml(source);
  if (doc.root.namespaceURI !== W || doc.root.localName !== 'comments') fail('root', 'Unexpected comments root');
  const ids = new Set<string>(), paragraphIds = new Set<string>(), byParagraph = new Map<string, CommentInfo>();
  if (mixedContent(doc.root, source)) finding(part, 'mixed-content');
  for (const node of doc.root.children) {
    if (node.namespaceURI !== W || node.localName !== 'comment') { finding(part, 'unknown-comment-child'); continue; }
    const id = decimalId(attribute(node, 'id', W));
    if (ids.has(id)) fail('identifier', 'Duplicate comment ID');
    ids.add(id);
    const paragraphs = node.children.filter(n => n.namespaceURI === W && n.localName === 'p');
    if (!paragraphs.length || paragraphs.length !== node.children.length || mixedContent(node, source)) finding(part, 'comment-body');
    for (const p of paragraphs) {
      const raw = attribute(p, 'paraId', W14);
      if (raw === undefined) continue;
      const para = paragraphId(raw);
      if (paragraphIds.has(para)) fail('identifier', 'Duplicate comment paragraph ID');
      paragraphIds.add(para);
    }
    const last = paragraphs.at(-1), rawPara = last && attribute(last, 'paraId', W14);
    const result: CommentInfo = {
      id, author: attribute(node, 'author', W), initials: attribute(node, 'initials', W), date: attribute(node, 'date', W),
      text: paragraphs.map(p => paragraphText(p, source, () => finding(part, 'comment-body'))).join('\n'),
      ...(rawPara === undefined ? {} : { paragraphId: paragraphId(rawPara) }), resolved: false,
    };
    comments.push(result);
    if (result.paragraphId) byParagraph.set(result.paragraphId, result);
  }
  if (!extension) return { comments, unsupported, entries };
  const xml = pkg.text(extension), ex = parseXml(xml);
  if (ex.root.namespaceURI !== W15 || ex.root.localName !== 'commentsEx') fail('root', 'Unexpected commentsExtended root');
  if (mixedContent(ex.root, xml)) finding(extension, 'mixed-content');
  for (const entry of ex.root.children) {
    if (entry.namespaceURI !== W15 || entry.localName !== 'commentEx') { finding(extension, 'unknown-extension-child'); continue; }
    if (entry.children.length || mixedContent(entry, xml)) finding(extension, 'extension-content');
    for (const key of Object.keys(entry.attributes)) {
      if (['paraId', 'paraIdParent', 'done'].includes(key.split(':').at(-1)!) && entry.attributeNamespaces[key] !== W15) fail('namespace', 'Comment metadata must use the commentsExtended namespace');
    }
    const key = paragraphId(attribute(entry, 'paraId', W15));
    if (entries.has(key)) fail('identifier', 'Duplicate commentEx paragraph ID');
    entries.set(key, entry);
    const comment = byParagraph.get(key);
    if (!comment) fail('graph', 'commentEx must address the last paragraph of an existing comment');
    const done = attribute(entry, 'done', W15);
    if (done !== undefined && !['0', '1', 'true', 'false', 'on', 'off'].includes(done)) fail('state', 'Invalid comment done state');
    comment.resolved = done === '1' || done === 'true' || done === 'on';
    const parent = attribute(entry, 'paraIdParent', W15);
    if (parent !== undefined) {
      const owner = byParagraph.get(paragraphId(parent));
      if (!owner) fail('graph', 'Reply parent does not exist');
      comment.parentId = owner.id;
    }
  }
  const byId = new Map(comments.map(c => [c.id, c]));
  const checked = new Set<string>();
  for (const comment of comments) {
    const chain = new Set<string>(); let current: CommentInfo | undefined = comment;
    while (current && !checked.has(current.id)) {
      if (chain.has(current.id)) fail('graph', 'Comment reply graph contains a cycle');
      chain.add(current.id); current = current.parentId === undefined ? undefined : byId.get(current.parentId);
    }
    for (const id of chain) checked.add(id);
  }
  return { comments, unsupported, entries, extension, xml };
}

function paragraphText(p: XmlElement, xml: string, unsupported: () => void): string {
  if (mixedContent(p, xml)) unsupported();
  let text = '';
  for (const child of p.children) {
    if (child.namespaceURI === W && child.localName === 'pPr') continue;
    if (child.namespaceURI !== W || child.localName !== 'r') { unsupported(); continue; }
    if (mixedContent(child, xml)) unsupported();
    for (const leaf of child.children) {
      if (leaf.namespaceURI !== W) { unsupported(); continue; }
      if (leaf.localName === 'rPr' || leaf.localName === 'annotationRef') continue;
      if (leaf.children.length) { unsupported(); continue; }
      if (leaf.localName === 't') text += leaf.text;
      else if (leaf.localName === 'tab') text += '\t';
      else if (leaf.localName === 'br' || leaf.localName === 'cr') text += '\n';
      else unsupported();
    }
  }
  return text;
}
function mixedContent(node: XmlElement, xml: string): boolean {
  if (node.selfClosing) return false;
  let cursor = node.openEnd;
  for (const child of node.children) { if (!/^[ \t\r\n]*$/.test(xml.slice(cursor, child.start))) return true; cursor = child.end; }
  return !/^[ \t\r\n]*$/.test(xml.slice(cursor, node.closeStart));
}
function decimalId(value: string | undefined): string {
  if (value === undefined || !/^[0-9]+$/.test(value) || !Number.isSafeInteger(Number(value))) fail('identifier', 'Comment IDs must be nonnegative safe decimal integers');
  return String(Number(value));
}
function paragraphId(value: string | undefined): string {
  if (!value || !/^[0-9A-Fa-f]{8}$/.test(value)) fail('identifier', 'Expected an eight-digit hexadecimal paragraph ID');
  return value.toUpperCase();
}
function assertUnprotected(pkg: OpcPackage): void {
  for (const rel of pkg.relationships(pkg.mainPart())) {
    if (rel.type !== OFFICE + 'settings') continue;
    if (rel.external) fail('protected', 'External settings cannot be checked');
    const settings = parseXml(pkg.text(rel.resolved!));
    if (settings.root.namespaceURI !== W || settings.root.localName !== 'settings') fail('protected', 'Invalid settings root');
    for (const node of settings.elements) if (node.namespaceURI === W && node.localName === 'documentProtection') {
      if (!/^(0|false|off)$/i.test(attribute(node, 'enforcement', W) ?? '')) fail('protected', 'Document protection refuses comment resolution');
    }
  }
}
