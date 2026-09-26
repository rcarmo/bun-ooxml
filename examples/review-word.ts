import { Document, OpcPackage, inspectStories, inspectRevisions, trackedReplace, resolveRevisions } from '../src/index.ts';

const document = Document.create();
document.addParagraph('Payment within thirty days.', { bold: true });
const pkg = await OpcPackage.open(await document.save());
const before = pkg.toBytes();
trackedReplace(pkg, pkg.mainPart(), 'thirty', 'sixty', {
  author: 'Reviewer', date: '2026-09-26T12:00:00Z',
});
const reviewed = await OpcPackage.open(pkg.toBytes());
if (inspectRevisions(reviewed).revisions.length !== 2) throw new Error('Missing redline');
for (const action of ['accept', 'reject'] as const) {
  const copy = await OpcPackage.open(reviewed.toBytes());
  resolveRevisions(copy, action);
  const saved = await OpcPackage.open(copy.toBytes());
  const text = inspectStories(saved).stories[0]!.paragraphs[0]!.text;
  const expected = action === 'accept' ? 'Payment within sixty days.' : 'Payment within thirty days.';
  if (text !== expected) throw new Error(`${action} text mismatch`);
  const original = await OpcPackage.open(before);
  for (const name of original.names()) {
    if (name !== original.mainPart() && !Buffer.from(original.get(name)!).equals(Buffer.from(saved.get(name)!))) {
      throw new Error(`Unrelated part changed: ${name}`);
    }
  }
}
console.log('Native Word redline saved/reopened; accept/reject text and unrelated parts verified.');
