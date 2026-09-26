import { lstatSync, readFileSync, realpathSync } from "node:fs";
import { join, resolve } from "node:path";

type ManifestFile = {
  id: string;
  path: string;
  bytes: number;
  sha256: string;
  role: string;
  format?: string;
  scenarioGroup?: string;
};

type Manifest = {
  schemaVersion: number;
  fixturePathBase: string;
  files: ManifestFile[];
};

const PROJECT_ROOT = resolve(import.meta.dir, "..");
const DEFAULT_FIXTURES_ROOT = join(PROJECT_ROOT, "references/fixtures-ooxml");
const HASH = /^[a-f0-9]{64}$/;

let cachedRoot: string | undefined;
let manifestCache: Manifest | undefined;
let filesByIdCache: Map<string, ManifestFile> | undefined;

export const F = {
  goWord: {
    default: "fixture-d9d6a313182a71a73d75a26a0ff3b7826dbd2e300e1d202114ec9f8fb018fda5",
    bulletList: "fixture-cb2c2610f99f786484b1941f25d068f6347cd0cd95213a57db849842d9d4a64b",
    comments: "fixture-2029abbda3bdacb270137f6210791ba2ee3f8be62affb8cc3181f3ac72a7c35a",
    complexTable: "fixture-28da984e3fc4c6079579c0e8b3b9a901a46ed3f9081f3a074e0fb7333b984695",
    formattedText: "fixture-9a92eba3dc84f293a82de9496e571c780fd258cc3bf276b74464faaef5835dbc",
    headersFooters: "fixture-3bafa155242222dbd3529b56af6b8f1939cbd2e954de5b80afeadbbd53a432aa",
    headings: "fixture-8513f05370714f5e288ec1ca2fb76fe21b74b458636f325666b20b100c9b021a",
    minimal: "fixture-9726b477472ddb7595875c9f30493df2577e587416b18418d0dc7221046690fe",
    numberedList: "fixture-37d3c408403dbecf4310f0a0b1313dc0c3756302e1778cc32d975823988ea3b2",
    sdtContentControls: "fixture-368fe96cb3ae55a0cc5fecbb599eda1d1058596d4914992f596083f291071ae4",
    simpleTable: "fixture-87e3c67cb73bdbf5c8389791fd2bb459ba0eaac862149af6fd7c5b641dafd56b",
    singleParagraph: "fixture-395b75b992c6662e2bf44c72401020efccc212b41a345e082cb09abee450d97f",
    styles: "fixture-d35e32ab35d4c95e81c15c5baac18b842696adcea960664e329a5a9a699ab7a1",
    trackChanges: "fixture-2e1022f7358f0deeaeb542a6dd10005f455c433a66f1b9b7e8964a48c7f0fb2a",
  },
  officeWord: {
    default: "fixture-b051c0c2ff43f2ab9213e19a52ccbc51217537cf3b336d54a68a91beb6670f9a",
    bulletList: "fixture-fad9dd22fb2d0ef888452d678a5b1bce2b15202af16fc4ed799acb230b33798d",
    comments: "fixture-ccdfb41723d543a8baf3444b16c3f8fa7d8473aead590e13825042ba6119da62",
    complexTable: "fixture-10737b881f1643cd13c3e6b664d3116da96c5cfdc9f9fb76548ff2dd852181f3",
    formattedText: "fixture-12183fb28e49ea1c1ac63de2252b011580c4cfed0a7cfd35efc1ffb20c94653e",
    headersFooters: "fixture-a99df5aa88c56707c994eabc15ccec7f811d5bd2b540b8cb77c90ebeecaa6933",
    headings: "fixture-0e3d6fb95187c7e0d5d5b1bf19451b58b617411ed8433911d8c95eab15f09f01",
    minimal: "fixture-291ea45fd599555956c86a7dd98c96c06ded43c5d03a57888422f6e34ea1203c",
    numberedList: "fixture-feaa592c7914153be8cb0ebe30b1a1b3d82fcb15d0b05aeca74f89bb3bbeee12",
    sdtContentControls: "fixture-e4f051ec2eb5f48b9b8299e931abb2bca1fa86ca5007865b3b2f9b83ba16676f",
    simpleTable: "fixture-8192955ef935f09eb61a9fe6805d4996c811efcf54c0c966f52d983e38e0a79c",
    singleParagraph: "fixture-dfca453a4b60629ec859b7e224040148b1e2fb92aa3a32b28a6ad9c6f4cc3526",
    styles: "fixture-9548a1ce68caae9df12bc85732f1c19a098658c5dce3d79488814e4145299e5e",
    trackChanges: "fixture-e3c5159fbf254f4d5423354773ae83a5535cb3cef8f603f1adca6d18adab11f2",
  },
  goSlides: {
    default: "fixture-151d747bc37d4f4988c1116f4abb45196b1c1644319ce342ee4dd56d111f3132",
    bulletPoints: "fixture-5c615f27a23cb0d5fe3d603080ba057c7fc320f2af1df277b5b0dd128f1a8e3d",
    comments: "fixture-ababb5bed8eeed4511f25e23da8e9039642ed3020b1517622d3ed71f5d5ef319",
    hiddenSlides: "fixture-e01ded1106a28f94a3439e8368f9a12ec360891f4a9e2810f6504c4c328ed79c",
    images: "fixture-88f5cf211cbedea688b081b2822c9207670e58a665afb1222506672fda3de2a8",
    layouts: "fixture-ef5de48dec43f165d513086506ba52c8f1b9ebbbc2fc99a6b4f56731f6115862",
    minimal: "fixture-6a28461a00850a9297690a0a53ad8c109aea9e6f279df309639e73be0cc4c762",
    multipleMasters: "fixture-e6bbd95968d9fd04bfc13cecc671a7438ab3c99987bc36b4fe118ef99241296f",
    notes: "fixture-04faba67841dda25dc3ff9e3e6e345e6feeeef1cf25a6b9065bf5fbdc83163dc",
    shapes: "fixture-2610748ab308da10d61ebc6e5677f774427a207bdb474b74ebd3ae234152589e",
    tables: "fixture-3465194945a0f8084084ceb843c219d59a4ad92b25a271d66182587d637db0be",
    titleSlide: "fixture-836b5c7d7917a5e053ca994d96af0409d5f6650dd28b02dda9f878f8600c8656",
  },
  officeSlides: {
    default: "fixture-c54a7b746c0328fc1930525edd91387eedbbca69a6f388f7ec024150187b6bab",
    bulletPoints: "fixture-50f9157c8ff3783b9cb6a976dddd834d5c683e42cc9a21170637b191800bc22e",
    comments: "fixture-2c771c00dfe223d0e30c3082de43cc023537d58babb562e26df8034e0a54c24d",
    hiddenSlides: "fixture-fa245a3df00fef7f7bf4739921ee840194040161e06490589e3d52cc9fa7a71d",
    images: "fixture-792277d2c358c5bd7c8a4053078a5594fa07bfb55a06b0bb30f2813c151ea459",
    layouts: "fixture-d902e187a376fb5085ab843234f4ee29b39d019c903b84c33ce134e596b40345",
    minimal: "fixture-e6b4859435d75b21f30a610d766ef64c4f0e926ee0836c9cb389ac2faf7ed5e8",
    multipleMasters: "fixture-7342e20ba487f1bcf7bf3961c93033b9465fae4ddd1e82a187e75cb99ab6fce7",
    notes: "fixture-e97c8d590798576b4b1c27e9e3d98ba73036d99ddeffeafa9ae81dc70a77ff1a",
    shapes: "fixture-10a6d7267ed96fa09cc5424695f4d7a20833bb5247a44a1e71cf57279c425b27",
    tables: "fixture-3998d8058f356eb629c301bbf31cf29625b9eb2c2a08b561932418fb56696a25",
    titleSlide: "fixture-1b848867cffb781112dc5778fa8cc7b9c9bd472c3636a348ec5d50005f05489e",
  },
  goSheets: {
    comments: "fixture-264be55e012d4bc2b3bf25e59824fdd30022e94f70869ea7d6ad960b803a902f",
    conditionalFormat: "fixture-7124469770b5368c036f9dcc13d0c09c382797827d1b58934d9b2a0b61efee28",
    dataTypes: "fixture-13bf5f08697e7de3da52b47d61e0f5af42081a4f1bdfbc1d51a3926910098b68",
    formatting: "fixture-dab1d6dbdfa594680115bc87fc1707fbd06adea87b45ef0a1115d7cf4c4f72b5",
    formulas: "fixture-61d8806a5fd3ff6c9eb62d716a6d929d0947e35ac6d9ac56626704690ecd3877",
    mergedCells: "fixture-691d1d459bd3d2ac3704b889bde121186b802381cf618c0a2daf12ca5b76ebec",
    minimal: "fixture-5140bb7ea22875bc86e3a9790b3dce686a8adc518f7c8597fbddd5d83264b59d",
    multipleSheets: "fixture-a88d1e934bea9d3483325937f29a5d0880065bd54f95bbeceb612f7d4b5575f6",
    namedRanges: "fixture-cc8da5b5969032cbd2bdc6a40b26fc61c2366cf87e3d9455d4fdb8308ceb14fa",
    singleCell: "fixture-78318622a642a298a0ec9626fc03a86719e0d9d2f170aa7e8d15db99ee29dbd9",
    tables: "fixture-51eaa2ee2fa09f384e5c177cebd62801d628f20724c57a221a25112cf374a8fb",
  },
  officeSheets: {
    comments: "fixture-c03ad65508025275563db84b36340934b426560ae259dd572f551a174c443dca",
    conditionalFormat: "fixture-80f3dc2dbfa2178cb0b5bb30dedea12091850ea955278dc9a1328a12044395b8",
    dataTypes: "fixture-bf5ecc732bd2387b953e05b888f5e40937934023f5f61de834a6fb6cd3e349c4",
    formatting: "fixture-ee0ee5c840165d52341ff0c75e1a44007521e176a6fe0232452a5b49973bd6a7",
    formulas: "fixture-9668136d1f23973e901042d770ff5058497eea0ff02ea0b201642e52cf024fea",
    mergedCells: "fixture-444855bd685a55fb37fd098d8024726c914d8f3cc14dffc8f115b5e80b7d2c47",
    minimal: "fixture-9145d25de350a3da09ac8f1976811464525a4eeaa9947baa17057ae85056c377",
    multipleSheets: "fixture-40628979d41435d00a452faf189d98cc00343ee1639070f20fcfc21b471a983d",
    namedRanges: "fixture-de71d259e267ca978b90556b64c6e76e46b81c6f4134ffbde755ffa77a815656",
    singleCell: "fixture-7c4584b7c6a6a9a474fd60dc5ef06a1e5f6d5d830ab1c7d841818bcb0230ea77",
    tables: "fixture-7a10c58d6f99e7285a905561690502f12438080697890924d80c90f9b67c05e8",
  },
  shared: {
    crossSheetCache: "fixture-8ba5708d5030adf93a4f7e4ae466563a1b66341067a200a6cb9b782484dcb5b1",
    defaultStyle: "fixture-38c2ed936696179d3b2359e9107ad2b8d62d71d69296f8f60bdfe1fe8f7f2439",
    presentPlaceholder: "fixture-535910216e3531e4f70959cccf83038f1c51febbf4a149f9f46abfbfa8667d90",
    titleAndSubtitle: "fixture-2aec94471f93c300d56ca4789106a974411085d1588f3424155362a06dd043f3",
  },
} as const;

export const SHARED_FIXTURE_IDS = {
  "cross-sheet-cache.xlsx": F.shared.crossSheetCache,
  "default-style.xlsx": F.shared.defaultStyle,
  "present-placeholder.docx": F.shared.presentPlaceholder,
  "title-and-subtitle.pptx": F.shared.titleAndSubtitle,
} as const;

export const GO_TESTDATA_PACKAGE_IDS = [
  F.goWord.default,
  F.goSlides.default,
  F.goWord.bulletList,
  F.goWord.comments,
  F.goWord.complexTable,
  F.goWord.formattedText,
  F.goWord.headersFooters,
  F.goWord.headings,
  F.goWord.minimal,
  F.goWord.numberedList,
  F.goWord.sdtContentControls,
  F.goWord.simpleTable,
  F.goWord.singleParagraph,
  F.goWord.styles,
  F.goWord.trackChanges,
  F.goSlides.bulletPoints,
  F.goSlides.comments,
  F.goSlides.hiddenSlides,
  F.goSlides.images,
  F.goSlides.layouts,
  F.goSlides.minimal,
  F.goSlides.multipleMasters,
  F.goSlides.notes,
  F.goSlides.shapes,
  F.goSlides.tables,
  F.goSlides.titleSlide,
  F.goSheets.comments,
  F.goSheets.conditionalFormat,
  F.goSheets.dataTypes,
  F.goSheets.formatting,
  F.goSheets.formulas,
  F.goSheets.mergedCells,
  F.goSheets.minimal,
  F.goSheets.multipleSheets,
  F.goSheets.namedRanges,
  F.goSheets.singleCell,
  F.goSheets.tables,
] as const;

export const PYTHON_TEMPLATE_PACKAGE_IDS = [
  F.officeWord.default,
  F.officeSlides.default,
  F.officeWord.bulletList,
  F.officeWord.comments,
  F.officeWord.complexTable,
  F.officeWord.formattedText,
  F.officeWord.headersFooters,
  F.officeWord.headings,
  F.officeWord.minimal,
  F.officeWord.numberedList,
  F.officeWord.sdtContentControls,
  F.officeWord.simpleTable,
  F.officeWord.singleParagraph,
  F.officeWord.styles,
  F.officeWord.trackChanges,
  F.officeSlides.bulletPoints,
  F.officeSlides.comments,
  F.officeSlides.hiddenSlides,
  F.officeSlides.images,
  F.officeSlides.layouts,
  F.officeSlides.minimal,
  F.officeSlides.multipleMasters,
  F.officeSlides.notes,
  F.officeSlides.shapes,
  F.officeSlides.tables,
  F.officeSlides.titleSlide,
  F.officeSheets.comments,
  F.officeSheets.conditionalFormat,
  F.officeSheets.dataTypes,
  F.officeSheets.formatting,
  F.officeSheets.formulas,
  F.officeSheets.mergedCells,
  F.officeSheets.minimal,
  F.officeSheets.multipleSheets,
  F.officeSheets.namedRanges,
  F.officeSheets.singleCell,
  F.officeSheets.tables,
] as const;

export const PYTHON_TESTDATA_PACKAGE_IDS = [
  F.officeWord.bulletList,
  F.officeWord.comments,
  F.officeWord.complexTable,
  F.officeWord.formattedText,
  F.officeWord.headersFooters,
  F.officeWord.headings,
  F.officeWord.minimal,
  F.officeWord.numberedList,
  F.officeWord.sdtContentControls,
  F.officeWord.simpleTable,
  F.officeWord.singleParagraph,
  F.officeWord.styles,
  F.officeWord.trackChanges,
  F.officeSlides.bulletPoints,
  F.officeSlides.comments,
  F.officeSlides.hiddenSlides,
  F.officeSlides.images,
  F.officeSlides.layouts,
  F.officeSlides.minimal,
  F.officeSlides.multipleMasters,
  F.officeSlides.notes,
  F.officeSlides.shapes,
  F.officeSlides.tables,
  F.officeSlides.titleSlide,
  F.officeSheets.comments,
  F.officeSheets.conditionalFormat,
  F.officeSheets.dataTypes,
  F.officeSheets.formatting,
  F.officeSheets.formulas,
  F.officeSheets.mergedCells,
  F.officeSheets.minimal,
  F.officeSheets.multipleSheets,
  F.officeSheets.namedRanges,
  F.officeSheets.singleCell,
  F.officeSheets.tables,
] as const;

export const CORPUS74_PACKAGE_IDS = [
  ...GO_TESTDATA_PACKAGE_IDS,
  ...PYTHON_TEMPLATE_PACKAGE_IDS,
] as const;

export const GO_TESTDATA_ALL_FILE_COUNT = 39;
export const GO_TESTDATA_PACKAGE_COUNT = GO_TESTDATA_PACKAGE_IDS.length;
export const PYTHON_TESTDATA_ALL_FILE_COUNT = 35;
export const PYTHON_TESTDATA_PACKAGE_COUNT = PYTHON_TESTDATA_PACKAGE_IDS.length;

export function fixturesRoot(): string {
  const root = resolve(process.env.OOXML_FIXTURES_ROOT ?? DEFAULT_FIXTURES_ROOT);
  assertDirectory(root, `Missing fixture root: ${root}`);
  if (lstatSync(root).isSymbolicLink()) {
    throw new Error(`Fixture root must not be a symlink: ${root}`);
  }
  return realpathSync(root);
}

export function fixtureManifest(): Manifest {
  const root = fixturesRoot();
  if (manifestCache && cachedRoot === root) {
    return manifestCache;
  }

  const manifestPath = join(root, "manifest.json");
  const manifest = JSON.parse(readFileSync(manifestPath, "utf8")) as Manifest;
  validateManifest(manifest);

  const filesById = new Map<string, ManifestFile>();
  for (const file of manifest.files) {
    if (filesById.has(file.id)) {
      throw new Error(`Duplicate fixture id: ${file.id}`);
    }
    filesById.set(file.id, file);
  }

  cachedRoot = root;
  manifestCache = manifest;
  filesByIdCache = filesById;
  return manifest;
}

export function fixturePath(id: string): string {
  const file = fixtureFile(id);
  if (file.role !== "fixture") {
    throw new Error(`Manifest entry is not a fixture: ${id}`);
  }
  validateFixtureFile(file, id);

  const root = fixturesRoot();
  const path = resolve(root, file.path);
  if (!path.startsWith(`${root}/`) && path !== root) {
    throw new Error(`Fixture path escapes root: ${id}`);
  }

  const stat = lstatSync(path, { throwIfNoEntry: false });
  if (!stat) {
    throw new Error(`Missing fixture bytes for ${id}: ${path}`);
  }
  if (stat.isSymbolicLink() || realpathSync(path) !== path) {
    throw new Error(`Fixture bytes must not be symlinked: ${path}`);
  }
  if (!stat.isFile()) {
    throw new Error(`Fixture bytes must be a file: ${path}`);
  }
  if (stat.size !== file.bytes) {
    throw new Error(`Fixture byte length drift for ${id}: ${stat.size} !== ${file.bytes}`);
  }
  const hash = new Bun.CryptoHasher('sha256').update(readFileSync(path)).digest('hex');
  if (hash !== file.sha256) throw new Error(`Fixture hash drift for ${id}`);
  return path;
}

export function fixturePaths(ids: readonly string[]): string[] {
  return ids.map((id) => fixturePath(id));
}

function fixtureFile(id: string): ManifestFile {
  fixtureManifest();
  const file = filesByIdCache?.get(id);
  if (!file) {
    throw new Error(`Unknown fixture id: ${id}`);
  }
  return file;
}

function validateManifest(manifest: Manifest): void {
  if (manifest.schemaVersion !== 2) {
    throw new Error(`Unsupported fixture manifest schema: ${manifest.schemaVersion}`);
  }
  if (manifest.fixturePathBase !== "repository-root") {
    throw new Error(`Unsupported fixture path base: ${manifest.fixturePathBase}`);
  }
  if (!Array.isArray(manifest.files) || manifest.files.length === 0) {
    throw new Error("Fixture manifest must contain files");
  }
  for (const file of manifest.files) {
    validateManifestFile(file);
  }
}

function validateManifestFile(file: ManifestFile): void {
  if (typeof file.id !== "string" || file.id.length === 0) {
    throw new Error("Fixture manifest id must be a non-empty string");
  }
  if (typeof file.path !== "string" || !isSafeRelativePath(file.path)) {
    throw new Error(`Unsafe fixture path: ${String(file.path)}`);
  }
  if (!Number.isSafeInteger(file.bytes) || file.bytes < 0) {
    throw new Error(`Invalid fixture byte length for ${file.id}`);
  }
  if (typeof file.sha256 !== "string" || !HASH.test(file.sha256)) {
    throw new Error(`Invalid fixture hash for ${file.id}`);
  }
  if (file.role === "fixture") {
    validateFixtureFile(file, file.id);
  }
}

function validateFixtureFile(file: ManifestFile, id: string): void {
  if (id !== 'fixture-' + file.sha256) throw new Error(`Fixture ID/hash mismatch: ${id}`);
  if (typeof file.format !== "string" || !/^(docx|pptx|xlsx|png)$/.test(file.format)) {
    throw new Error(`Fixture format missing or invalid for ${id}`);
  }
  if (typeof file.scenarioGroup !== "string" || !/^[a-z0-9-]+$/.test(file.scenarioGroup)) {
    throw new Error(`Fixture scenario group missing or invalid for ${id}`);
  }
  const expectedPrefix = `fixtures/${file.format}/${file.scenarioGroup}/`;
  if (!file.path.startsWith(expectedPrefix)) {
    throw new Error(`Fixture path is not grouped canonically for ${id}: ${file.path}`);
  }
  if (!file.path.endsWith(`.${file.format}`)) {
    throw new Error(`Fixture extension mismatch for ${id}: ${file.path}`);
  }
}

function isSafeRelativePath(path: string): boolean {
  if (path.startsWith("/") || path.includes("\\") || /[\u0000-\u001f:]/.test(path)) {
    return false;
  }
  return path.split("/").every((segment) => segment.length > 0 && segment !== "." && segment !== "..");
}

function assertDirectory(path: string, message: string): void {
  const stat = lstatSync(path, { throwIfNoEntry: false });
  if (!stat || !stat.isDirectory()) {
    throw new Error(message);
  }
}
