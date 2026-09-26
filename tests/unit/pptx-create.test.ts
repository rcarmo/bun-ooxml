import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import { describe, expect, test } from "bun:test";
import { join, resolve } from "node:path";

import { bindings as createBindings, runAppendPreservationScenario, runCreateMinimalScenario, runCreateRefusalScenario } from "../acceptance/create-pptx.ts";
import { bindings as parentBindings } from "../acceptance/steps.ts";
import {
  executeAcceptance,
  newAcceptanceRunId,
  parseFeature,
  type AcceptanceFeature,
  type AcceptanceInventory,
} from "../../scripts/gherkin.ts";

const OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const PRES_PROPS_RELATIONSHIP = `${OFFICE_REL_NS}/presProps`;
const SLIDE_RELATIONSHIP = `${OFFICE_REL_NS}/slide`;
const SLIDE_LAYOUT_RELATIONSHIP = `${OFFICE_REL_NS}/slideLayout`;
const SLIDE_MASTER_RELATIONSHIP = `${OFFICE_REL_NS}/slideMaster`;
const TABLE_STYLES_RELATIONSHIP = `${OFFICE_REL_NS}/tableStyles`;
const THEME_RELATIONSHIP = `${OFFICE_REL_NS}/theme`;
const VIEW_PROPS_RELATIONSHIP = `${OFFICE_REL_NS}/viewProps`;

const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const CONTENT_TYPES_PART = "[Content_Types].xml";
const ROOT_RELS_PART = "_rels/.rels";
const PRESENTATION_PART = "ppt/presentation.xml";
const PRESENTATION_RELS_PART = "ppt/_rels/presentation.xml.rels";
const PRES_PROPS_PART = "ppt/presProps.xml";
const VIEW_PROPS_PART = "ppt/viewProps.xml";
const TABLE_STYLES_PART = "ppt/tableStyles.xml";
const THEME_PART = "ppt/theme/theme1.xml";
const SLIDE_MASTER_PART = "ppt/slideMasters/slideMaster1.xml";
const SLIDE_MASTER_RELS_PART = "ppt/slideMasters/_rels/slideMaster1.xml.rels";
const SLIDE_LAYOUT_PART = "ppt/slideLayouts/slideLayout1.xml";
const SLIDE_LAYOUT_RELS_PART = "ppt/slideLayouts/_rels/slideLayout1.xml.rels";
const SLIDE1_PART = "ppt/slides/slide1.xml";
const SLIDE1_RELS_PART = "ppt/slides/_rels/slide1.xml.rels";
const SLIDE2_PART = "ppt/slides/slide2.xml";
const SLIDE2_RELS_PART = "ppt/slides/_rels/slide2.xml.rels";

describe("Presentation.create and addTextSlide", () => {
  test("authors a minimal presentation whose master, layout and theme links reopen through the native reader", async () => {
    const result = await runCreateMinimalScenario();

    expect(result.slideCount).toBe(2);
    expect(result.slideMasterCount).toBe(1);
    expect(result.slideLayoutCount).toBe(1);
    expect(result.mainPartName).toBe(PRESENTATION_PART);
    expect(result.rootOfficeDocumentTarget).toBe(PRESENTATION_PART);
    expect(result.reopenedTexts).toEqual([
      ["Created title 1", "Created subtitle 1"],
      ["Created title 2", "Created subtitle 2"],
    ]);
    expect(result.savedPartNames).toEqual([
      CONTENT_TYPES_PART,
      ROOT_RELS_PART,
      PRESENTATION_RELS_PART,
      PRES_PROPS_PART,
      PRESENTATION_PART,
      SLIDE_LAYOUT_RELS_PART,
      SLIDE_LAYOUT_PART,
      SLIDE_MASTER_RELS_PART,
      SLIDE_MASTER_PART,
      SLIDE1_RELS_PART,
      SLIDE2_RELS_PART,
      SLIDE1_PART,
      SLIDE2_PART,
      TABLE_STYLES_PART,
      THEME_PART,
      VIEW_PROPS_PART,
    ].sort());
    expect(result.contentTypeOverrides).toEqual([
      "/ppt/presProps.xml",
      "/ppt/presentation.xml",
      "/ppt/slideLayouts/slideLayout1.xml",
      "/ppt/slideMasters/slideMaster1.xml",
      "/ppt/slides/slide1.xml",
      "/ppt/slides/slide2.xml",
      "/ppt/tableStyles.xml",
      "/ppt/theme/theme1.xml",
      "/ppt/viewProps.xml",
    ]);
    expect(result.presentationRelationshipTypes.sort()).toEqual([
      PRES_PROPS_RELATIONSHIP,
      SLIDE_MASTER_RELATIONSHIP,
      SLIDE_RELATIONSHIP,
      SLIDE_RELATIONSHIP,
      TABLE_STYLES_RELATIONSHIP,
      THEME_RELATIONSHIP,
      VIEW_PROPS_RELATIONSHIP,
    ].sort());
    expect(result.presentationSlideMasterTarget).toBe(SLIDE_MASTER_PART);
    expect(result.presentationThemeTarget).toBe(THEME_PART);
    expect(result.presentationPresPropsTarget).toBe(PRES_PROPS_PART);
    expect(result.presentationViewPropsTarget).toBe(VIEW_PROPS_PART);
    expect(result.presentationTableStylesTarget).toBe(TABLE_STYLES_PART);
    expect(result.presentationSlideTargets).toEqual([SLIDE1_PART, SLIDE2_PART]);
    expect(result.slideLayoutTargets).toEqual([SLIDE_LAYOUT_PART, SLIDE_LAYOUT_PART]);
    expect(result.masterRelationshipTypes.sort()).toEqual([
      SLIDE_LAYOUT_RELATIONSHIP,
      THEME_RELATIONSHIP,
    ].sort());
    expect(result.masterLayoutTargets).toEqual([SLIDE_LAYOUT_PART]);
    expect(result.masterThemeTarget).toBe(THEME_PART);
    expect(result.layoutRelationshipTypes).toEqual([SLIDE_MASTER_RELATIONSHIP]);
    expect(result.layoutMasterTarget).toBe(SLIDE_MASTER_PART);
    expect(result.reopenedDiff).toEqual({ added: [], changed: [], removed: [] });
  });

  test("appends to a compatible deck without disturbing first-slide bytes until the earlier anchor edits them", async () => {
    const result = await runAppendPreservationScenario();

    expect(result.reopenedTexts).toEqual([
      ["Edited original title", "Original subtitle"],
      ["Appended title", "Appended subtitle"],
    ]);
    expect(result.originalTitleAnchorVersion).toBe(result.afterAppendTitleAnchorVersion);
    expect(result.originalTitleAnchorParagraphIndex).toBe(result.afterAppendTitleAnchorParagraphIndex);
    expect(result.slide1XmlAfterAppend).toBe(result.originalSlide1Xml);
    expect(result.finalSlide1Xml).toBe(result.expectedSlide1Xml);
    expect(result.finalSlide1Rels).toEqual(result.originalSlide1Rels);
    expect(result.savedPartNames).toContain(SLIDE1_RELS_PART);
    expect(result.savedPartNames).toContain(SLIDE2_RELS_PART);
    expect(result.presentationSlideTargets).toEqual([SLIDE1_PART, SLIDE2_PART]);
    expect(result.slideLayoutTargets).toEqual([SLIDE_LAYOUT_PART, SLIDE_LAYOUT_PART]);
    expect(result.diff).toEqual({
      added: [SLIDE2_RELS_PART, SLIDE2_PART],
      changed: [CONTENT_TYPES_PART, PRESENTATION_RELS_PART, PRESENTATION_PART, SLIDE1_PART],
      removed: [],
    });
  });

  test("refuses invalid arguments and unsafe title-layout selection atomically", async () => {
    const result = await runCreateRefusalScenario();

    expect(result.invalidError.code).toBe("PPTX_ARGUMENT_INVALID");
    expect(result.invalidAfter).toEqual(result.invalidBefore);
    expect(result.invalidDiff).toEqual({ added: [], changed: [], removed: [] });
    expect(result.unsafeError.code).toBe("PPTX_LAYOUT_UNSAFE");
    expect(result.unsafeAfter).toEqual(result.unsafeBefore);
    expect(result.unsafeDiff).toEqual({ added: [], changed: [], removed: [] });
  });

  test("executes the PPTX create feature with dedicated bindings and optional parent aggregation", async () => {
    const featurePath = "references/fixtures-ooxml/workflows/pptx/creation.feature";
    const feature = parseFeature(featurePath, (await Bun.file(join(fixturesRoot(), "workflows/pptx/creation.feature")).text()).replace(/^@planned/m, '@implemented @bun'));
    const inventory = inventoryFor(feature);

    for (const bindings of [createBindings, parentBindings]) {
      const execution = await executeAcceptance(inventory, bindings, newAcceptanceRunId());
      expect(execution.failures).toEqual([]);
      expect(execution.counts.features).toEqual({ passed: 1, failed: 0, planned: 0, total: 1 });
      expect(execution.counts.scenarios).toEqual({ passed: 3, failed: 0, planned: 0, total: 3 });
      expect(execution.counts.cases).toEqual({ passed: 3, failed: 0, planned: 0, total: 3 });
      expect(execution.counts.steps.failed).toBe(0);
      expect(execution.counts.steps.undefined).toBe(0);
      expect(execution.counts.steps.ambiguous).toBe(0);
    }
  }, 300_000);
});

function inventoryFor(feature: AcceptanceFeature): AcceptanceInventory {
  const cases = feature.scenarios.flatMap((scenario) => scenario.cases);
  const steps = cases.flatMap((acceptanceCase) => acceptanceCase.steps).length;
  const count = (implemented: number) => ({ implemented, planned: 0, total: implemented });
  return {
    root: PROJECT_ROOT,
    features: [feature],
    counts: {
      features: count(1),
      scenarios: count(feature.scenarios.length),
      cases: count(cases.length),
      steps: count(steps),
    },
  };
}
