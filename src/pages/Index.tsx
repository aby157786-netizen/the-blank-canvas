import { ChangeEvent, useMemo, useState } from "react";
import pptxgen from "pptxgenjs";
import {
  BarChart3,
  BriefcaseBusiness,
  Check,
  Clipboard,
  Columns3,
  Copy,
  Download,
  FileJson,
  GraduationCap,
  Image,
  Layers3,
  Lightbulb,
  LineChart,
  ListChecks,
  Palette,
  Presentation,
  Rocket,
  ScrollText,
  Sparkles,
  Table2,
  Upload,
  UserRound,
} from "lucide-react";

type TemplateId =
  | "business"
  | "business-skeuo"
  | "education"
  | "education-glass"
  | "sales"
  | "sales-brutal"
  | "proposal"
  | "proposal-minimal"
  | "resume"
  | "resume-neumorph"
  | "timeline"
  | "timeline-skeuo"
  | "comparison"
  | "comparison-minimal"
  | "brainstorming"
  | "brainstorming-brutal"
  | "report"
  | "report-glass"
  | "creative"
  | "creative-neumorph";

type UiStyle = "Classic" | "Skeuomorphism" | "Neobrutalism" | "Glassmorphism" | "Neumorphism" | "Minimalism";

type LayoutType =
  | "cover"
  | "title-content"
  | "section-header"
  | "two-content"
  | "title-only"
  | "content-caption"
  | "bullets"
  | "comparison"
  | "timeline"
  | "process"
  | "image"
  | "cards"
  | "metrics"
  | "profile"
  | "idea-wall"
  | "proposal"
  | "matrix";

type SlideContent = {
  title: string;
  subtitle?: string;
  bullets: string[];
  notes?: string;
  layout?: LayoutType;
  image?: string;
  imagePrompt?: string;
  comparison?: string[];
  steps?: string[];
  metrics?: string[];
  cards?: string[];
  columns?: string[];
  rows?: string[];
};

type DeckContent = {
  title: string;
  subtitle: string;
  sections: SlideContent[];
};

type Template = {
  id: TemplateId;
  name: string;
  scenario: string;
  uiStyle: UiStyle;
  bestFor: string;
  icon: typeof Presentation;
  layoutBias: LayoutType[];
  colors: {
    bg: string;
    ink: string;
    muted: string;
    panel: string;
    accent: string;
    accent2: string;
    border: string;
    shadow: string;
  };
};

type Outline = {
  id: LayoutType;
  name: string;
  description: string;
  fields: string;
  prompt: string;
  icon: typeof Presentation;
};

const layoutJsonExamples: Record<LayoutType, string> = {
  cover: `{
  "layout": "cover",
  "title": "Topic title",
  "subtitle": "Audience + promise"
}`,
  "title-content": `{
  "layout": "title-content",
  "title": "Main idea",
  "subtitle": "Short context line",
  "bullets": ["Supporting idea 1", "Supporting idea 2", "Supporting idea 3"]
}`,
  "section-header": `{
  "layout": "section-header",
  "title": "Section title",
  "subtitle": "What this section will explain next"
}`,
  "two-content": `{
  "layout": "two-content",
  "title": "Two related content blocks",
  "columns": ["Left side topic", "Right side topic"],
  "bullets": ["Left detail 1", "Left detail 2", "Right detail 1", "Right detail 2"]
}`,
  "title-only": `{
  "layout": "title-only",
  "title": "Big statement or transition title",
  "subtitle": "Optional short explanation"
}`,
  "content-caption": `{
  "layout": "content-caption",
  "title": "Visual or content focus",
  "imagePrompt": "Describe the main image, chart, diagram, or screenshot",
  "bullets": ["Caption insight", "Why it matters"]
}`,
  bullets: `{
  "layout": "bullets",
  "title": "Key takeaways",
  "bullets": ["Point 1", "Point 2", "Point 3", "Point 4"]
}`,
  comparison: `{
  "layout": "comparison",
  "title": "Option A vs Option B",
  "comparison": ["Option A: strength", "Option A: risk", "Option B: strength", "Option B: risk"]
}`,
  timeline: `{
  "layout": "timeline",
  "title": "Roadmap",
  "steps": ["Phase 1: discovery", "Phase 2: build", "Phase 3: launch", "Phase 4: improve"]
}`,
  process: `{
  "layout": "process",
  "title": "Workflow",
  "subtitle": "Short explanation shown below or near the process",
  "steps": ["Step 1", "Step 2", "Step 3", "Step 4", "Step 5"],
  "bullets": ["Explain the purpose of the sequence", "Add any important note below the steps"]
}`,
  image: `{
  "layout": "image",
  "title": "Visual proof",
  "imagePrompt": "Describe the screenshot, photo, diagram, chart, or scene to place here",
  "bullets": ["Caption point", "Why it matters"]
}`,
  cards: `{
  "layout": "cards",
  "title": "Four modules",
  "cards": ["Card 1", "Card 2", "Card 3", "Card 4"]
}`,
  metrics: `{
  "layout": "metrics",
  "title": "Performance snapshot",
  "metrics": ["42%: conversion lift", "$1.2M: revenue impact", "18 days: time saved"]
}`,
  profile: `{
  "layout": "profile",
  "title": "Candidate profile",
  "bullets": ["Role summary", "Core skill", "Experience highlight"],
  "metrics": ["8 years: experience", "12: shipped projects"]
}`,
  "idea-wall": `{
  "layout": "idea-wall",
  "title": "Brainstorm board",
  "cards": ["Idea 1", "Idea 2", "Idea 3", "Idea 4", "Idea 5", "Idea 6"]
}`,
  proposal: `{
  "layout": "proposal",
  "title": "Problem and solution",
  "bullets": ["Problem", "Proposed solution", "Expected impact", "Next move"]
}`,
  matrix: `{
  "layout": "matrix",
  "title": "Decision matrix",
  "columns": ["Criteria", "Option A", "Option B", "Decision"],
  "rows": ["Cost", "Speed", "Risk", "Fit"]
}`,
};

const layoutVariants: Record<LayoutType, string[]> = {
  cover: ["Centered title cover", "Title + audience", "Title + outcome promise", "Title + source-material context"],
  "title-content": ["Title and content", "Title with explanation", "Title with short list"],
  "section-header": ["Section divider", "Chapter opener", "Topic transition"],
  "two-content": ["Two content columns", "Text + text blocks", "Two related lists"],
  "title-only": ["Big transition statement", "Quote-style title", "Simple announcement"],
  "content-caption": ["Picture with caption", "Chart with explanation", "Screenshot with notes"],
  bullets: ["Key takeaways", "Problem list", "Recommendation list"],
  comparison: ["Before vs after", "Option A vs Option B", "Pros and cons"],
  timeline: ["Date-based roadmap", "Phase-based milestones", "Past-present-future story"],
  process: ["Step-by-step workflow", "Implementation plan", "Learning sequence"],
  image: ["Screenshot placeholder", "Diagram placeholder", "Photo or scene placeholder"],
  cards: ["Feature cards", "Module cards", "Chapter cards"],
  metrics: ["KPI snapshot", "Research findings", "Impact numbers"],
  profile: ["Resume profile", "Team profile", "Founder/operator profile"],
  "idea-wall": ["Brainstorm notes", "Creative concepts", "Workshop clusters"],
  proposal: ["Problem-solution", "Scope-value", "Ask-next steps"],
  matrix: ["Feature table", "Decision table", "Evaluation scorecard"],
};

const uiStyles: UiStyle[] = ["Classic", "Skeuomorphism", "Neobrutalism", "Glassmorphism", "Neumorphism", "Minimalism"];

const layoutOutlines: Outline[] = [
  { id: "cover", name: "Cover page", description: "Deck title, subtitle, topic and audience framing. Can be used as the first slide or duplicated for major parts.", fields: "layout: cover, title, subtitle", prompt: "Create a strong cover using title and subtitle only. Choose a title-first, audience-first, or outcome-first arrangement.", icon: Presentation },
  { id: "title-content", name: "Title and content", description: "Official PPT-style title with one main content area for explanation, notes or bullets.", fields: "layout: title-content, title, subtitle, bullets[]", prompt: "Use when the slide needs a title, then a clear explanation or list under it.", icon: FileJson },
  { id: "section-header", name: "Section header", description: "Official PPT-style section divider for a new chapter, topic or theme.", fields: "layout: section-header, title, subtitle", prompt: "Use to separate the deck into parts before continuing with content slides.", icon: Layers3 },
  { id: "two-content", name: "Two content", description: "Official PPT-style two-panel slide for two related blocks of content.", fields: "layout: two-content, columns[], bullets[]", prompt: "Use when content naturally splits into left and right groups, such as concept/details or issue/response.", icon: Columns3 },
  { id: "title-only", name: "Title only", description: "Official PPT-style large title slide for statements, transitions or simple section messages.", fields: "layout: title-only, title, subtitle", prompt: "Use for a short transition, statement, quote-like idea, or divider with very little text.", icon: Presentation },
  { id: "content-caption", name: "Content with caption", description: "Official PPT-style visual/content area with a small caption or explanation.", fields: "layout: content-caption, imagePrompt, bullets[]", prompt: "Use when a chart, screenshot, diagram or image needs a short caption or supporting note.", icon: Image },
  { id: "bullets", name: "Dot listing", description: "Clear title with short bullet points.", fields: "layout: bullets, bullets[]", prompt: "Use concise bullets with one idea per bullet.", icon: ListChecks },
  { id: "comparison", name: "Comparison", description: "Two-column or multi-option contrast.", fields: "layout: comparison, comparison[] or columns[]", prompt: "Create balanced comparison points for options, pros/cons, before/after, or alternatives.", icon: Columns3 },
  { id: "timeline", name: "Timeline", description: "Milestones, dates, phases or chronological progress.", fields: "layout: timeline, steps[]", prompt: "Turn the topic into ordered milestones with date or phase labels.", icon: ScrollText },
  { id: "process", name: "Process steps", description: "Sequential actions or workflow stages.", fields: "layout: process, steps[]", prompt: "Write 3-6 sequential steps with action-oriented labels.", icon: Rocket },
  { id: "image", name: "Image placeholder", description: "Large visual area with supporting text and image prompt.", fields: "layout: image, imagePrompt, bullets[]", prompt: "Include an imagePrompt describing the visual that should be placed on the slide.", icon: Image },
  { id: "cards", name: "2x2 cards", description: "Modular ideas, features, chapters or learning blocks.", fields: "layout: cards, cards[]", prompt: "Create four compact cards with short labels and useful details.", icon: Layers3 },
  { id: "metrics", name: "Report stats", description: "Large numbers, KPIs, findings and evidence.", fields: "layout: metrics, metrics[]", prompt: "Extract measurable facts, percentages, dates, counts or clear findings.", icon: BarChart3 },
  { id: "profile", name: "Profile/resume", description: "Skills, experience, bio or personal positioning.", fields: "layout: profile, bullets[], metrics[]", prompt: "Structure the person or team profile into highlights, skills and proof points.", icon: UserRound },
  { id: "idea-wall", name: "Idea wall", description: "Brainstorm clusters, concepts and creative directions.", fields: "layout: idea-wall, cards[]", prompt: "Generate diverse ideas grouped as short sticky-note style cards.", icon: Lightbulb },
  { id: "proposal", name: "Proposal outline", description: "Problem, solution, scope, value and next steps.", fields: "layout: proposal, bullets[] or steps[]", prompt: "Frame the section as problem, solution, impact and next move.", icon: BriefcaseBusiness },
  { id: "matrix", name: "Table matrix", description: "Rows and columns for decisions, grading or feature maps.", fields: "layout: matrix, columns[], rows[]", prompt: "Create clear column headers and row values for a matrix-style slide.", icon: Table2 },
];

const templates: Template[] = [
  { id: "business", name: "Boardroom Brief", scenario: "Business", uiStyle: "Classic", bestFor: "Executive summaries and operating plans", icon: BriefcaseBusiness, layoutBias: ["bullets", "metrics", "proposal"], colors: { bg: "101820", ink: "F8FAFC", muted: "CBD5E1", panel: "1C2A35", accent: "F2AA4C", accent2: "86C5DA", border: "334155", shadow: "05080B" } },
  { id: "education", name: "Lesson Builder", scenario: "Education", uiStyle: "Classic", bestFor: "Lessons, modules and curriculum notes", icon: GraduationCap, layoutBias: ["cards", "process", "image"], colors: { bg: "F7F3E8", ink: "20322E", muted: "5B6B64", panel: "FFFFFF", accent: "2F7D6D", accent2: "E3A535", border: "D8CEB5", shadow: "B8AD90" } },
  { id: "sales", name: "Revenue Pitch", scenario: "Sales pitch", uiStyle: "Classic", bestFor: "Commercial pitches and pipeline reviews", icon: LineChart, layoutBias: ["metrics", "comparison", "process"], colors: { bg: "0B1320", ink: "FFFFFF", muted: "B6C2D2", panel: "162235", accent: "22C55E", accent2: "38BDF8", border: "1E3A5F", shadow: "020617" } },
  { id: "proposal", name: "Project Case", scenario: "Project proposal", uiStyle: "Classic", bestFor: "Problem, scope, impact and approval decks", icon: Rocket, layoutBias: ["proposal", "timeline", "process"], colors: { bg: "F4F6F8", ink: "1E293B", muted: "64748B", panel: "FFFFFF", accent: "0F766E", accent2: "D97706", border: "CBD5E1", shadow: "94A3B8" } },
  { id: "business-skeuo", name: "Executive Desk", scenario: "Business", uiStyle: "Skeuomorphism", bestFor: "Premium reports with tactile panels", icon: BriefcaseBusiness, layoutBias: ["metrics", "matrix", "proposal"], colors: { bg: "2B261F", ink: "FFF7E8", muted: "D8C4A0", panel: "5A4632", accent: "C8A45D", accent2: "7FB7BE", border: "8B6F47", shadow: "140F0A" } },
  { id: "timeline-skeuo", name: "Field Journal", scenario: "Timeline", uiStyle: "Skeuomorphism", bestFor: "Milestones, research logs and roadmaps", icon: ScrollText, layoutBias: ["timeline", "image", "cards"], colors: { bg: "EFE3CA", ink: "2D241F", muted: "765B4B", panel: "FFF8E8", accent: "9D4E2F", accent2: "2E6B62", border: "C8A879", shadow: "A88961" } },
  { id: "sales-brutal", name: "Deal Hammer", scenario: "Sales pitch", uiStyle: "Neobrutalism", bestFor: "High-energy pitches and bold offers", icon: LineChart, layoutBias: ["comparison", "metrics", "process"], colors: { bg: "F7F255", ink: "111111", muted: "3A3A3A", panel: "FFFFFF", accent: "FF3B30", accent2: "00A6FF", border: "111111", shadow: "111111" } },
  { id: "brainstorming-brutal", name: "Sticky Riot", scenario: "Brainstorming", uiStyle: "Neobrutalism", bestFor: "Idea workshops and creative sprints", icon: Lightbulb, layoutBias: ["idea-wall", "cards", "image"], colors: { bg: "FFB6E6", ink: "111111", muted: "3D2435", panel: "FFFDF2", accent: "00D084", accent2: "FFD400", border: "111111", shadow: "111111" } },
  { id: "education-glass", name: "Crystal Lesson", scenario: "Education", uiStyle: "Glassmorphism", bestFor: "Modern learning and polished explainers", icon: GraduationCap, layoutBias: ["cards", "image", "process"], colors: { bg: "DDF7FF", ink: "12313A", muted: "456B75", panel: "F7FEFF", accent: "0EA5E9", accent2: "8B5CF6", border: "B8E8F5", shadow: "7DD3FC" } },
  { id: "report-glass", name: "Insight Frost", scenario: "Report", uiStyle: "Glassmorphism", bestFor: "Research findings and analytics summaries", icon: BarChart3, layoutBias: ["metrics", "matrix", "comparison"], colors: { bg: "16213E", ink: "F8FBFF", muted: "C6D4E8", panel: "243B63", accent: "2DD4BF", accent2: "F472B6", border: "5B7AA7", shadow: "09111F" } },
  { id: "resume-neumorph", name: "Soft Portfolio", scenario: "Resume", uiStyle: "Neumorphism", bestFor: "Personal branding and profile decks", icon: UserRound, layoutBias: ["profile", "metrics", "cards"], colors: { bg: "E9EEF5", ink: "1F2937", muted: "667085", panel: "E9EEF5", accent: "2563EB", accent2: "14B8A6", border: "D7DEE8", shadow: "AEB8C6" } },
  { id: "creative-neumorph", name: "Soft Studio", scenario: "Creative", uiStyle: "Neumorphism", bestFor: "Concept decks with gentle dimensionality", icon: Palette, layoutBias: ["image", "idea-wall", "cards"], colors: { bg: "F3EAF2", ink: "342638", muted: "75637C", panel: "F3EAF2", accent: "C026D3", accent2: "F59E0B", border: "E2D2E4", shadow: "C8B5CB" } },
  { id: "proposal-minimal", name: "Clear Case", scenario: "Project proposal", uiStyle: "Minimalism", bestFor: "Calm approval decks and strategic memos", icon: Rocket, layoutBias: ["proposal", "process", "timeline"], colors: { bg: "FAFAF8", ink: "1C1C1A", muted: "6B6B63", panel: "FFFFFF", accent: "1F7A68", accent2: "B45309", border: "DAD7CF", shadow: "C9C4B8" } },
  { id: "comparison-minimal", name: "Quiet Matrix", scenario: "Comparison", uiStyle: "Minimalism", bestFor: "Decision frameworks and option analysis", icon: Columns3, layoutBias: ["comparison", "matrix", "bullets"], colors: { bg: "F6F7F8", ink: "101828", muted: "667085", panel: "FFFFFF", accent: "344054", accent2: "0F766E", border: "D0D5DD", shadow: "EAECF0" } },
  { id: "resume", name: "Career Profile", scenario: "Resume", uiStyle: "Classic", bestFor: "Candidate, founder or team profiles", icon: UserRound, layoutBias: ["profile", "cards", "metrics"], colors: { bg: "111111", ink: "FAFAFA", muted: "D4D4D4", panel: "252525", accent: "E11D48", accent2: "F5F5F5", border: "404040", shadow: "000000" } },
  { id: "timeline", name: "Milestone Map", scenario: "Timeline", uiStyle: "Classic", bestFor: "Roadmaps and phased launches", icon: ScrollText, layoutBias: ["timeline", "process", "proposal"], colors: { bg: "FEF7ED", ink: "2D241F", muted: "765B4B", panel: "FFFFFF", accent: "C2410C", accent2: "2563EB", border: "FED7AA", shadow: "FDBA74" } },
  { id: "comparison", name: "Decision Matrix", scenario: "Comparison", uiStyle: "Classic", bestFor: "Alternatives, pros and cons, feature maps", icon: Columns3, layoutBias: ["comparison", "matrix", "cards"], colors: { bg: "F8FAFC", ink: "172033", muted: "64748B", panel: "FFFFFF", accent: "4F46E5", accent2: "14B8A6", border: "CBD5E1", shadow: "94A3B8" } },
  { id: "brainstorming", name: "Idea Wall", scenario: "Brainstorming", uiStyle: "Classic", bestFor: "Workshops and concept exploration", icon: Lightbulb, layoutBias: ["idea-wall", "cards", "image"], colors: { bg: "FFF8DB", ink: "302A12", muted: "766B35", panel: "FFFFFF", accent: "F59E0B", accent2: "DB2777", border: "FDE68A", shadow: "FCD34D" } },
  { id: "report", name: "Insight Report", scenario: "Report", uiStyle: "Classic", bestFor: "KPIs, findings and recommendations", icon: BarChart3, layoutBias: ["metrics", "matrix", "bullets"], colors: { bg: "F1F5F9", ink: "0F172A", muted: "475569", panel: "FFFFFF", accent: "0369A1", accent2: "65A30D", border: "CBD5E1", shadow: "94A3B8" } },
  { id: "creative", name: "Studio Concept", scenario: "Creative", uiStyle: "Classic", bestFor: "Campaign ideas and visual concepts", icon: Palette, layoutBias: ["image", "idea-wall", "cards"], colors: { bg: "201A2E", ink: "FFF7ED", muted: "E9D5FF", panel: "382B4A", accent: "FB7185", accent2: "FBBF24", border: "6B4B84", shadow: "120D1A" } },
];

const starterJson = JSON.stringify(
  {
    title: "Presentation title",
    subtitle: "Short subtitle for the cover page",
    slides: [
      {
        layout: "bullets",
        title: "Overview",
        subtitle: "Optional short context",
        bullets: ["Short point 1", "Short point 2", "Short point 3"],
      },
      {
        layout: "metrics",
        title: "Important numbers",
        metrics: ["42%: result or KPI", "12 weeks: timeline", "$1.2M: impact"],
      },
      {
        layout: "comparison",
        title: "Options compared",
        comparison: ["Option A: strength", "Option A: risk", "Option B: strength", "Option B: risk"],
      },
      {
        layout: "image",
        title: "Visual slide",
        imagePrompt: "Describe the image, chart, diagram, screenshot, or scene needed here",
        bullets: ["Caption point", "Why this visual matters"],
      },
    ],
  },
  null,
  2,
);

const sampleJson = starterJson;

const safeText = (value: unknown, fallback = "Untitled") => {
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    const text = String(value).trim();
    return text.length ? text.slice(0, 240) : fallback;
  }
  return fallback;
};

const valuesToBullets = (value: unknown): string[] => {
  if (Array.isArray(value)) return value.flatMap(valuesToBullets).slice(0, 10);
  if (value && typeof value === "object") {
    return Object.entries(value)
      .map(([key, item]) => `${key}: ${safeText(item, JSON.stringify(item).slice(0, 140))}`)
      .slice(0, 10);
  }
  return [safeText(value, "Detail")];
};

const parseLayout = (value: unknown): LayoutType | undefined => {
  const text = safeText(value, "").toLowerCase().replace(/_/g, "-");
  return layoutOutlines.some((outline) => outline.id === text) ? (text as LayoutType) : undefined;
};

const normalizeDeck = (jsonText: string): DeckContent => {
  const parsed = JSON.parse(jsonText) as Record<string, unknown> | unknown[];
  const root = Array.isArray(parsed) ? { title: "Generated presentation", slides: parsed } : parsed;
  const rawSlides = [root.slides, root.sections, root.items, root.content].find(Array.isArray) as unknown[] | undefined;

  const entries = rawSlides?.length ? rawSlides : Object.entries(root).filter(([, value]) => value && typeof value === "object");
  const sections = entries.slice(0, 16).map((entry, index): SlideContent => {
    const value = Array.isArray(rawSlides) ? entry : (entry as [string, unknown])[1];
    const key = Array.isArray(rawSlides) ? `Section ${index + 1}` : safeText((entry as [string, unknown])[0], `Section ${index + 1}`);

    if (value && typeof value === "object" && !Array.isArray(value)) {
      const record = value as Record<string, unknown>;
      const title = safeText(record.title ?? record.heading ?? record.name ?? key, key);
      const bulletsSource = record.bullets ?? record.points ?? record.items ?? record.details ?? record.description ?? record.summary ?? record.takeaways ?? record;
      return {
        title,
        subtitle: safeText(record.subtitle ?? record.kicker ?? record.context ?? "", ""),
        bullets: valuesToBullets(bulletsSource).filter((bullet) => bullet !== title).slice(0, 8),
        notes: safeText(record.notes ?? "", ""),
        layout: parseLayout(record.layout ?? record.type),
        image: safeText(record.image ?? record.imageUrl ?? "", ""),
        imagePrompt: safeText(record.imagePrompt ?? record.visual ?? record.visualPrompt ?? "", ""),
        comparison: valuesToBullets(record.comparison ?? record.compare ?? record.options ?? []).slice(0, 6),
        steps: valuesToBullets(record.steps ?? record.timeline ?? record.milestones ?? record.process ?? []).slice(0, 6),
        metrics: valuesToBullets(record.metrics ?? record.stats ?? record.kpis ?? []).slice(0, 6),
        cards: valuesToBullets(record.cards ?? record.ideas ?? record.features ?? []).slice(0, 8),
        columns: valuesToBullets(record.columns ?? record.headers ?? []).slice(0, 5),
        rows: valuesToBullets(record.rows ?? record.table ?? record.matrix ?? []).slice(0, 8),
      };
    }

    return { title: key, bullets: valuesToBullets(value).slice(0, 8) };
  });

  if (!sections.length) {
    sections.push({ title: "Overview", layout: "bullets", bullets: Object.entries(root).map(([key, value]) => `${key}: ${safeText(value)}`).slice(0, 8) });
  }

  return {
    title: safeText(root.title ?? root.name ?? root.topic, "Generated presentation"),
    subtitle: safeText(root.subtitle ?? root.description ?? root.summary, "Built from your JSON content"),
    sections,
  };
};

const resolveLayout = (item: SlideContent, template: Template, index: number): LayoutType => {
  if (item.layout) return item.layout;
  if (item.metrics?.length) return "metrics";
  if (item.comparison?.length || item.columns?.length) return "comparison";
  if (item.steps?.length) return template.layoutBias.includes("timeline") ? "timeline" : "process";
  if (item.image || item.imagePrompt) return "image";
  if (item.cards?.length) return template.layoutBias.includes("idea-wall") ? "idea-wall" : "cards";
  return template.layoutBias[index % template.layoutBias.length] === "cover" ? "bullets" : template.layoutBias[index % template.layoutBias.length];
};

const styleSettings = (template: Template) => {
  switch (template.uiStyle) {
    case "Neobrutalism":
      return { radius: 0.02, lineWidth: 2.25, shadowOffset: 0.09, panelTransparency: 0 };
    case "Glassmorphism":
      return { radius: 0.12, lineWidth: 1, shadowOffset: 0.04, panelTransparency: 18 };
    case "Neumorphism":
      return { radius: 0.16, lineWidth: 0.7, shadowOffset: 0.06, panelTransparency: 0 };
    case "Minimalism":
      return { radius: 0.02, lineWidth: 0.55, shadowOffset: 0.02, panelTransparency: 0 };
    case "Skeuomorphism":
      return { radius: 0.1, lineWidth: 1.2, shadowOffset: 0.07, panelTransparency: 0 };
    default:
      return { radius: 0.06, lineWidth: 0.8, shadowOffset: 0.04, panelTransparency: 0 };
  }
};

const addPanel = (slide: pptxgen.Slide, template: Template, x: number, y: number, w: number, h: number, fill = template.colors.panel) => {
  const style = styleSettings(template);
  if (style.shadowOffset > 0.03) {
    slide.addShape("roundRect", { x: x + style.shadowOffset, y: y + style.shadowOffset, w, h, rectRadius: style.radius, fill: { color: template.colors.shadow, transparency: template.uiStyle === "Neobrutalism" ? 0 : 55 }, line: { color: template.colors.shadow, transparency: 100 } });
  }
  slide.addShape("roundRect", { x, y, w, h, rectRadius: style.radius, fill: { color: fill, transparency: style.panelTransparency }, line: { color: template.colors.border, width: style.lineWidth } });
  if (template.uiStyle === "Skeuomorphism") {
    slide.addShape("line", { x: x + 0.12, y: y + 0.14, w: w - 0.24, h: 0, line: { color: "FFFFFF", width: 1, transparency: 35 } });
  }
};

const addFooter = (slide: pptxgen.Slide, template: Template, index: number, layout: LayoutType) => {
  slide.addText(layout.toUpperCase(), { x: 0.6, y: 6.83, w: 2.6, h: 0.22, fontFace: "Aptos", fontSize: 7.5, bold: true, color: template.colors.muted, margin: 0 });
  slide.addShape("line", { x: 3.25, y: 6.93, w: 8.6, h: 0, line: { color: template.colors.border, width: 0.8 } });
  slide.addText(String(index).padStart(2, "0"), { x: 12.05, y: 6.78, w: 0.75, h: 0.22, fontFace: "Aptos", fontSize: 8, color: template.colors.muted, align: "right", margin: 0 });
};

const addTitle = (slide: pptxgen.Slide, item: SlideContent, template: Template) => {
  slide.addText(item.title, { x: 0.62, y: 0.46, w: 8.8, h: 0.56, fontFace: "Aptos Display", fontSize: 27, bold: true, color: template.colors.ink, fit: "shrink", margin: 0 });
  if (item.subtitle) slide.addText(item.subtitle, { x: 0.64, y: 1.06, w: 8.3, h: 0.28, fontFace: "Aptos", fontSize: 10, color: template.colors.muted, fit: "shrink", margin: 0 });
};

const addTitleSlide = (pptx: pptxgen, deck: DeckContent, template: Template) => {
  const slide = pptx.addSlide();
  slide.background = { color: template.colors.bg };
  if (template.uiStyle === "Neobrutalism") {
    slide.addShape("rect", { x: 0.25, y: 0.28, w: 12.2, h: 6.5, fill: { color: template.colors.panel }, line: { color: template.colors.border, width: 3 } });
    slide.addShape("rect", { x: 0.65, y: 0.72, w: 2.35, h: 0.55, fill: { color: template.colors.accent2 }, line: { color: template.colors.border, width: 2 } });
  } else {
    slide.addShape("rect", { x: 0, y: 0, w: 3.1, h: 7.5, fill: { color: template.colors.accent }, line: { color: template.colors.accent } });
    slide.addShape("rect", { x: 0.42, y: 0.58, w: 2.2, h: 0.12, fill: { color: template.colors.accent2 }, line: { color: template.colors.accent2 } });
  }
  slide.addText(`${template.scenario.toUpperCase()} · ${template.uiStyle.toUpperCase()}`, { x: 3.65, y: 1.08, w: 8.2, h: 0.32, fontFace: "Aptos", fontSize: 11, bold: true, color: template.colors.accent, margin: 0 });
  slide.addText(deck.title, { x: 3.6, y: 1.66, w: 8.45, h: 1.5, fontFace: "Aptos Display", fontSize: 41, bold: true, color: template.colors.ink, fit: "shrink", margin: 0 });
  slide.addText(deck.subtitle, { x: 3.65, y: 3.48, w: 7.3, h: 0.65, fontFace: "Aptos", fontSize: 15.5, color: template.colors.muted, fit: "shrink", margin: 0 });
  addPanel(slide, template, 3.65, 5.38, 3.25, 0.62, template.colors.panel);
  slide.addText(`${deck.sections.length} content slides`, { x: 3.9, y: 5.58, w: 2.7, h: 0.22, fontFace: "Aptos", fontSize: 12.5, bold: true, color: template.colors.ink, margin: 0 });
  slide.addShape("arc", { x: 10.9, y: 5.12, w: 1.65, h: 1.65, line: { color: template.colors.accent2, width: 5, transparency: 10 } });
};

const getPrimaryItems = (item: SlideContent, layout: LayoutType) => {
  if (layout === "metrics") return item.metrics?.length ? item.metrics : item.bullets;
  if (layout === "comparison") return item.comparison?.length ? item.comparison : item.columns?.length ? item.columns : item.bullets;
  if (layout === "timeline" || layout === "process") return item.steps?.length ? item.steps : item.bullets;
  if (layout === "cards" || layout === "idea-wall") return item.cards?.length ? item.cards : item.bullets;
  if (layout === "matrix") return item.rows?.length ? item.rows : item.bullets;
  return item.bullets;
};

const addBulletSlide = (slide: pptxgen.Slide, item: SlideContent, template: Template) => {
  addTitle(slide, item, template);
  getPrimaryItems(item, "bullets").slice(0, 7).forEach((bullet, i) => {
    const y = 1.58 + i * 0.68;
    slide.addShape("ellipse", { x: 0.78, y: y + 0.07, w: 0.22, h: 0.22, fill: { color: i % 2 ? template.colors.accent2 : template.colors.accent }, line: { color: i % 2 ? template.colors.accent2 : template.colors.accent } });
    slide.addText(bullet, { x: 1.18, y, w: 9.55, h: 0.34, fontFace: "Aptos", fontSize: 15, color: template.colors.ink, fit: "shrink", margin: 0.02 });
  });
  addPanel(slide, template, 10.3, 1.28, 1.85, 4.35, template.colors.accent);
  slide.addText("Outline", { x: 10.68, y: 3.12, w: 1.1, h: 0.28, rotate: 90, fontFace: "Aptos", fontSize: 14, bold: true, color: template.colors.bg, margin: 0 });
};

const addComparisonSlide = (slide: pptxgen.Slide, item: SlideContent, template: Template) => {
  addTitle(slide, item, template);
  const items = getPrimaryItems(item, "comparison").slice(0, 6);
  const half = Math.ceil(items.length / 2);
  [items.slice(0, half), items.slice(half)].forEach((group, col) => {
    const x = 0.72 + col * 6.08;
    addPanel(slide, template, x, 1.55, 5.45, 4.85, col ? template.colors.panel : template.colors.accent);
    slide.addText(col ? "Option B" : "Option A", { x: x + 0.28, y: 1.82, w: 2.2, h: 0.26, fontFace: "Aptos", fontSize: 12, bold: true, color: col ? template.colors.ink : template.colors.bg, margin: 0 });
    group.forEach((text, i) => slide.addText(text, { x: x + 0.3, y: 2.35 + i * 0.82, w: 4.8, h: 0.42, fontFace: "Aptos", fontSize: 13.5, color: col ? template.colors.ink : template.colors.bg, fit: "shrink", margin: 0.02 }));
  });
};

const addTimelineSlide = (slide: pptxgen.Slide, item: SlideContent, template: Template) => {
  addTitle(slide, item, template);
  const items = getPrimaryItems(item, "timeline").slice(0, 5);
  slide.addShape("line", { x: 1.15, y: 3.55, w: 10.65, h: 0, line: { color: template.colors.accent, width: 2.5 } });
  items.forEach((text, i) => {
    const x = 0.75 + i * (10.7 / Math.max(items.length - 1, 1));
    slide.addShape("ellipse", { x, y: 3.28, w: 0.55, h: 0.55, fill: { color: i % 2 ? template.colors.accent2 : template.colors.accent }, line: { color: template.colors.border, width: 1 } });
    slide.addText(String(i + 1), { x: x + 0.18, y: 3.43, w: 0.2, h: 0.14, fontFace: "Aptos", fontSize: 8, bold: true, color: template.colors.bg, align: "center", margin: 0 });
    addPanel(slide, template, Math.min(x - 0.42, 10.85), i % 2 ? 4.25 : 1.72, 1.85, 0.95);
    slide.addText(text, { x: Math.min(x - 0.24, 11.02), y: i % 2 ? 4.48 : 1.95, w: 1.48, h: 0.42, fontFace: "Aptos", fontSize: 9.5, color: template.colors.ink, fit: "shrink", margin: 0.02 });
  });
};

const addProcessSlide = (slide: pptxgen.Slide, item: SlideContent, template: Template) => {
  addTitle(slide, item, template);
  const items = getPrimaryItems(item, "process").slice(0, 6);
  items.forEach((text, i) => {
    const x = 0.72 + (i % 3) * 4.05;
    const y = 1.62 + Math.floor(i / 3) * 2.25;
    addPanel(slide, template, x, y, 3.45, 1.45);
    slide.addText(String(i + 1).padStart(2, "0"), { x: x + 0.22, y: y + 0.18, w: 0.6, h: 0.28, fontFace: "Aptos Display", fontSize: 16, bold: true, color: template.colors.accent, margin: 0 });
    slide.addText(text, { x: x + 0.24, y: y + 0.66, w: 2.92, h: 0.52, fontFace: "Aptos", fontSize: 12.5, color: template.colors.ink, fit: "shrink", margin: 0.02 });
  });
};

const addImageSlide = (slide: pptxgen.Slide, item: SlideContent, template: Template) => {
  addTitle(slide, item, template);
  addPanel(slide, template, 0.72, 1.52, 6.25, 4.75, template.colors.panel);
  slide.addShape("rect", { x: 1.03, y: 1.86, w: 5.64, h: 3.42, fill: { color: template.colors.accent, transparency: 12 }, line: { color: template.colors.border, transparency: 20 } });
  slide.addText("IMAGE PLACEHOLDER", { x: 1.65, y: 3.12, w: 4.4, h: 0.3, fontFace: "Aptos", fontSize: 15, bold: true, color: template.colors.bg, align: "center", margin: 0 });
  slide.addText(item.imagePrompt || item.image || "Add image prompt or source URL in the JSON.", { x: 1.2, y: 5.55, w: 5.28, h: 0.32, fontFace: "Aptos", fontSize: 10, color: template.colors.muted, fit: "shrink", margin: 0 });
  getPrimaryItems(item, "bullets").slice(0, 4).forEach((bullet, i) => slide.addText(bullet, { x: 7.35, y: 1.72 + i * 0.9, w: 4.3, h: 0.42, fontFace: "Aptos", fontSize: 14, color: template.colors.ink, fit: "shrink", margin: 0.02 }));
};

const addCardsSlide = (slide: pptxgen.Slide, item: SlideContent, template: Template, layout: LayoutType) => {
  addTitle(slide, item, template);
  const items = getPrimaryItems(item, layout).slice(0, 8);
  items.forEach((text, i) => {
    const x = 0.72 + (i % 4) * 3.05;
    const y = 1.62 + Math.floor(i / 4) * 2.15;
    addPanel(slide, template, x, y, 2.58, 1.55, layout === "idea-wall" && i % 2 ? template.colors.accent2 : template.colors.panel);
    slide.addText(text, { x: x + 0.2, y: y + 0.28, w: 2.12, h: 0.72, fontFace: "Aptos", fontSize: 12, bold: layout === "idea-wall", color: layout === "idea-wall" && i % 2 ? template.colors.bg : template.colors.ink, fit: "shrink", margin: 0.02 });
  });
};

const addMetricsSlide = (slide: pptxgen.Slide, item: SlideContent, template: Template) => {
  addTitle(slide, item, template);
  const items = getPrimaryItems(item, "metrics").slice(0, 6);
  items.forEach((text, i) => {
    const x = 0.72 + (i % 3) * 4.05;
    const y = 1.62 + Math.floor(i / 3) * 2.0;
    addPanel(slide, template, x, y, 3.45, 1.36, i % 2 ? template.colors.panel : template.colors.accent);
    const [value, ...label] = text.split(/[:–-]/);
    slide.addText(value.trim(), { x: x + 0.2, y: y + 0.26, w: 3, h: 0.38, fontFace: "Aptos Display", fontSize: 22, bold: true, color: i % 2 ? template.colors.accent : template.colors.bg, fit: "shrink", margin: 0 });
    slide.addText((label.join("-") || text).trim(), { x: x + 0.23, y: y + 0.82, w: 2.92, h: 0.22, fontFace: "Aptos", fontSize: 9.5, color: i % 2 ? template.colors.muted : template.colors.bg, fit: "shrink", margin: 0 });
  });
};

const addMatrixSlide = (slide: pptxgen.Slide, item: SlideContent, template: Template) => {
  addTitle(slide, item, template);
  const headers = item.columns?.length ? item.columns.slice(0, 4) : ["Criteria", "Option", "Impact", "Notes"];
  const rows = getPrimaryItems(item, "matrix").slice(0, 5);
  const x = 0.72;
  const y = 1.58;
  const colW = 11.45 / headers.length;
  headers.forEach((head, i) => {
    slide.addShape("rect", { x: x + i * colW, y, w: colW, h: 0.52, fill: { color: template.colors.accent }, line: { color: template.colors.border, width: 1 } });
    slide.addText(head, { x: x + i * colW + 0.1, y: y + 0.17, w: colW - 0.2, h: 0.16, fontFace: "Aptos", fontSize: 8.5, bold: true, color: template.colors.bg, fit: "shrink", margin: 0 });
  });
  rows.forEach((row, r) => {
    headers.forEach((_, c) => {
      slide.addShape("rect", { x: x + c * colW, y: y + 0.52 + r * 0.72, w: colW, h: 0.72, fill: { color: template.colors.panel }, line: { color: template.colors.border, width: 0.75 } });
      slide.addText(c === 0 ? row : "", { x: x + c * colW + 0.1, y: y + 0.75 + r * 0.72, w: colW - 0.2, h: 0.22, fontFace: "Aptos", fontSize: 8.5, color: template.colors.ink, fit: "shrink", margin: 0 });
    });
  });
};

const addContentSlide = (pptx: pptxgen, item: SlideContent, template: Template, index: number) => {
  const slide = pptx.addSlide();
  const layout = resolveLayout(item, template, index);
  slide.background = { color: template.colors.bg };
  if (template.uiStyle === "Glassmorphism") {
    slide.addShape("ellipse", { x: 9.8, y: -0.6, w: 2.7, h: 2.7, fill: { color: template.colors.accent2, transparency: 45 }, line: { color: template.colors.accent2, transparency: 100 } });
    slide.addShape("ellipse", { x: -0.8, y: 5.2, w: 2.5, h: 2.5, fill: { color: template.colors.accent, transparency: 55 }, line: { color: template.colors.accent, transparency: 100 } });
  }

  if (layout === "comparison" || layout === "proposal") addComparisonSlide(slide, item, template);
  else if (layout === "timeline") addTimelineSlide(slide, item, template);
  else if (layout === "process") addProcessSlide(slide, item, template);
  else if (layout === "image") addImageSlide(slide, item, template);
  else if (layout === "cards" || layout === "idea-wall" || layout === "profile") addCardsSlide(slide, item, template, layout === "profile" ? "cards" : layout);
  else if (layout === "metrics") addMetricsSlide(slide, item, template);
  else if (layout === "matrix") addMatrixSlide(slide, item, template);
  else addBulletSlide(slide, item, template);

  addFooter(slide, template, index, layout);
};

const createPresentation = async (deck: DeckContent, template: Template) => {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "JSON PPT Generator";
  pptx.subject = `${template.scenario} · ${template.uiStyle}`;
  pptx.title = deck.title;
  pptx.company = "Lovable";
  pptx.theme = { headFontFace: "Aptos Display", bodyFontFace: "Aptos" };

  addTitleSlide(pptx, deck, template);
  deck.sections.forEach((section, index) => addContentSlide(pptx, section, template, index + 1));

  const fileName = `${deck.title.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "presentation"}-${template.id}.pptx`;
  await pptx.writeFile({ fileName });
};

const buildPrompt = () => `I need you to create JSON for a PowerPoint deck. The JSON will be pasted into a JSON-to-PPT web app, so please follow this guide naturally and carefully.

Topic: [replace with topic]
Audience: [replace with audience]
Goal: [replace with presentation goal]
Optional source materials: [paste notes, article text, curriculum, resume, data, transcript, research, or image requirements here]

Your final answer must be JSON only. Do not write an introduction, explanation, markdown, code fence, comments, or anything outside the JSON.

Think of the deck like this:
1. The whole presentation is one JSON object.
2. The object has a main "title", a short "subtitle", and a "slides" array.
3. Each item inside "slides" is one PowerPoint slide.
4. Each slide chooses one layout, then uses only the fields that make sense for that layout.

The required overall structure is:
{
  "title": "Presentation title",
  "subtitle": "Short subtitle for the cover page",
  "slides": [
    {
      "layout": "bullets",
      "title": "Slide title",
      "subtitle": "Optional short context",
      "bullets": ["Short point", "Short point"],
      "imagePrompt": "Optional image description"
    }
  ]
}

Natural writing rules for this JSON:
- Use clear, human slide titles. A title should sound like a slide headline, not a long paragraph.
- Keep all text short because it will be placed inside PowerPoint slides.
- Use arrays of strings, not nested objects, for slide content.
- Do not use long paragraphs inside bullets, cards, metrics, rows, or comparison items.
- Do not invent facts if source materials are provided. Summarize and organize what is given.
- If a slide needs a picture, screenshot, chart, diagram, person, place, product, or scene, write an "imagePrompt" that describes what should be shown.
- If no real image URL is provided, do not fake an image URL. Use "imagePrompt" instead.
- It is okay to repeat layouts when they fit the topic.
- It is okay to skip layouts that do not fit the topic.
- It is not required to use every layout.
- You may combine helpful fields, for example an "image" slide can include both "imagePrompt" and "bullets".

Supported fields:
- Top-level fields: "title", "subtitle", "slides".
- Common slide fields: "layout", "title", "subtitle", "notes".
- List fields: "bullets", "points", "items", "details", "takeaways".
- Visual fields: "imagePrompt", "visual", "visualPrompt", "image", "imageUrl".
- Structured fields: "comparison", "compare", "options", "steps", "timeline", "milestones", "process", "metrics", "stats", "kpis", "cards", "ideas", "features", "columns", "headers", "rows", "table", "matrix".

Use the standard field names when possible:
- Use "bullets" for normal lists.
- Use "steps" for process, timeline, or ordered sequence slides.
- Use "comparison" for before/after, pros/cons, or option comparisons.
- Use "metrics" for KPIs, numbers, findings, results, dates, or proof points.
- Use "cards" for features, modules, ideas, chapters, or grouped concepts.
- Use "columns" and "rows" for matrix/table-style slides.
- Use "imagePrompt" for any visual placeholder.

Supported layouts and how to use them:
${layoutOutlines.map((layout) => `
Layout value: "${layout.id}"
Layout name: ${layout.name}
Use this when: ${layout.description}
Good outline types: ${layoutVariants[layout.id].join("; ")}
Best fields to use: ${layout.fields}
Natural instruction: ${layout.prompt}
JSON example:
${layoutJsonExamples[layout.id]}`).join("\n")}

Here is a complete example of good JSON:
{
  "title": "Market Expansion Plan",
  "subtitle": "Opportunities, risks, and launch priorities",
  "slides": [
    {
      "layout": "metrics",
      "title": "Current position",
      "metrics": ["18%: revenue growth", "3.4x: pipeline coverage", "6 pts: retention lift"]
    },
    {
      "layout": "comparison",
      "title": "Expansion options",
      "comparison": ["Enterprise: higher ACV", "Enterprise: longer sales cycle", "SMB: faster adoption", "SMB: higher churn risk"]
    },
    {
      "layout": "timeline",
      "title": "Launch roadmap",
      "steps": ["Month 1: validate segment", "Month 2: build campaign", "Month 3: pilot", "Month 4: scale"]
    },
    {
      "layout": "image",
      "title": "Customer workflow",
      "imagePrompt": "Diagram showing the customer journey from signup to activation",
      "bullets": ["Show friction points", "Highlight activation moment"]
    },
    {
      "layout": "proposal",
      "title": "Recommended next move",
      "bullets": ["Start with enterprise pilot", "Use current case studies", "Review results after 45 days"]
    }
  ]
}

Before returning the JSON, check it against this final checklist:
- Is the answer valid JSON only?
- Does it start with one object using "title", "subtitle", and "slides"?
- Is "slides" an array?
- Is every slide an object?
- Does every slide have "layout" and "title"?
- Are content lists arrays of short strings?
- Are layouts chosen naturally based on the material?
- Is any needed visual described with "imagePrompt"?
- Are there no markdown fences, comments, trailing commas, or extra text?`;

const MiniLayoutPreview = ({ outline, template }: { outline: Outline; template: Template }) => {
  const panelStyle = { borderColor: `#${template.colors.border}`, backgroundColor: `#${template.colors.panel}` };
  const accentStyle = { backgroundColor: `#${template.colors.accent}` };
  const accentAltStyle = { backgroundColor: `#${template.colors.accent2}` };

  return (
    <div className="mt-4 border p-2" style={{ borderColor: `#${template.colors.border}`, backgroundColor: `#${template.colors.bg}` }}>
      <div className="mb-2 h-1.5 w-2/5" style={accentStyle} />
      {outline.id === "comparison" ? <div className="grid grid-cols-2 gap-2"><div className="h-14 border" style={panelStyle} /><div className="h-14 border" style={panelStyle} /></div> : null}
      {outline.id === "timeline" || outline.id === "process" ? <div className="flex items-center gap-1 py-4">{[0, 1, 2, 3].map((item) => <div key={item} className="flex flex-1 items-center gap-1"><span className="h-4 w-4 rounded-full" style={item % 2 ? accentAltStyle : accentStyle} />{item < 3 ? <span className="h-0.5 flex-1" style={panelStyle} /> : null}</div>)}</div> : null}
      {outline.id === "image" ? <div className="grid grid-cols-[1.5fr_1fr] gap-2"><div className="flex h-16 items-center justify-center border text-[9px] font-black uppercase" style={accentStyle}>Image</div><div className="space-y-1"><div className="h-3 border" style={panelStyle} /><div className="h-3 border" style={panelStyle} /><div className="h-3 border" style={panelStyle} /></div></div> : null}
      {outline.id === "cards" || outline.id === "idea-wall" || outline.id === "profile" ? <div className="grid grid-cols-2 gap-2">{[0, 1, 2, 3].map((item) => <div key={item} className="h-8 border" style={item % 2 && outline.id === "idea-wall" ? accentAltStyle : panelStyle} />)}</div> : null}
      {outline.id === "metrics" ? <div className="grid grid-cols-3 gap-2">{[0, 1, 2].map((item) => <div key={item} className="h-12 border" style={item % 2 ? panelStyle : accentStyle} />)}</div> : null}
      {outline.id === "matrix" ? <div className="space-y-1">{[0, 1, 2].map((row) => <div key={row} className="grid grid-cols-4 gap-1">{[0, 1, 2, 3].map((col) => <div key={col} className="h-4 border" style={row === 0 ? accentStyle : panelStyle} />)}</div>)}</div> : null}
      {outline.id === "cover" || outline.id === "bullets" || outline.id === "proposal" ? <div className="space-y-2 py-2"><div className="h-5 w-3/4 border" style={panelStyle} /><div className="h-2 w-full" style={accentStyle} /><div className="h-2 w-4/5" style={accentAltStyle} /><div className="h-2 w-2/3" style={panelStyle} /></div> : null}
    </div>
  );
};

const Index = () => {
  const [jsonText, setJsonText] = useState(sampleJson);
  const [selectedTemplate, setSelectedTemplate] = useState<TemplateId>("business");
  const [selectedStyle, setSelectedStyle] = useState<UiStyle | "All">("All");
  const [selectedOutline, setSelectedOutline] = useState<LayoutType>("image");
  const [error, setError] = useState("");
  const [copyStatus, setCopyStatus] = useState("");
  const [isGenerating, setIsGenerating] = useState(false);

  const activeTemplate = templates.find((template) => template.id === selectedTemplate) ?? templates[0];
  const aiPrompt = useMemo(() => buildPrompt(), []);
  const filteredTemplates = selectedStyle === "All" ? templates : templates.filter((template) => template.uiStyle === selectedStyle);
  const deck = useMemo(() => {
    try {
      return normalizeDeck(jsonText);
    } catch {
      return null;
    }
  }, [jsonText]);

  const handleUpload = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!file.name.endsWith(".json") && file.type !== "application/json") {
      setError("Please upload a valid .json file.");
      return;
    }
    setJsonText(await file.text());
    setError("");
  };

  const handleGenerate = async () => {
    try {
      setError("");
      setIsGenerating(true);
      const normalized = normalizeDeck(jsonText);
      await createPresentation(normalized, activeTemplate);
    } catch (caught) {
      console.error("PPT generation failed", caught);
      setError(caught instanceof SyntaxError ? "The JSON could not be parsed. Check quotes, commas, and brackets." : caught instanceof Error ? caught.message : "Could not generate the presentation from this JSON.");
    } finally {
      setIsGenerating(false);
    }
  };

  const copyPrompt = async () => {
    await navigator.clipboard.writeText(aiPrompt);
    setCopyStatus("Prompt copied");
    window.setTimeout(() => setCopyStatus(""), 1800);
  };

  const useStarterJson = () => {
    setJsonText(starterJson);
    setError("");
  };

  return (
    <main className="min-h-screen bg-background text-foreground">
      <section className="border-b bg-[radial-gradient(circle_at_top_left,hsl(var(--accent)),transparent_28%),linear-gradient(135deg,hsl(var(--background)),hsl(var(--muted)))]">
        <div className="mx-auto grid max-w-7xl gap-8 px-5 py-8 lg:grid-cols-[0.9fr_1.1fr] lg:px-8 lg:py-10">
          <div className="flex min-h-[360px] flex-col justify-between">
            <div>
              <div className="mb-5 inline-flex items-center gap-2 border border-border bg-card px-3 py-2 text-xs font-semibold uppercase tracking-widest text-muted-foreground">
                <Presentation className="h-4 w-4 text-primary" /> JSON to PowerPoint
              </div>
              <h1 className="max-w-2xl text-4xl font-black leading-tight tracking-normal md:text-6xl">Generate adaptable slide decks from structured JSON.</h1>
              <p className="mt-5 max-w-xl text-base leading-7 text-muted-foreground md:text-lg">
                Upload or paste JSON, choose from 20 templates across six UI styles, map content into layout outlines, then download a finished `.pptx` file.
              </p>
            </div>
            <div className="mt-8 grid max-w-xl grid-cols-3 border border-border bg-card">
              <div className="border-r border-border p-4"><p className="text-2xl font-black">20</p><p className="text-xs text-muted-foreground">templates</p></div>
              <div className="border-r border-border p-4"><p className="text-2xl font-black">12</p><p className="text-xs text-muted-foreground">layout outlines</p></div>
              <div className="p-4"><p className="text-2xl font-black">.pptx</p><p className="text-xs text-muted-foreground">download</p></div>
            </div>
          </div>

          <div className="border border-border bg-card p-4 shadow-2xl shadow-primary/10">
            <div className="mb-3 flex items-center justify-between gap-3">
              <div>
                <p className="text-sm font-bold">JSON input</p>
                <p className="text-xs text-muted-foreground">Supports layouts, image prompts, metrics, comparisons, steps, cards and matrix rows.</p>
              </div>
              <div className="flex flex-wrap gap-2">
                <button type="button" onClick={useStarterJson} className="inline-flex items-center gap-2 border border-border bg-secondary px-3 py-2 text-sm font-semibold text-secondary-foreground transition-colors hover:bg-accent"><FileJson className="h-4 w-4" /> Starter</button>
                <label className="inline-flex cursor-pointer items-center gap-2 border border-border bg-secondary px-3 py-2 text-sm font-semibold text-secondary-foreground transition-colors hover:bg-accent">
                  <Upload className="h-4 w-4" /> Upload
                  <input className="sr-only" type="file" accept="application/json,.json" onChange={handleUpload} />
                </label>
              </div>
            </div>
            <div className="mb-3 grid gap-3 border border-border bg-muted/50 p-3 text-xs leading-5 text-muted-foreground md:grid-cols-[0.85fr_1.15fr]">
              <div>
                <p className="font-black text-foreground">How JSON should look</p>
                <p className="mt-1">Use one object with <span className="font-mono">title</span>, <span className="font-mono">subtitle</span>, and a <span className="font-mono">slides</span> array. Every slide needs a <span className="font-mono">layout</span> and <span className="font-mono">title</span>.</p>
              </div>
              <pre className="overflow-auto border border-border bg-background p-2 font-mono text-[11px] text-foreground">{`{
  "title": "Presentation title",
  "subtitle": "Short cover subtitle",
  "slides": [
    { "layout": "bullets", "title": "Overview", "bullets": ["Point 1", "Point 2"] }
  ]
}`}</pre>
            </div>
            <textarea
              value={jsonText}
              onChange={(event) => {
                setJsonText(event.target.value);
                setError("");
              }}
              spellCheck={false}
              className="min-h-[360px] w-full resize-y border border-input bg-background p-4 font-mono text-sm leading-6 text-foreground outline-none ring-offset-background transition-shadow focus:ring-2 focus:ring-ring focus:ring-offset-2"
            />
            {error ? <p className="mt-3 text-sm font-semibold text-destructive">{error}</p> : null}
          </div>
        </div>
      </section>

      <section className="mx-auto grid max-w-7xl gap-6 px-5 py-8 lg:grid-cols-[1fr_360px] lg:px-8">
        <div>
          <div className="mb-4 flex flex-wrap items-end justify-between gap-4">
            <div>
              <h2 className="text-2xl font-black tracking-normal">Choose a template UI style</h2>
              <p className="text-sm text-muted-foreground">Every style has two or more templates with different colors, shapes and slide behavior.</p>
            </div>
            <div className="flex flex-wrap gap-2">
              {["All", ...uiStyles].map((style) => (
                <button key={style} type="button" onClick={() => setSelectedStyle(style as UiStyle | "All")} className={`border px-3 py-2 text-xs font-bold transition-colors ${selectedStyle === style ? "border-primary bg-primary text-primary-foreground" : "border-border bg-card hover:bg-accent"}`}>{style}</button>
              ))}
            </div>
          </div>
          <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
            {filteredTemplates.map((template) => {
              const Icon = template.icon;
              const isSelected = template.id === selectedTemplate;
              return (
                <button key={template.id} type="button" onClick={() => setSelectedTemplate(template.id)} className={`group border p-4 text-left transition-all ${isSelected ? "border-primary bg-primary text-primary-foreground shadow-lg shadow-primary/20" : "border-border bg-card hover:-translate-y-0.5 hover:bg-accent"}`}>
                  <div className="mb-4 flex items-center justify-between">
                    <Icon className="h-5 w-5" />
                    {isSelected ? <Check className="h-4 w-4" /> : <span className="h-4 w-10 border border-current" style={{ background: `linear-gradient(90deg,#${template.colors.accent},#${template.colors.accent2})` }} />}
                  </div>
                  <p className="text-sm font-black">{template.name}</p>
                  <p className={`mt-1 text-xs ${isSelected ? "text-primary-foreground/80" : "text-muted-foreground"}`}>{template.scenario} · {template.uiStyle}</p>
                  <p className={`mt-3 text-xs leading-5 ${isSelected ? "text-primary-foreground/70" : "text-muted-foreground"}`}>{template.bestFor}</p>
                </button>
              );
            })}
          </div>
        </div>

        <aside className="border border-border bg-card p-5">
          <div className="mb-5 flex items-start gap-3">
            <div className="border border-border bg-secondary p-2 text-secondary-foreground"><FileJson className="h-5 w-5" /></div>
            <div>
              <p className="font-black">Deck preview</p>
              <p className="text-sm text-muted-foreground">{deck ? `${deck.sections.length + 1} total slides` : "Invalid JSON"}</p>
            </div>
          </div>
          {deck ? (
            <div className="space-y-4">
              <div>
                <p className="text-xs font-bold uppercase tracking-widest text-muted-foreground">Title</p>
                <p className="mt-1 text-lg font-black leading-tight">{deck.title}</p>
                <p className="mt-1 text-sm text-muted-foreground">{deck.subtitle}</p>
              </div>
              <div className="max-h-52 space-y-2 overflow-auto border-y border-border py-3">
                {deck.sections.slice(0, 8).map((section, index) => (
                  <div key={`${section.title}-${index}`} className="flex gap-3 text-sm">
                    <span className="text-muted-foreground">{String(index + 1).padStart(2, "0")}</span>
                    <span className="font-semibold">{section.title}</span>
                    <span className="ml-auto text-xs text-muted-foreground">{resolveLayout(section, activeTemplate, index)}</span>
                  </div>
                ))}
              </div>
              <button type="button" onClick={handleGenerate} disabled={isGenerating} className="inline-flex w-full items-center justify-center gap-2 bg-primary px-4 py-3 text-sm font-black text-primary-foreground transition-colors hover:bg-primary/90 disabled:pointer-events-none disabled:opacity-60">
                <Download className="h-4 w-4" /> {isGenerating ? "Generating..." : "Generate & download PPT"}
              </button>
            </div>
          ) : <p className="text-sm text-muted-foreground">Paste valid JSON to preview the generated slide structure.</p>}
        </aside>
      </section>

      <section className="border-y border-border bg-muted/40">
        <div className="mx-auto max-w-7xl px-5 py-8 lg:px-8">
          <div className="mb-5 flex items-end justify-between gap-4">
            <div>
              <h2 className="text-2xl font-black tracking-normal">Supported layout outlines</h2>
              <p className="text-sm text-muted-foreground">Use these `layout` values in JSON to control slide structure, from comparisons to image placeholders.</p>
            </div>
          </div>
          <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
            {layoutOutlines.map((outline) => {
              const Icon = outline.icon;
              const selected = outline.id === selectedOutline;
              return (
                <button key={outline.id} type="button" onClick={() => setSelectedOutline(outline.id)} className={`border p-4 text-left transition-colors ${selected ? "border-primary bg-primary text-primary-foreground" : "border-border bg-card hover:bg-accent"}`}>
                  <div className="mb-3 flex items-center gap-2"><Icon className="h-4 w-4" /><p className="text-sm font-black">{outline.name}</p></div>
                  <p className={`text-xs leading-5 ${selected ? "text-primary-foreground/75" : "text-muted-foreground"}`}>{outline.description}</p>
                  <p className={`mt-3 text-[11px] font-bold uppercase tracking-wide ${selected ? "text-primary-foreground/80" : "text-muted-foreground"}`}>Types: {layoutVariants[outline.id].join(" · ")}</p>
                  <p className={`mt-3 font-mono text-[11px] ${selected ? "text-primary-foreground/75" : "text-muted-foreground"}`}>{outline.fields}</p>
                  <MiniLayoutPreview outline={outline} template={activeTemplate} />
                  <pre className={`mt-3 max-h-28 overflow-auto whitespace-pre-wrap border p-2 font-mono text-[10px] leading-4 ${selected ? "border-primary-foreground/30 bg-primary-foreground/10 text-primary-foreground/80" : "border-border bg-background text-muted-foreground"}`}>{layoutJsonExamples[outline.id]}</pre>
                </button>
              );
            })}
          </div>
        </div>
      </section>

      <section className="mx-auto grid max-w-7xl gap-6 px-5 py-8 lg:grid-cols-[0.8fr_1.2fr] lg:px-8">
        <div>
          <div className="mb-4 inline-flex items-center gap-2 border border-border bg-card px-3 py-2 text-xs font-semibold uppercase tracking-widest text-muted-foreground"><Sparkles className="h-4 w-4 text-primary" /> Prompt builder</div>
          <h2 className="text-2xl font-black tracking-normal">Copy a prompt for another AI to create compatible JSON.</h2>
          <p className="mt-3 text-sm leading-6 text-muted-foreground">The prompt explains the exact JSON shape, uniform slide rules, supported fields, layout options and optional source-material instructions.</p>
          <div className="mt-5 border border-border bg-card p-4">
            <p className="text-sm font-black">Prompt format</p>
            <p className="mt-2 text-sm text-muted-foreground">Neutral JSON command · 12 supported layouts · copyable examples</p>
          </div>
        </div>
        <div className="border border-border bg-card p-4">
          <div className="mb-3 flex items-center justify-between gap-3">
            <div className="flex items-center gap-2 text-sm font-black"><Clipboard className="h-4 w-4" /> AI JSON prompt</div>
            <button type="button" onClick={copyPrompt} className="inline-flex items-center gap-2 border border-border bg-secondary px-3 py-2 text-sm font-bold text-secondary-foreground hover:bg-accent"><Copy className="h-4 w-4" /> {copyStatus || "Copy prompt"}</button>
          </div>
          <textarea readOnly value={aiPrompt} className="min-h-[360px] w-full resize-y border border-input bg-background p-4 font-mono text-xs leading-5 text-foreground outline-none" />
        </div>
      </section>
    </main>
  );
};

export default Index;
