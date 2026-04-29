import { ChangeEvent, useMemo, useState } from "react";
import pptxgen from "pptxgenjs";
import {
  BarChart3,
  BriefcaseBusiness,
  Check,
  Columns3,
  Download,
  FileJson,
  GraduationCap,
  Lightbulb,
  LineChart,
  Palette,
  Presentation,
  Rocket,
  ScrollText,
  Upload,
  UserRound,
} from "lucide-react";

type TemplateId =
  | "business"
  | "education"
  | "sales"
  | "proposal"
  | "resume"
  | "timeline"
  | "comparison"
  | "brainstorming"
  | "report"
  | "creative";

type SlideContent = {
  title: string;
  subtitle?: string;
  bullets: string[];
  notes?: string;
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
  icon: typeof Presentation;
  colors: {
    bg: string;
    ink: string;
    muted: string;
    panel: string;
    accent: string;
    accent2: string;
  };
};

const templates: Template[] = [
  {
    id: "business",
    name: "Boardroom Brief",
    scenario: "Business",
    icon: BriefcaseBusiness,
    colors: { bg: "101820", ink: "F8FAFC", muted: "CBD5E1", panel: "1C2A35", accent: "F2AA4C", accent2: "86C5DA" },
  },
  {
    id: "education",
    name: "Lesson Builder",
    scenario: "Education",
    icon: GraduationCap,
    colors: { bg: "F7F3E8", ink: "20322E", muted: "5B6B64", panel: "FFFFFF", accent: "2F7D6D", accent2: "E3A535" },
  },
  {
    id: "sales",
    name: "Revenue Pitch",
    scenario: "Sales pitch",
    icon: LineChart,
    colors: { bg: "0B1320", ink: "FFFFFF", muted: "B6C2D2", panel: "162235", accent: "22C55E", accent2: "38BDF8" },
  },
  {
    id: "proposal",
    name: "Project Case",
    scenario: "Project proposal",
    icon: Rocket,
    colors: { bg: "F4F6F8", ink: "1E293B", muted: "64748B", panel: "FFFFFF", accent: "0F766E", accent2: "D97706" },
  },
  {
    id: "resume",
    name: "Career Profile",
    scenario: "Resume",
    icon: UserRound,
    colors: { bg: "111111", ink: "FAFAFA", muted: "D4D4D4", panel: "252525", accent: "E11D48", accent2: "F5F5F5" },
  },
  {
    id: "timeline",
    name: "Milestone Map",
    scenario: "Timeline",
    icon: ScrollText,
    colors: { bg: "FEF7ED", ink: "2D241F", muted: "765B4B", panel: "FFFFFF", accent: "C2410C", accent2: "2563EB" },
  },
  {
    id: "comparison",
    name: "Decision Matrix",
    scenario: "Comparison",
    icon: Columns3,
    colors: { bg: "F8FAFC", ink: "172033", muted: "64748B", panel: "FFFFFF", accent: "4F46E5", accent2: "14B8A6" },
  },
  {
    id: "brainstorming",
    name: "Idea Wall",
    scenario: "Brainstorming",
    icon: Lightbulb,
    colors: { bg: "FFF8DB", ink: "302A12", muted: "766B35", panel: "FFFFFF", accent: "F59E0B", accent2: "DB2777" },
  },
  {
    id: "report",
    name: "Insight Report",
    scenario: "Report",
    icon: BarChart3,
    colors: { bg: "F1F5F9", ink: "0F172A", muted: "475569", panel: "FFFFFF", accent: "0369A1", accent2: "65A30D" },
  },
  {
    id: "creative",
    name: "Studio Concept",
    scenario: "Creative",
    icon: Palette,
    colors: { bg: "201A2E", ink: "FFF7ED", muted: "E9D5FF", panel: "382B4A", accent: "FB7185", accent2: "FBBF24" },
  },
];

const sampleJson = JSON.stringify(
  {
    title: "Quarterly Growth Plan",
    subtitle: "A JSON-powered deck generated in the browser",
    slides: [
      { title: "Executive summary", bullets: ["Expand self-serve acquisition", "Improve enterprise conversion", "Prioritize onboarding speed"] },
      { title: "Key metrics", bullets: ["Revenue up 18% quarter over quarter", "Pipeline coverage at 3.4x", "Retention improved by 6 points"] },
      { title: "Next steps", bullets: ["Launch vertical campaign", "Run pricing experiment", "Create weekly KPI review"] },
    ],
  },
  null,
  2,
);

const safeText = (value: unknown, fallback = "Untitled") => {
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    const text = String(value).trim();
    return text.length ? text.slice(0, 220) : fallback;
  }
  return fallback;
};

const valuesToBullets = (value: unknown): string[] => {
  if (Array.isArray(value)) return value.flatMap(valuesToBullets).slice(0, 8);
  if (value && typeof value === "object") {
    return Object.entries(value)
      .map(([key, item]) => `${key}: ${safeText(item, JSON.stringify(item).slice(0, 120))}`)
      .slice(0, 8);
  }
  return [safeText(value, "Detail")];
};

const normalizeDeck = (jsonText: string): DeckContent => {
  const parsed = JSON.parse(jsonText) as Record<string, unknown>;
  const root = Array.isArray(parsed) ? { title: "Generated presentation", slides: parsed } : parsed;
  const rawSlides = [root.slides, root.sections, root.items, root.content].find(Array.isArray) as unknown[] | undefined;

  const sections = (rawSlides?.length ? rawSlides : Object.entries(root).filter(([, value]) => typeof value === "object"))
    .slice(0, 12)
    .map((entry, index): SlideContent => {
      const value = Array.isArray(rawSlides) ? entry : (entry as [string, unknown])[1];
      const key = Array.isArray(rawSlides) ? `Section ${index + 1}` : safeText((entry as [string, unknown])[0], `Section ${index + 1}`);

      if (value && typeof value === "object" && !Array.isArray(value)) {
        const record = value as Record<string, unknown>;
        const title = safeText(record.title ?? record.heading ?? record.name ?? key, key);
        const bulletsSource = record.bullets ?? record.points ?? record.items ?? record.details ?? record.description ?? record.summary ?? record;
        return {
          title,
          subtitle: safeText(record.subtitle ?? record.kicker ?? "", ""),
          bullets: valuesToBullets(bulletsSource).filter((bullet) => bullet !== title).slice(0, 6),
          notes: safeText(record.notes ?? "", ""),
        };
      }

      return { title: key, bullets: valuesToBullets(value).slice(0, 6) };
    });

  if (!sections.length) {
    sections.push({ title: "Overview", bullets: Object.entries(root).map(([key, value]) => `${key}: ${safeText(value)}`).slice(0, 6) });
  }

  return {
    title: safeText(root.title ?? root.name ?? root.topic, "Generated presentation"),
    subtitle: safeText(root.subtitle ?? root.description ?? root.summary, "Built from your JSON content"),
    sections,
  };
};

const addFooter = (slide: pptxgen.Slide, template: Template, index: number) => {
  slide.addShape("line", { x: 0.55, y: 6.92, w: 12.2, h: 0, line: { color: template.colors.accent, width: 1 } });
  slide.addText(String(index).padStart(2, "0"), { x: 12.05, y: 6.72, w: 0.8, h: 0.25, fontFace: "Aptos", fontSize: 8, color: template.colors.muted, align: "right" });
};

const addTitleSlide = (pptx: pptxgen, deck: DeckContent, template: Template) => {
  const slide = pptx.addSlide();
  slide.background = { color: template.colors.bg };
  slide.addShape("rect", { x: 0, y: 0, w: 3.1, h: 7.5, fill: { color: template.colors.accent }, line: { color: template.colors.accent } });
  slide.addShape("rect", { x: 0.45, y: 0.6, w: 2.2, h: 0.12, fill: { color: template.colors.accent2 }, line: { color: template.colors.accent2 } });
  slide.addText(template.scenario.toUpperCase(), { x: 3.65, y: 1.2, w: 7.8, h: 0.32, fontFace: "Aptos", fontSize: 12, bold: true, color: template.colors.accent });
  slide.addText(deck.title, { x: 3.6, y: 1.72, w: 8.5, h: 1.55, fontFace: "Aptos Display", fontSize: 42, bold: true, color: template.colors.ink, breakLine: false, fit: "shrink" });
  slide.addText(deck.subtitle, { x: 3.65, y: 3.55, w: 7.1, h: 0.65, fontFace: "Aptos", fontSize: 16, color: template.colors.muted, fit: "shrink" });
  slide.addText(`${deck.sections.length} content slides`, { x: 3.65, y: 5.7, w: 2.6, h: 0.34, fontFace: "Aptos", fontSize: 13, bold: true, color: template.colors.ink });
  slide.addShape("arc", { x: 10.9, y: 5.25, w: 1.6, h: 1.6, line: { color: template.colors.accent2, width: 5, transparency: 12 } });
};

const addContentSlide = (pptx: pptxgen, item: SlideContent, template: Template, index: number) => {
  const slide = pptx.addSlide();
  slide.background = { color: template.colors.bg };
  slide.addText(item.title, { x: 0.62, y: 0.52, w: 8.6, h: 0.55, fontFace: "Aptos Display", fontSize: 28, bold: true, color: template.colors.ink, fit: "shrink" });
  if (item.subtitle) slide.addText(item.subtitle, { x: 0.66, y: 1.08, w: 7.8, h: 0.28, fontFace: "Aptos", fontSize: 10, color: template.colors.muted, fit: "shrink" });

  const bullets = item.bullets.length ? item.bullets : ["No bullet details were provided in this JSON section."];
  const cols = template.id === "comparison" ? 2 : bullets.length > 4 ? 2 : 1;
  const cardW = cols === 2 ? 5.7 : 8.6;
  bullets.slice(0, 8).forEach((bullet, bulletIndex) => {
    const col = cols === 2 ? bulletIndex % 2 : 0;
    const row = cols === 2 ? Math.floor(bulletIndex / 2) : bulletIndex;
    const x = 0.72 + col * 6.05;
    const y = 1.72 + row * 1.08;
    slide.addShape(pptxgen.ShapeType.roundRect, { x, y, w: cardW, h: 0.78, rectRadius: 0.05, fill: { color: template.colors.panel, transparency: 0 }, line: { color: template.colors.panel } });
    slide.addShape(pptxgen.ShapeType.ellipse, { x: x + 0.2, y: y + 0.2, w: 0.34, h: 0.34, fill: { color: bulletIndex % 2 ? template.colors.accent2 : template.colors.accent }, line: { color: bulletIndex % 2 ? template.colors.accent2 : template.colors.accent } });
    slide.addText(bullet, { x: x + 0.72, y: y + 0.16, w: cardW - 1.0, h: 0.44, fontFace: "Aptos", fontSize: 13, color: template.colors.ink, fit: "shrink", breakLine: false });
  });

  if (["timeline", "proposal", "report"].includes(template.id)) {
    slide.addShape(pptxgen.ShapeType.line, { x: 9.85, y: 1.6, w: 0, h: 4.6, line: { color: template.colors.accent, width: 2 } });
    [0, 1, 2].forEach((dot) => slide.addShape(pptxgen.ShapeType.ellipse, { x: 9.68, y: 1.8 + dot * 1.85, w: 0.34, h: 0.34, fill: { color: template.colors.accent2 }, line: { color: template.colors.accent2 } }));
  }

  addFooter(slide, template, index);
};

const createPresentation = async (deck: DeckContent, template: Template) => {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "JSON PPT Generator";
  pptx.subject = template.scenario;
  pptx.title = deck.title;
  pptx.company = "Lovable";
  pptx.theme = {
    headFontFace: "Aptos Display",
    bodyFontFace: "Aptos",
  };

  addTitleSlide(pptx, deck, template);
  deck.sections.forEach((section, index) => addContentSlide(pptx, section, template, index + 1));

  const fileName = `${deck.title.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "presentation"}-${template.id}.pptx`;
  await pptx.writeFile({ fileName });
};

const Index = () => {
  const [jsonText, setJsonText] = useState(sampleJson);
  const [selectedTemplate, setSelectedTemplate] = useState<TemplateId>("business");
  const [error, setError] = useState("");
  const [isGenerating, setIsGenerating] = useState(false);

  const activeTemplate = templates.find((template) => template.id === selectedTemplate) ?? templates[0];
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
                Upload or paste JSON, select one of ten fixed presentation scenarios, then download a finished `.pptx` file generated locally in your browser.
              </p>
            </div>
            <div className="mt-8 grid max-w-xl grid-cols-3 border border-border bg-card">
              <div className="border-r border-border p-4">
                <p className="text-2xl font-black">10</p>
                <p className="text-xs text-muted-foreground">fixed templates</p>
              </div>
              <div className="border-r border-border p-4">
                <p className="text-2xl font-black">12</p>
                <p className="text-xs text-muted-foreground">slides max</p>
              </div>
              <div className="p-4">
                <p className="text-2xl font-black">.pptx</p>
                <p className="text-xs text-muted-foreground">download</p>
              </div>
            </div>
          </div>

          <div className="border border-border bg-card p-4 shadow-2xl shadow-primary/10">
            <div className="mb-3 flex items-center justify-between gap-3">
              <div>
                <p className="text-sm font-bold">JSON input</p>
                <p className="text-xs text-muted-foreground">Supports `title`, `subtitle`, `slides`, `sections`, `items`, or nested objects.</p>
              </div>
              <label className="inline-flex cursor-pointer items-center gap-2 border border-border bg-secondary px-3 py-2 text-sm font-semibold text-secondary-foreground transition-colors hover:bg-accent">
                <Upload className="h-4 w-4" /> Upload
                <input className="sr-only" type="file" accept="application/json,.json" onChange={handleUpload} />
              </label>
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
          <div className="mb-4 flex items-end justify-between gap-4">
            <div>
              <h2 className="text-2xl font-black tracking-normal">Choose a fixed template</h2>
              <p className="text-sm text-muted-foreground">Each style adapts the same JSON into scenario-specific slide composition.</p>
            </div>
          </div>
          <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5">
            {templates.map((template) => {
              const Icon = template.icon;
              const isSelected = template.id === selectedTemplate;
              return (
                <button
                  key={template.id}
                  type="button"
                  onClick={() => setSelectedTemplate(template.id)}
                  className={`group border p-4 text-left transition-all ${isSelected ? "border-primary bg-primary text-primary-foreground shadow-lg shadow-primary/20" : "border-border bg-card hover:-translate-y-0.5 hover:bg-accent"}`}
                >
                  <div className="mb-5 flex items-center justify-between">
                    <Icon className="h-5 w-5" />
                    {isSelected ? <Check className="h-4 w-4" /> : null}
                  </div>
                  <p className="text-sm font-black">{template.name}</p>
                  <p className={`mt-1 text-xs ${isSelected ? "text-primary-foreground/75" : "text-muted-foreground"}`}>{template.scenario}</p>
                </button>
              );
            })}
          </div>
        </div>

        <aside className="border border-border bg-card p-5">
          <div className="mb-5 flex items-start gap-3">
            <div className="border border-border bg-secondary p-2 text-secondary-foreground">
              <FileJson className="h-5 w-5" />
            </div>
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
              <div className="max-h-48 space-y-2 overflow-auto border-y border-border py-3">
                {deck.sections.slice(0, 6).map((section, index) => (
                  <div key={`${section.title}-${index}`} className="flex gap-3 text-sm">
                    <span className="text-muted-foreground">{String(index + 1).padStart(2, "0")}</span>
                    <span className="font-semibold">{section.title}</span>
                  </div>
                ))}
              </div>
              <button
                type="button"
                onClick={handleGenerate}
                disabled={isGenerating}
                className="inline-flex w-full items-center justify-center gap-2 bg-primary px-4 py-3 text-sm font-black text-primary-foreground transition-colors hover:bg-primary/90 disabled:pointer-events-none disabled:opacity-60"
              >
                <Download className="h-4 w-4" /> {isGenerating ? "Generating..." : "Generate & download PPT"}
              </button>
            </div>
          ) : (
            <p className="text-sm text-muted-foreground">Paste valid JSON to preview the generated slide structure.</p>
          )}
        </aside>
      </section>
    </main>
  );
};

export default Index;
