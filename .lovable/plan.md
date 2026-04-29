I’ll update the app so templates are not just color skins, but full presentation design systems with selectable layout behavior, visible outlines, image placeholders, and copyable AI prompts.

## What will change

1. Expand templates from 10 to 20
- Keep the existing 10 scenario templates: business, education, sales pitch, project proposal, resume, timeline, comparison, brainstorming, report, creative.
- Add a `uiStyle` category to every template.
- Create 2 templates for each UI style:
  - Skeuomorphism
  - Neobrutalism
  - Glassmorphism
  - Neumorphism
  - Minimalism
  - Existing/original adaptive style groups, so the original set remains represented
- Each template will have a different visual identity: color scheme, card treatment, borders, shadows, accent shapes, typography choices, and slide composition.

2. Add layout types and outline types
- Add a layout system separate from templates, so users can pick or preview how JSON will be arranged.
- Include layout types such as:
  - Cover page
  - Dot/bullet listing
  - Comparison columns
  - Timeline/milestones
  - Process/steps
  - Image placeholder + text
  - 2x2 cards
  - Report/stat callouts
  - Resume/profile sections
  - Brainstorming/idea wall
  - Proposal/problem-solution
  - Table/matrix-style outline
- The home page will show these layouts as “outline cards” so users understand what JSON structure the app supports.

3. Make generated PPT slides use varied layouts
- Instead of every content slide becoming the same card/bullet listing, generation will choose a layout based on:
  - selected template/scenario,
  - section fields in the JSON such as `layout`, `image`, `comparison`, `steps`, `metrics`, `timeline`, `columns`, `quote`, or `cards`,
  - fallback logic when JSON is simple.
- Add image placeholder slide layouts. Since users are providing JSON rather than actual images, the PPT will reserve designed image areas with labels like “Image placeholder” or use any image URL/text metadata from the JSON as a caption.

4. Add prompt generator section on the home page
- Add a new section where users can copy a prompt for another AI.
- The prompt will instruct the other AI to produce JSON in the exact shape this app understands.
- The prompt will include placeholders for:
  - topic,
  - audience,
  - presentation goal,
  - template/scenario,
  - preferred layout types,
  - optional source materials.
- Users will be able to select a template/layout and copy the matching prompt.

5. Improve the visible home page UI
- Add filters or grouping for templates by UI style.
- Show each template’s scenario, UI style, palette, and best use case.
- Add a “supported JSON outline” section with examples.
- Add a prompt-copy section with a textarea/code block and copy button.
- Keep upload/paste JSON, deck preview, template selection, and download flow intact.

## Technical plan

1. Refactor template definitions
- Extend `TemplateId` and `Template` with:
  - `uiStyle`
  - `layoutBias`
  - `shapeStyle`
  - `shadowStyle`
  - `borderStyle`
  - richer color tokens
  - prompt guidance metadata
- Replace the 10-template array with a 20-template catalog.

2. Add layout model
- Create layout type definitions such as `LayoutType` and `OutlineType`.
- Add layout metadata for name, description, JSON fields, and prompt instructions.
- Update deck normalization to preserve more structured content instead of flattening everything into bullets too early.

3. Update PPT generation
- Add slide rendering functions for multiple layouts:
  - cover/title
  - bullet list
  - comparison
  - timeline
  - process steps
  - image placeholder
  - card grid
  - stat/report
  - matrix/table-like
- Route each section to the best renderer using JSON keys and selected template style.
- Apply UI-style-specific styling:
  - skeuomorphism: dimensional panels, highlights, realistic controls
  - neobrutalism: hard black borders, offset shadows, loud colors
  - glassmorphism: translucent panels, soft gradients, blur-like layered shapes
  - neumorphism: soft same-color raised/inset surfaces
  - minimalism: whitespace, thin lines, restrained palette

4. Add prompt-copy UI
- Generate a prompt string from selected template + selected layout guidance.
- Add copy-to-clipboard behavior with success feedback.
- Include a sample JSON schema users can copy or hand to another AI.

5. Validate and clean up
- Remove unused dependencies if still present, such as `file-saver` if it is no longer imported.
- Check TypeScript compatibility with `pptxgenjs` shape and text options.
- Verify the generate button still works with the current sample JSON and with richer structured JSON examples.

## Expected result

The page will become a full JSON-to-PPT design studio: users can upload/paste JSON, choose from 20 visually distinct templates, understand and copy supported layout outlines, generate decks with more than cover + bullet-list slides, and copy a ready-made prompt for another AI to produce compatible JSON for any topic or source material.