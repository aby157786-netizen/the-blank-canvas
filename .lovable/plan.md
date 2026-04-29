I’ll update the JSON command/prompt so it is no longer vague, and make the app explain the exact JSON format in a uniform, beginner-friendly way.

## What will change

1. Rewrite the AI JSON prompt
- Make the prompt say clearly: “Create one JSON object with `title`, `subtitle`, and `slides`.”
- Explain that every slide should use the same simple structure:

```json
{
  "layout": "bullets",
  "title": "Slide title",
  "subtitle": "Optional short context",
  "bullets": ["Short point", "Short point"],
  "imagePrompt": "Optional image description"
}
```

- Make it clear that not every slide needs every field.
- Make layout choice simple: choose one layout value per slide, then use the matching fields for that layout.
- Keep the instruction neutral, not tied to a specific designer/template/outline.

2. Add a “uniform JSON rules” section to the prompt
The prompt will list rules like:
- Always return valid JSON only.
- Use `slides`, not random section names.
- Every slide must be an object.
- Use arrays of strings for `bullets`, `steps`, `metrics`, `cards`, `comparison`, `columns`, and `rows`.
- Keep text short to avoid PowerPoint overflow.
- Use `imagePrompt` when an image is needed but no file/URL exists.

3. Improve supported layout explanations
For every supported layout, the prompt will show:
- Layout name
- When to use it
- Required or best fields
- Simple JSON example
- 2–3 possible outline types

Example:

```text
Layout: bullets
Use for: key points, recommendations, problem lists
Best fields: layout, title, bullets
Example:
{
  "layout": "bullets",
  "title": "Key takeaways",
  "bullets": ["Point 1", "Point 2", "Point 3"]
}
```

4. Add clearer examples in the app
- Replace the current sample JSON with a cleaner uniform example.
- Add a small “How JSON should look” helper near the JSON input so users understand the required format before generating.
- Add a copyable starter JSON template that users can edit.

5. Fix wording in the Prompt Builder section
- The text currently says the prompt includes selected template/UI style/layout, but the prompt was made neutral. I’ll correct that wording so it accurately says the prompt explains the app’s supported JSON format and layouts.

## Technical plan

- Update `sampleJson` in `src/pages/Index.tsx` to a simpler, uniform deck example.
- Refactor `buildPrompt()` to use a clearer structure:
  - task
  - input placeholders
  - required JSON shape
  - uniform slide object rules
  - layout-by-layout guide
  - full example JSON
  - common mistakes to avoid
- Improve `layoutJsonExamples` where needed so examples follow the same naming and field style.
- Add a small helper block in the UI above or below the JSON textarea showing the basic JSON shape.
- Update Prompt Builder copy to remove misleading references to selected templates.

## Expected result

Users and other AI tools will understand exactly how to create JSON for the app. The JSON format will feel consistent, easy to copy, and easier to edit without guessing which fields the app supports.