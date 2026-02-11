---
name: slide-deck
description: Generate PowerPoint (.pptx) presentations with Databricks branding that can be imported into Google Slides. Use when user asks to create slides, presentations, pitch decks, or slide decks.
allowed-tools: Read, Write, Bash, Glob, AskUserQuestion, TodoWrite
---

# Slide Deck Generator

Generate professional PowerPoint presentations (.pptx) with Databricks branding. Output files can be imported directly into Google Slides.

## Getting Started

When a user wants to create a slide deck, first determine which mode to use:

**Ask the user:**
1. I have existing content as the primary source. → **Content Mode**
2. Build the deck through discussion — you can still share supporting docs. → **Interview Mode**

---

## Content Mode

For users who have existing content as the primary source. Claude analyzes the material and asks targeted follow-ups to fill gaps.

The content provides the substance. Read deeply, then align on how to present it.

Examples: PRDs, design docs, QBRs, post-mortems, RCAs, decision docs, one-pagers, documentation pages, code implementations, or query results.

> **Note:** Content can include multiple files from different sources. Users may also ask Claude to run commands or queries (e.g., SQL, CLI tools) to pull data as part of the content gathering.

### Flow

1. **Before reading any files, ask the user:**

   Understand what you're about to read and why—this sets the lens for analysis.

   - What are you providing? (PRD, design doc, post-mortem, analysis, etc.)
   - What's the context? Why are you creating this deck?
   - Who's the audience?
   - Any specific focus or angle you want?

2. **Read and analyze the content deeply**
   - Review all provided files/text thoroughly
   - Identify key themes, points, data, and structure

3. **Return with clarifying questions**

   After reading, ask a few questions to fill gaps in your knowledge—the content provides the depth.

   - Clarify anything unclear or ambiguous in the content
   - Confirm the main message and angle
   - Check what to emphasize vs. summarize
   - Check if anything should be added or excluded

   A few targeted questions, not deep discovery.

4. **Design the outline**
   - Map content to slide types and sequence
   - See **Design Principles** and **Slide Types** in Reference below

5. **Confirm understanding**
   - Present your proposed structure/outline
   - Get user approval

6. **Generate the deck**

---

## Interview Mode

For users who want to build the deck through discussion. Claude will ask questions to discover the content. Users can still provide supplemental documentation.

The conversation provides the substance. Interview deeply to discover what to present.

### Flow

1. **Conduct the interview**

   Interview the user in depth using the AskUserQuestion tool. This is how you discover the content—be thorough.

   - Understand the story they want to tell
   - Dig into key messages, supporting evidence, the "why" behind points
   - Explore what the audience cares about and what they should do after
   - Ask for specifics: examples, data, context
   - Uncover constraints or things to avoid
   - If user provides supporting docs, incorporate them

   Don't stop after one round. Continue interviewing in multiple rounds until you meet the success criteria below.

2. **Design the outline**
   - Map content to slide types and sequence
   - See **Design Principles** and **Slide Types** in Reference below

3. **Confirm understanding**
   - Present your proposed structure/outline
   - Get user approval

4. **Generate the deck**

### Success Criteria

Before designing, you must understand:

**Required:**
- Topic and scope
- Target audience
- Purpose/goal of the presentation
- Key message(s)
- Main sections/points to cover
- Desired length
- Important details for each section (data, examples, context)

**If relevant:**
- Specific metrics or stats
- Customer stories or examples
- Technical details
- Constraints (what to include/avoid)

Keep asking until you have enough detail.

---

## Generating the Deck

Once you have designed your presentation:

1. **Generate content JSON** following the structure below

2. **Write JSON** to a temp file (e.g., `/tmp/slides-content.json`)

3. **Create output directory if needed:**
   ```bash
   mkdir -p ./generated-slides
   ```

4. **Run the generator:**
   ```bash
   python3 {baseDir}/scripts/generate-pptx.py \
     --input /tmp/slides-content.json \
     --output ./generated-slides/presentation.pptx
   ```

5. **Tell user** where to find the file and how to import to Google Slides:
   > Done! Your presentation is at `./generated-slides/[name].pptx`
   >
   > To import to Google Slides:
   > 1. Go to slides.google.com
   > 2. File → Open → Upload
   > 3. Select the .pptx file

6. **Invite feedback**: Let the user know they can request changes. Ask if they'd like any edits—different wording, additional slides, reordering, different slide types, etc.

---

## Iterating on the Deck

When the user requests changes after generation:

1. Read the existing JSON from `/tmp/slides-content.json`
2. Make targeted edits to the specific slides they mentioned
3. Regenerate with the same command

For structural changes (adding, removing, reordering slides), edit the `slides` array directly. For content changes, locate the slide by its type and title, then update the relevant fields.

Continue iterating until the user is satisfied.

---

# Reference

## Design Principles

When designing your outline, consider:

### Every piece of information has a shape

Before choosing a slide type, identify the shape of what you're presenting:

| Shape | What it looks like | Consider using |
|-------|-------------------|----------------|
| Sequential | Steps, phases, process, workflow | `timeline`, `agenda`, `cards` |
| Comparative | A vs B, options, trade-offs | `two-column`, `pros-cons`, `comparison` |
| Categorical | Features, types, capabilities, pillars | `icon-grid`, `three-column`, `three-column-icons` |
| Emphatic | One key stat, one bold claim | `big-number`, `callout` |
| Evidence | Proof, credibility, testimonial | `quote`, `logos`, `stat-row` |
| Status | Progress, done/not done | `checklist` |
| Mixed content | Text with visual/diagram | `card-left`, `card-right`, `card-full` |

**Bullets (`content`) are not the default.** They're one option among many. If you're about to make your third bullet slide in a row, stop and ask: what shape does this information actually have?

### Think about rhythm

A presentation is a sequence. Each slide exists in context of what came before and after.

- Don't repeat the same layout back-to-back-to-back
- Dense slides (lots of bullets, details) need breathing room (callout, big-number, section break)
- Vary the visual treatment to maintain audience engagement
- Section slides (`section`) create natural breaks between topics

### Design checklist

Before generating, mentally walk through your slide sequence:

- [ ] Does each slide type match the shape of its content?
- [ ] Is there visual variety, or are there long runs of identical layouts?
- [ ] Do section breaks fall at natural topic transitions?
- [ ] Are there moments of impact (big-number, callout, quote) at key points?
- [ ] Does the rhythm feel right—density balanced with breathing room?

---

## Content JSON Structure

```json
{"title": "Deck Title", "author": "Name", "date": "January 2025", "slides": [...]}
```

## Slide Types Reference

See "Designing the Presentation" above for how to choose between these.

### Standard Slides

| Type | Description | Schema |
|------|-------------|--------|
| `title` | Opening slide (dark) | title, subtitle?, author?, date? |
| `section` | Section divider (dark) | title, subtitle? |
| `section-description` | Section with body text | title, subtitle?, description? or bullets? |
| `content` | Bullet points | title, subtitle?, bullets: [strings], notes? |
| `one-column` | Narrow single column | title, content? or bullets? |
| `closing` | Final slide (dark) | title |

### Column Layouts

| Type | Description | Schema |
|------|-------------|--------|
| `two-column` | Side-by-side columns | title, subtitle?, left_header?, left: [strings], right_header?, right: [strings] |
| `two-column-icons` | Two columns with icons | title, subtitle?, columns: [{header, items, icon?}] |
| `three-column` | Three columns | title, subtitle?, columns: [{header, items}] |
| `three-column-icons` | Three columns with icons | title, subtitle?, columns: [{header, items, icon?}] |

### Cards & Layouts

| Type | Description | Schema |
|------|-------------|--------|
| `cards` | Three-card layout | title, subtitle?, cards: [{header, content? or items?}] |
| `card-right` | Content left, card right | title, subtitle?, bullets? or content?, card_content? |
| `card-left` | Card left, content right | title, subtitle?, card_content?, bullets? or content? |
| `card-full` | Full-width card | title, subtitle?, content? |

### Data & Metrics

| Type | Description | Schema |
|------|-------------|--------|
| `big-number` | Hero stat | number, text, subtitle? |
| `stat-row` | Multiple metrics in row | title, stats: [{value, label}] |
| `comparison` | VS layout (diamond) | title, left_label, right_label |
| `pros-cons` | Pro/con lists | title, pros_header?, cons_header?, pros: [strings], cons: [strings] |

### Visual & Sequential

| Type | Description | Schema |
|------|-------------|--------|
| `agenda` | Numbered hexagon list | title, items: [strings] |
| `timeline` | Sequential steps | title, steps: [{title, description}] |
| `icon-grid` | Feature grid with icons | title, items: [{icon (emoji recommended), title, description?}] |
| `checklist` | Checkbox items | title, items: [{text, checked: bool}] |

### Quotes & Social Proof

| Type | Description | Schema |
|------|-------------|--------|
| `quote` | Testimonial (dark) | quote, attribution? |
| `callout` | Bold statement (dark) | text, source? |
| `logos` | Logo/partner display | title, subtitle?, logos: [strings] |

### Architecture Diagrams (Imported)

| Type | Description | Schema |
|------|-------------|--------|
| `architecture` | Pre-built architecture diagram from catalog | catalog_slide (1-based slide number), notes? |

**Schema notation:** `field?` = optional, `[strings]` = string array, `[{a, b}]` = object array, `bool` = true/false

**Total: 26 slide types**

## Template & Branding

The generator uses the official Databricks corporate template. Branding, backgrounds, and styling are handled automatically:

- **Dark backgrounds**: Title, section, callout, quote, and closing slides use Databricks dark templates for visual impact
- **Light backgrounds**: All other content slides use light templates for readability

All slides automatically include proper footer, branding, and visual consistency.

## Architecture Diagram Catalog

A catalog of 55 pre-built Databricks architecture diagrams is available for import into generated decks. Use these for technical audiences (L200+) when discussing platform components, deployment models, security, or data architecture.

### When to Use

- Technical deep-dives with SAs, architects, or engineering teams
- Platform component discussions (compute, security, data access)
- Reference architecture walkthroughs
- Any deck at L200+ depth that benefits from official architecture visuals

### JSON Format

```json
{"type": "architecture", "catalog_slide": 14, "notes": "Optional custom speaker notes"}
```

- `catalog_slide` (required): 1-based slide number from the catalog (see index below)
- `notes` (optional): Custom speaker notes. If omitted, the catalog's built-in talk track is used.

### Usage Guidelines

- Architecture slides are imported as-is with all shapes, images, and formatting preserved
- They use a white background and do not carry over the catalog's slide master
- Mix freely with other slide types — they work alongside generated slides
- Multiple architecture slides can reference the same catalog slide
- Section header slides from the catalog (kind=section) can also be imported, but prefer using the native `section` slide type for visual consistency

### Catalog Index

Organized by topic section. Slides marked with shape counts to indicate complexity.

**Data Intelligence Platform**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 6 | Modern Data Platform - Framework | 26 | overview, L100-200 |
| 7 | Databricks "Data Intelligence Platform" | 42 | overview, L100-200 |

**Databricks AI**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 9 | Mosaic AI (overview) | 16 | overview, L100-200 |
| 10 | Mosaic AI (detailed) | 59 | L200 |

**Lakehouse & Data Intelligence Platform**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 12 | From Lakehouse to Data Intelligence Platform | 21 | overview, L100-200 |

**High Level Architecture**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 14 | Data Intelligence Platform Overview | 61 | overview, L100-200 |
| 15 | Data Intelligence Platform Classic | 63 | overview, L100-200 |
| 16 | Data Intelligence Platform Serverless | 65 | overview, L100-200 |
| 17 | Data Intelligence Platform Serverless (storage) | 66 | overview, L100-200 |

**Classic Compute**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 19 | High Level Architecture | 23 | overview, L100-200 |
| 20 | Detailed High Level Architecture | 46 | overview |

**SQL Warehouse**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 22 | High Level Architecture | 26 | overview, L100-200 |
| 23 | Detailed High Level Architecture | 53 | overview |

**Serverless Compute**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 25 | High Level Architecture | 39 | overview, L100-200 |
| 26 | High Level Architecture (SQL) | 34 | overview, L100-200 |
| 27 | High Level Architecture - Details | 33 | overview |

**Serverless Model Serving, Vector Search, Online Tables**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 29 | High Level Architecture | 48 | overview, L100-200 |

**Online Transaction Processing**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 31 | High Level Architecture | 37 | — |
| 32 | Lakebase integration | 58 | — |

**Security**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 34 | Security Highlights | 36 | overview |
| 35 | Secure Cluster Communication | 62 | overview |
| 36 | Private Link | 33 | overview |
| 37 | Customer Managed Keys | 29 | overview |
| 38 | Lakebase security highlights | 29 | — |

**Scaling Databricks**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 40 | Scaling Databricks Compute | 21 | overview |
| 41 | Multi Workspace Architecture | 43 | overview |

**Data Architecture**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 44 | Medallion, the best practice pipeline | 31 | overview, L100-200 |
| 45 | Data Modeling: Dimensional Modeling for DWHs | 47 | overview |
| 46 | Data Modeling: Modern use cases (ML and AI) | 28 | overview |
| 47 | Enhanced medallion architecture | 30 | overview |

**Data Access**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 49 | High Level Data Access from Databricks | 29 | overview, L100-200 |
| 50 | Data Access Security | 25 | overview |
| 51 | Life of a query with Unity Catalog | 26 | L200 |
| 52 | Life of a query with Unity Catalog (detailed) | 28 | L200 |
| 53 | Querying database sources with Unity | 39 | — |

**Delta & Iceberg**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 55 | Data Formats - Parquet | 6 | overview |
| 56 | Delta: Reliability and performance features | 6 | overview, L100-200 |
| 57 | Reading and writing Iceberg and Delta | 83 | — |

**Lakehouse Monitoring and Observability**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 59 | Lakehouse Observability | 58 | — |
| 60 | Lakehouse Monitoring | 59 | — |

**Filesystems in Databricks**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 62 | DBFS (Databricks File System) | 26 | overview |
| 63 | Filesystems in Databricks | 20 | overview |

**Reference Architecture**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 65 | Databricks Data Intelligence Platform (full reference) | 133 | — |
| 66 | 1 Built in ingestion from SaaS and databases | 99 | — |
| 67 | 2 Batch ingestion and ETL | 119 | — |
| 68 | 3 Streaming and Change Data Capture | 123 | — |
| 69 | 4 Machine Learning (traditional) | 139 | — |
| 70 | 5 Gen AI: Agents | 151 | — |
| 71 | 6 Business Intelligence | 109 | — |
| 72 | 7 Business Apps | 110 | — |
| 73 | 8 Lakehouse Federation | 87 | — |
| 74 | 9 Catalog Federation | 89 | — |
| 75 | 10a Sharing Data | 94 | — |
| 76 | 10b Consuming Shared Data | 90 | — |

**Appendix**

| Slide | Title | Shapes | Tags |
|-------|-------|--------|------|
| 78 | Communication with DBFS | 34 | azure |

## Content Best Practices

- **Titles**: Max 8 words, clear and action-oriented
- **Bullets**: 3-5 per slide, max 12 words each
- **One idea per slide** - don't overcrowd
- **Include speaker notes** for important context
- **Use section slides** to break up long presentations
