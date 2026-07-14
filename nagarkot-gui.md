---
name: nagarkot-gui
description: Applies logo-based color theme, structured header (logo left, title centered), single footer, full-screen layout, and consistent spacing while preserving functional flexibility.
---

# 1. Objective

Automatically transform any existing GUI application so that it conforms to the Nagarkot Forwarders Pvt Ltd. visual identity while preserving functionality.

The workflow must ensure:

1. Strictly use existing virtual environment or else create a new Virtual Environment, installed required Packages and then only execute any commands.
1. Ask User to attach an image of existing GUI if not attached and then decide if a fullscreen is required, if yes then implement fullscreen and make sure to utilise all space or else dont go for full screen..
2. Logo positioned at top-left (fixed height = 20 units).
3. Tool name centered across entire window width (true visual center).
4. Optional subtitle/tagline centered below tool name.
5. Brand color theme applied consistently.
6. “Nagarkot Forwarders Pvt Ltd.” appears once per screen, bottom-left.
7. Layout remains flexible and function-driven.
8. No framework-specific hardcoding.

---

# 2. Core Non-Negotiable Branding Rules

## 2.1 Header Requirements

* Logo:

  * Positioned top-left.
  * Fixed height: 20 units.
  * Width auto-scaled proportionally.
  * Non-clickable.
  * Proper left padding.

* Tool Name:

  * Must be Centered relative to total window width and not relative to remaining space after logo.
  * Bold, high-contrast.

* Subtitle (Optional):

  * Centered below tool name.
  * Muted color.
  * Smaller than title.

Header height must remain flexible.
The coding LLM determines the optimal height for visual balance.

---

## 2.2 Footer Requirement

* Text: "Nagarkot Forwarders Pvt. Ltd. \u00A9"
* Appears once per screen.
* Positioned bottom-left.
* Small, muted gray.
* Minimal visual dominance.

In multi-screen applications:

* Once per screen/page.
* If persistent layout exists, reuse footer component.

---

# 3. Brand Design System

## 3.1 Color Palette (Derived from Logo)

Primary Blue: `#1F3F6E`
Accent Red: `#D8232A`
Dark Text: `#1E1E1E`
Muted Gray: `#6B7280`
Light Background: `#F4F6F8`
Panel White: `#FFFFFF`
Border Gray: `#E5E7EB`
Hover Blue: `#2A528F`

If exact color sampling is available, override with precise extracted values.

---

## 3.2 Typography Guidelines

* Clean sans-serif.
* Title: Bold, large.
* Subtitle: Medium, muted.
* Body text: Neutral dark.
* Footer: Small, subtle.

No decorative fonts.

---

# 4. Structural Refactor Model

Transform application into logical zones:

```
ROOT (Full Screen)
│
├── HEADER (Dynamic Height)
│   ├── Logo (Left)
│   ├── Centered Title Block (Absolute Center)
│
├── BODY (Flexible Content Area)
│   ├── Existing Functional Components
│
└── FOOTER
    └── Company Name (Bottom-Left)
```

The layout structure must adapt to the application’s needs.
Do not impose fixed grid structures unless appropriate.

---

# 5. Automatic Refactor Steps

## Step 1: Analyze Existing UI

* Detect root container.
* Identify layout method.
* Detect existing header/footer components.
* Map component hierarchy.
* Identify repeated color patterns.

---

## Step 2: Normalize Window Behavior

* Remove fixed window sizing.
* Enable full-screen launch.
* Ensure responsive resizing.
* Preserve usability constraints.

---

## Step 3: Inject or Refactor Header

If no header exists:

* Create header container at top.
* Height determined dynamically.

If header exists:

* Refactor to comply with brand rules.

### Title Centering Implementation (Critical)

Must achieve true horizontal center alignment.

Allowed strategies:

* Absolute positioning relative to window width.
* Overlay layered container.
* Grid with independent center alignment.
* Anchor-based centering.

The coding LLM chooses method based on framework.

---

## Step 4: Refactor Body Content

Do NOT impose rigid layout template.

Instead:

* Wrap existing UI into a body container.
* Apply consistent padding.
* Normalize spacing between sections.
* Group related fields logically.
* Improve hierarchy without breaking functionality.

Background:

* Light Background color.

---

## Step 5: Component Standardization

### Buttons

Primary:

* Background: Primary Blue.
* Text: White.
* No heavy borders.
* Hover: Hover Blue.
* Subtle rounding.

Secondary:

* White background.
* Blue border.
* Blue text.

### Inputs

* White background.
* Light border.
* Blue focus state.
* Consistent height.

### Tables

* Header: Light tinted background.
* Alternating row shading subtle.
* Borders light and consistent.

### Alerts

Error:

* Accent Red.
* Clear visibility.

Warning:

* Muted orange/red derivative.

Success:

* Blue-toned confirmation (avoid green if possible to maintain brand harmony).

---

# 6. Layout Flexibility Rules

The workflow must NOT:

* Force equal column distribution.
* Force centered content.
* Force panel symmetry.
* Restrict custom field arrangements.

Instead:

* Respect existing functional grouping.
* Improve spacing consistency.
* Improve alignment clarity.
* Maintain logical visual flow.

---

# 7. Multi-Screen Handling

## Separate Windows

Apply entire workflow independently per window.

## Tab-Based Applications

Single header + footer.
Body content changes per tab.

## Frame/Page Switching

Create reusable header and footer components.
Inject into each page.

Footer appears once per visible screen.

---

# 8. Color Application Enforcement

* Header text: Primary Blue.
* Primary actions: Blue.
* Critical states: Red only when necessary.
* Background: Light neutral.
* Avoid introducing unrelated color tones.

Remove inconsistent blues or random accent colors.

---

# 9. Automatic Validation Checklist

After refactor:

1. Application launches full-screen.
2. Logo height equals 20 units.
3. Tool name visually centered to total width.
4. Subtitle centered if present.
5. Footer appears once per screen.
6. All primary buttons use brand blue.
7. No conflicting color schemes remain.
8. Layout remains functional.
9. No framework-specific hardcoding introduced.
10. Header height determined dynamically and visually balanced.

---

# 10. Antigravity IDE Execution Prompt

Use the following transformation instruction:

“Refactor the entire GUI to comply with Nagarkot Brand Standard.
Convert root window to full-screen.
Inject or refactor header with left-aligned logo (height 20), tool name centered across full window width, optional centered subtitle, and dynamic header height.
Preserve flexible body layout but normalize spacing and visual hierarchy.
Inject bottom-left footer containing ‘Nagarkot Forwarders Pvt Ltd.’ once per screen.
Apply consistent brand color tokens to all UI components.
Do not hardcode framework-specific styling.
Perform transformation automatically while preserving functionality.”

---

# 11. Anti-Patterns to Eliminate

* Title aligned beside logo.
* Multiple color palettes.
* Footer centered.
* Overuse of red.
* Mixed font sizes without hierarchy.
* Inconsistent spacing.
* Cluttered header bars.

---

# 12. Outcome

All applications will:

* Maintain unique functional layout.
* Present consistent Nagarkot branding.
* Display professional hierarchy.
* Scale across tools without redesign effort.
* Preserve usability while enforcing identity consistency.