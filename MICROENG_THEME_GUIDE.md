# MicroEng Design System & Theme Guide

A concise style guide for MicroEng apps. Logos for this project live in `Logos/`.

---

## Logos

Located in `Logos/`:

| File | Purpose |
|------|---------|
| `microeng-logo.png`   | Primary logo (full) |
| `microeng-logo2.png`  | Logo variant 2 |
| `microeng-logo3.png`  | Logo variant 3 |
| `microeng_logotray.png` | System tray / small icon |

Example (web/React):
```jsx
import logoImage from '../Logos/microeng-logo.png';
import trayIcon from '../Logos/microeng_logotray.png';
```

---

## Color System

Light theme (default):
```css
:root {
  --color-bg-body: #f5f7fb;
  --color-bg-panel: #ffffff;
  --color-bg-elevated: #f9fafc;
  --color-bg-muted: #eef1f4;
  --color-overlay: rgba(15, 23, 42, 0.45);

  --color-text-primary: #111827;
  --color-text-secondary: #374151;
  --color-text-tertiary: #6b7280;

  --color-accent: #8ba9d9;
  --color-accent-strong: #6b89c9;
  --color-accent-soft: rgba(139, 169, 217, 0.15);

  --color-success: #2d8a4e;
  --color-danger: #c74b4b;
  --color-warning: #b8860b;
  --color-info: #2b7a8c;

  --status-neutral-bg: #f6f8fa;
  --status-neutral-border: #d0d7de;
  --status-neutral-color: #4b5563;
  --status-info-bg: #f0f9ff;
  --status-info-border: #7eb8e4;
  --status-info-color: #0969da;
  --status-success-bg: #f0fff4;
  --status-success-border: #4ade80;
  --status-success-color: #166534;
  --status-warning-bg: #fffbeb;
  --status-warning-border: #d4a528;
  --status-warning-color: #92400e;
  --status-danger-bg: #fef2f2;
  --status-danger-border: #f87171;
  --status-danger-color: #b91c1c;

  --color-border: rgba(15, 23, 42, 0.28);
  --color-border-strong: rgba(15, 23, 42, 0.45);
  --shadow-soft: 0 18px 30px rgba(15, 23, 42, 0.08);
  --shadow-inset: inset 0 0 0 1px rgba(255, 255, 255, 0.6);
}
```

Dark theme (apply `.theme-dark`):
```css
.theme-dark {
  --color-bg-body: #111418;
  --color-bg-panel: #171b20;
  --color-bg-elevated: #1f242c;
  --color-bg-muted: #242a34;
  --color-overlay: rgba(0, 0, 0, 0.55);

  --color-text-primary: #f5f7fb;
  --color-text-secondary: #c4ccd8;
  --color-text-tertiary: #8f99ad;

  --color-border: rgba(255, 255, 255, 0.32);
  --color-border-strong: rgba(255, 255, 255, 0.5);
  --shadow-soft: 0 10px 30px rgba(0, 0, 0, 0.35);

  --color-accent: #4f7afc;
  --color-accent-strong: #3d66eb;
  --color-accent-soft: rgba(79, 122, 252, 0.18);

  --color-success: #4dbf86;
  --color-danger: #ff6b6b;
  --color-warning: #ffae42;
  --color-info: #40c4ff;
}
```

Node/card header colors (pastel):
```css
:root {
  --node-header-navisworks: #a3d9b1;
  --node-header-general: #f4e4b4;
  --node-header-projectwise: #b4d8e8;
  --node-header-text: #000000;
}
```

---

## Typography

Font stack:
```css
font-family: "Segoe UI", system-ui, -apple-system, BlinkMacSystemFont, Arial, sans-serif;
-webkit-font-smoothing: antialiased;
-moz-osx-font-smoothing: grayscale;
```

Type scale (px):
- h1: 24 / 700 / -0.025em
- h2: 20 / 600 / -0.025em
- h3: 18 / 700 / -0.025em
- Category header: 14 / 600 / uppercase / 0.05em
- Body: 14 / 400
- Small/Label: 12-13 / 500-600
- Tiny/Badge: 10-11 / 500-600 / 0.05em

---

## Spacing, Radius, Motion

```css
:root {
  --radius-sm: 6px;
  --radius-md: 10px;
  --radius-lg: 14px;

  --transition-fast: 0.16s ease;
  --transition-medium: 0.24s ease;
}
```

Common spacing:
- Component gap: 8px
- Card padding: 16px
- Section margin: 24px
- Panel padding: 20px
- Small gap: 4px
- Large gap: 12px

---

## Component Patterns (samples)

Buttons:
```css
.btn {
  background: var(--color-accent);
  border: none;
  color: #fff;
  font-size: 12px;
  font-weight: 500;
  padding: 8px 12px;
  border-radius: var(--radius-sm);
  cursor: pointer;
  transition: transform var(--transition-fast),
              box-shadow var(--transition-fast),
              background-color var(--transition-fast);
  box-shadow: 0 6px 20px rgba(79, 122, 252, 0.25);
}
.btn:hover { background: var(--color-accent-strong); transform: translateY(-1px); }
.btn:disabled { opacity: 0.6; cursor: not-allowed; transform: none; }
.btn-secondary { background: var(--color-bg-muted); color: var(--color-text-secondary); border: 1px solid var(--color-border); box-shadow: none; }
.btn-danger { background: var(--color-danger); box-shadow: 0 6px 20px rgba(199, 75, 75, 0.25); }
```

Card:
```css
.card {
  background: var(--color-bg-panel);
  border: 1px solid var(--color-border);
  border-radius: var(--radius-md);
  padding: 16px;
  box-shadow: var(--shadow-soft);
  transition: all 0.2s ease;
}
.card:hover {
  border-color: var(--color-accent);
  transform: translateY(-2px);
  box-shadow: 0 8px 25px rgba(0, 0, 0, 0.12);
}
```

Inputs:
```css
input, textarea, select {
  font-family: inherit;
  color: var(--color-text-primary);
  background: var(--color-bg-panel);
  border: 1px solid var(--color-border);
  border-radius: var(--radius-sm);
  padding: 8px 10px;
  font-size: 14px;
  transition: border-color var(--transition-fast),
              background-color var(--transition-fast);
}
input:hover { border-color: var(--color-border-strong); }
input:focus {
  outline: none;
  border-color: var(--color-accent);
  box-shadow: 0 0 0 3px var(--color-accent-soft);
}
```

Modal shell:
```css
.modal-overlay {
  position: fixed;
  inset: 0;
  background: var(--color-overlay);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 2000;
}
.modal-content {
  background: var(--color-bg-panel);
  border-radius: var(--radius-lg);
  max-width: 520px;
  width: 92%;
  box-shadow: 0 18px 48px rgba(0, 0, 0, 0.35);
  border: 1px solid var(--color-border);
}
```

---

## Accessibility

- Maintain WCAG 2.1 AA contrast.
- Visible focus states (outline or box-shadow).
- Hover states for all interactive elements; disabled state uses reduced opacity and no hover lift.

---

## Quick Reference (variables)

```css
/* Backgrounds */
var(--color-bg-body)      /* Page background */
var(--color-bg-panel)     /* Cards, modals */
var(--color-bg-elevated)  /* Headers, elevated */
var(--color-bg-muted)     /* Subdued areas */

/* Text */
var(--color-text-primary)
var(--color-text-secondary)
var(--color-text-tertiary)

/* Accents */
var(--color-accent)
var(--color-accent-strong)
var(--color-success)
var(--color-danger)
var(--color-warning)
var(--color-info)

/* Borders */
var(--color-border)
var(--color-border-strong)

/* Radius */
var(--radius-sm)  /* 6px */
var(--radius-md)  /* 10px */
var(--radius-lg)  /* 14px */

/* Motion */
var(--transition-fast)    /* 0.16s */
var(--transition-medium)  /* 0.24s */
```

---

*MicroEng Theme Guide v1.1 - aligned to Navisworks plugin assets (Logos/).*
