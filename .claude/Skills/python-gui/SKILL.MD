# Python GUI Design Style Guide
**Version:** 1.0.0  
**Target:** Professional Desktop Applications

---

## 1. Color Palette Standards

Professional UIs follow strict contrast and semantic rules. Never use raw primary colors. Use the **60-30-10 Rule**: 60% dominant background, 30% structural elements, 10% accent focus.

### 1.1 Dark Theme Profile
*   **Primary Background (60%):** `#121212` (Deep charcoal, avoids OLED black smear)
*   **Surface/Card Background (30%):** `#1E1E1E` (Elevated components, sidebars)
*   **Accent Color (10%):** `#0A84FF` (System blue, high visibility)
*   **Success Semantic:** `#30D158` (Soft emerald green)
*   **Warning/Error Semantic:** `#FF453A` (Desaturated crimson red)
*   **Primary Text:** `#FFFFFF` (90% Opacity)
*   **Secondary Text:** `#EBEBF5` (60% Opacity)

### 1.2 Light Theme Profile
*   **Primary Background (60%):** `#F8F9FA` (Soft off-white, reduces eye strain)
*   **Surface/Card Background (30%):** `#FFFFFF` (Pure white for cards, inputs)
*   **Accent Color (10%):** `#007AFF` (Classic corporate blue)
*   **Success Semantic:** `#34C759` (Vibrant green)
*   **Warning/Error Semantic:** `#FF3B30` (Alert red)
*   **Primary Text:** `#1C1C1E` (Near black)
*   **Secondary Text:** `#3C3C43` (60% Opacity gray)

---

## 2. Typography & Fonts

### 2.1 Font Families
Do not use non-standard fonts that require manual installation unless bundled directly. Use cross-platform system fonts.

*   **Windows Target:** `Segoe UI`, `Aptos`
*   **macOS Target:** `SF Pro Display`, `San Francisco`
*   **Linux Target:** `Ubuntu Sans`, `Cantarell`
*   **Fallback Standard:** `Helvetica Neue`, `Arial`, `sans-serif`
*   **Monospace (Code/Data):** `Consolas`, `SF Mono`, `JetBrains Mono`

### 2.2 Type Scale Hierarchy


| Role | Size (pt) | Weight | Line Spacing | Usage |
| :--- | :--- | :--- | :--- | :--- |
| **Heading 1** | 24pt | Bold (700) | 1.2 | Main Window Title / Dashboard Header |
| **Heading 2** | 18pt | SemiBold (600) | 1.3 | Section / Tab Headers |
| **Subtitle** | 14pt | Medium (500) | 1.4 | Card Headers / Sub-sections |
| **Body Large**| 12pt | Regular (400) | 1.5 | Primary Form Labels / Inputs |
| **Body Small**| 10pt | Regular (400) | 1.5 | Data Tables / Body Copy / Settings |
| **Caption** | 9pt | Light (300) | 1.2 | Tooltips / Help Text / Status Bar |

---

## 3. Theme & Widget Specifications

### 3.1 Layout & Grid Spacing
*   **Base Unit:** 8px grid system. All margins, padding, and gaps must be multiples of 8 (8px, 16px, 24px, 32px).
*   **Window Padding:** Minimum 16px padding from window borders to content.
*   **Widget Gap:** 8px horizontal/vertical gap between tightly coupled elements.

### 3.2 Visual Attributes
*   **Corner Radius (Border Radius):** 
    *   Small Widgets (Buttons, Inputs): `4px` or `6px`
    *   Large Containers (Cards, Frames): `8px` or `12px`
*   **Borders:** `1px` solid outline. Use a contrast ratio of at least 3:1 against background colors.
*   **Shadows:** Avoid heavy black dropshadows. Use a soft blur: `rgba(0, 0, 0, 0.15)` with a 4px blur radius.

### 3.3 Interactive States
*   **Default:** Standard theme color.
*   **Hover:** Darken light themes by 10% / Lighten dark themes by 10%.
*   **Active/Pressed:** Shift color value toward accent or compress size slightly.
*   **Disabled:** Drop opacity to `35%` and change cursor to default arrow.

---

## 4. Framework Implementation Matrix
| Framework | Layout System | Theming Support | Widget Library | Threading Model |
| :--- | :--- | :--- | :--- | :--- |
/* NiceUI | Flexbox-based | Built-in theming engine | Custom widgets with Tailwind CSS classes | Asyncio-compatible */

Use Tailwind's spacing utilities, which naturally scale on a 4px/8px progression.
*   **Page Padding:** Always wrap your main layout in a container with `classes('p-6')` (24px) or `classes('p-8')` (32px).
*   **Component Spacing:** Use `classes('gap-4')` (16px) for standard item spacing inside rows and columns.


## 5. UI Layout Best Practices

*   **Responsive Resizing:** Always anchor elements properly using layouts (`QVBoxLayout`/`QHBoxLayout` in Qt, or `grid(sticky="nsew")` in Tkinter). Never use absolute positioning (`place()` method).
*   **Dynamic Loading Indicators:** Provide a visual spinner or progress bar if an action takes longer than 200ms.
*   **Threading Separation:** Run heavy calculations or API fetches inside a background worker thread (`QThread` or Python `threading`). Never block the primary UI thread.
