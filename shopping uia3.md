# Shopee Search Field — UIA Context (Junior AI Reference)

## Platform identification (Mercado Livre & Shopee)

These rules apply to **both** Mercado Livre and Shopee when implementing or debugging shortcuts:

1. **Both platforms**: The shortcut **identification** issue (e.g. search field not found, wrong element) can occur on **both** Mercado Livre and Shopee. Use the same identification and URL-verification approach for each.

2. **Do not use window title for platform detection.** The browser window title changes dynamically (e.g. to the product name). Do **not** use `WinGetTitle` or title substring (e.g. "Mercado Livre", "Shopee") to decide if the active page is one of these e-commerce platforms.

3. **Use the URL to verify platform.** Extract and analyze the **URL** (e.g. from the Chrome address bar via UIA: `Type: 50004`, `AccessKey: "Ctrl+L"`, then `addressBar.Value`) to verify whether the active page belongs to Mercado Livre (`mercadolivre.com` / `mercadolibre.com`) or Shopee (`shopee.com`).

4. **Execute shortcuts only after URL confirmation.** Run the corresponding keyboard shortcuts (e.g. Shift+S for search focus) **only after** the platform has been confirmed via the URL. Do not enable or fire shortcuts based on title or other unstable signals.

5. **Prioritize operational efficiency.** During both the identification step (getting and checking the URL) and the execution step (resolving the target element and invoking the shortcut), prioritize efficiency (e.g. cache by HWND, bounded UIA calls, avoid redundant work).

---

## Target element

**Shopee eCommerce search text field** — the main site search input on shopee.com.br.

## Selection method

- **Keyboard shortcut**: **Shift + S**
- **Implementation**: Hotkey `+s` in [Shift keys.ahk](Shift%20keys.ahk) under `#HotIf IsShopeeActive()` (Shopee Brazil shortcuts).

## UIA data

### Numeric path from document root

Use this path when the root is the browser document (e.g. `Shopee_GetDocRoot()` or `GetCurrentDocumentElement()`):

```
1,1,2,3,1
```

Meaning: 1st child → 1st child → 2nd child → 3rd child → 1st child (Group → main Group → shopee-top → shopee-searchbar → ComboBox).

### Full path from window (reference)

Path from the Chrome window root (e.g. `UIA.ElementFromHandle(hwnd)`):

```
2,1,1,2,2,1,1,1,1,1,2,3,1
```

The document (RootWebArea) is at `2,1,1,2,2,1,1,1,1,1`.

### Structured path (Type/ClassName steps, documentation only)

Not parsed by current code; for reference and possible future tooling:

```
{T:33,CN:"BrowserRootView"}, {T:33}, {T:33}, {T:33,CN:"BrowserView"}, {T:33,CN:"View"}, {T:33}, {T:33}, {T:30}, {T:26}, {T:26}, {T:26,CN:"shopee-top shopee-top--sticky"}, {T:26,CN:"shopee-searchbar"}, {T:3}
```

### Element properties

| Property      | Value                                                                                  |
| ------------- | -------------------------------------------------------------------------------------- |
| Type          | 50003 (ComboBox)                                                                       |
| Name          | "Buscar na Shopee"                                                                     |
| LocalizedType | "caixa de combinação"                                                                  |
| ClassName     | "shopee-searchbar-input\_\_input" (tree dump may show `shopee-searchbar-input**input`) |

---

## Instructions for Junior AI

1. **Shortcut**: The shortcut (Shift + S) is working properly on Shop T (Shopee). Do not treat the shortcut as the source of failure when debugging.

2. **Focus on identification**: When something goes wrong, treat the **identification process** as the likely source of the issue (e.g. UIA resolution, FindElement conditions, path, or tree changes). Prioritize investigating and fixing how the search field is located, not the hotkey or activation.

3. **Efficiency**: Do not apply or rely on the efficiency canon of MT5 when evaluating efficiency in this context.

---

When automating the **Shopee search field**:

1. **Root**: Use the document root (e.g. `Shopee_GetDocRoot()` or the browser’s current document element). Do not use the window root unless you are using the full path from window.

2. **Resolve the element** (in order of preference):
   - Prefer property-based: `root.FindElement({ Type: 50003, Name: "Buscar na Shopee" })`.
   - If that fails or for robustness, use the numeric path from document root: `root.ElementFromPath("1,1,2,3,1")` or `root.ElementFromPathExist("1,1,2,3,1")`.

3. **Interact**: Call `SetFocus()` on the element, or if that fails, `Click()`.

No new processing layer is required; the existing UIA library supports numeric paths. This file is the single source of truth for path and properties.
