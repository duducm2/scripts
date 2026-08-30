from pathlib import Path

path = Path(r"C:\Users\eduev\Meu Drive\17 - Projects\scripts\mnemonics\web\index.html")
text = path.read_text(encoding="utf-8")
start = text.find(
    "      /* Palace overlay cockpit: ~20% snapshot + full atom text (no crop) */"
)
if start < 0:
    raise SystemExit("start marker not found")
end = text.find("    </style>", start)
if end < 0:
    raise SystemExit("end marker not found")

new_css = r"""      /* Palace overlay: 20% snapshot + magazine grid (atoms use space under image) */
      #overlay.open {
        display: flex;
        flex-direction: column;
        overflow: hidden;
      }
      #overlay .overlay-body {
        display: flex;
        flex-direction: column;
        flex: 1 1 auto;
        min-height: 0;
        gap: 0;
        padding: 0;
        overflow: auto;
      }
      #overlay .overlay-stage {
        display: grid;
        grid-template-columns: minmax(120px, 20%) repeat(3, minmax(0, 1fr));
        grid-template-rows: auto auto;
        align-content: start;
        gap: 0.5rem 0.55rem;
        box-sizing: border-box;
        height: calc(100vh - 3.6rem);
        min-height: calc(100vh - 3.6rem);
        max-height: calc(100vh - 3.6rem);
        padding: 0.55rem 0.65rem 0.65rem;
        overflow: auto;
        flex: 0 0 auto;
      }
      #overlay .overlay-image {
        grid-column: 1;
        grid-row: 1;
        align-self: start;
        justify-self: stretch;
        display: flex;
        align-items: flex-start;
        justify-content: center;
        width: 100%;
        height: auto;
        max-height: 26vh;
        margin: 0;
        padding: 0;
        border: 1px solid var(--line);
        border-radius: 8px;
        background: #07080b;
        cursor: zoom-in;
        overflow: hidden;
      }
      #overlay .overlay-image img {
        width: 100%;
        height: auto;
        max-height: 26vh;
        object-fit: contain;
        object-position: top center;
        display: block;
      }
      #overlay .overlay-image .missing-lg {
        padding: 0.75rem;
        font-size: 0.8rem;
      }
      /* Let beast cards participate in the stage magazine grid */
      #overlay .overlay-atoms {
        display: contents;
      }
      #overlay .overlay-atoms > h3 {
        display: none;
      }
      #overlay:not(.image-expanded) #ovAtoms {
        display: contents;
      }
      /* Default (3 beasts): sit beside the image */
      #overlay:not(.image-expanded) #ovAtoms > .beast-cluster {
        height: auto !important;
        min-height: 0;
        overflow: visible !important;
        gap: 0.35rem;
        padding: 0.45rem 0.55rem 0.5rem;
        border-radius: 8px;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="1"] > .beast-cluster:nth-child(1) {
        grid-column: 2 / -1;
        grid-row: 1;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="2"] > .beast-cluster:nth-child(1) {
        grid-column: 2 / 3;
        grid-row: 1;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="2"] > .beast-cluster:nth-child(2) {
        grid-column: 3 / -1;
        grid-row: 1;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="3"] > .beast-cluster:nth-child(1) {
        grid-column: 2;
        grid-row: 1;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="3"] > .beast-cluster:nth-child(2) {
        grid-column: 3;
        grid-row: 1;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="3"] > .beast-cluster:nth-child(3) {
        grid-column: 4;
        grid-row: 1;
      }
      /* Four: 3 beside image + 1 full-width under */
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="4"] > .beast-cluster:nth-child(1) {
        grid-column: 2;
        grid-row: 1;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="4"] > .beast-cluster:nth-child(2) {
        grid-column: 3;
        grid-row: 1;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="4"] > .beast-cluster:nth-child(3) {
        grid-column: 4;
        grid-row: 1;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="4"] > .beast-cluster:nth-child(4) {
        grid-column: 1 / -1;
        grid-row: 2;
      }
      /* Five: 3 beside image + 2 under image (uses former black void) */
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="5"] > .beast-cluster:nth-child(1) {
        grid-column: 2;
        grid-row: 1;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="5"] > .beast-cluster:nth-child(2) {
        grid-column: 3;
        grid-row: 1;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="5"] > .beast-cluster:nth-child(3) {
        grid-column: 4;
        grid-row: 1;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="5"] > .beast-cluster:nth-child(4) {
        grid-column: 1 / 3;
        grid-row: 2;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="5"] > .beast-cluster:nth-child(5) {
        grid-column: 3 / -1;
        grid-row: 2;
      }
      #overlay:not(.image-expanded) .beast-cluster--lead {
        border-left-color: #e8c56a;
        box-shadow: inset 0 0 0 1px rgba(232, 197, 106, 0.18);
      }
      #overlay:not(.image-expanded) .beast-cluster-head {
        padding-bottom: 0.25rem;
      }
      #overlay:not(.image-expanded) .beast-cluster-head .lbl {
        font-size: 0.62rem;
      }
      #overlay:not(.image-expanded) .beast-cluster-head .beast-name {
        font-size: 0.9rem;
        line-height: 1.25;
      }
      #overlay:not(.image-expanded) .beast-cluster-head .sensory-chip {
        font-size: 0.72rem;
        max-width: min(12rem, 48%);
        white-space: normal;
      }
      #overlay:not(.image-expanded) .beast-cluster-atoms {
        display: grid !important;
        grid-template-columns: 1fr !important;
        gap: 0.4rem !important;
        overflow: visible !important;
        min-height: 0;
      }
      #overlay:not(.image-expanded) .atom-card {
        padding: 0.45rem 0.55rem;
        border-radius: 6px;
        overflow: visible !important;
        height: auto !important;
        display: block;
        min-height: 0;
      }
      #overlay:not(.image-expanded) .atom-card .atom-card-head {
        margin: 0 0 0.25rem;
        padding-bottom: 0.2rem;
      }
      #overlay:not(.image-expanded) .atom-card .field {
        margin: 0.3rem 0 0;
        font-size: 0.8rem;
        line-height: 1.35;
        max-height: none !important;
        overflow: visible !important;
      }
      #overlay:not(.image-expanded) .atom-card .lbl {
        font-size: 0.62rem;
        margin-bottom: 0.05rem;
      }
      #overlay:not(.image-expanded) .atom-card .field-concept {
        margin-top: 0.3rem;
        padding: 0.35rem 0.45rem;
        border-radius: 5px;
        max-height: none !important;
        overflow: visible !important;
      }
      #overlay .overlay-details {
        display: grid;
        grid-template-columns: 1fr;
        gap: 1rem;
        padding: 1rem 1.1rem 2rem;
        border-top: 1px solid var(--line);
        background: #101218;
      }
      @media (min-width: 1100px) {
        #overlay .overlay-details {
          grid-template-columns: 1.1fr 1fr;
        }
        #overlay .overlay-gallery {
          grid-column: 1 / -1;
        }
      }
      #overlay.image-expanded .overlay-stage {
        display: flex;
        grid-template-columns: none;
        grid-template-rows: none;
        height: calc(100vh - 3.6rem);
        min-height: calc(100vh - 3.6rem);
        max-height: calc(100vh - 3.6rem);
      }
      #overlay.image-expanded .overlay-atoms,
      #overlay.image-expanded .overlay-details {
        display: none !important;
      }
      #overlay.image-expanded .overlay-image {
        flex: 1 1 auto;
        width: 100%;
        max-width: none;
        max-height: none;
        height: 100%;
        border: none;
        border-radius: 0;
        cursor: zoom-out;
        align-items: center;
        grid-column: auto;
        grid-row: auto;
      }
      #overlay.image-expanded .overlay-image img {
        width: 100%;
        height: 100%;
        max-height: none;
        object-fit: contain;
      }
      @media (max-width: 900px) {
        #overlay .overlay-stage {
          display: flex;
          flex-direction: column;
          height: auto;
          min-height: auto;
          max-height: none;
        }
        #overlay .overlay-image {
          flex: 0 0 auto;
          width: 100%;
          max-width: none;
          max-height: 32vh;
          grid-column: auto;
          grid-row: auto;
        }
        #overlay .overlay-image img {
          max-height: 32vh;
        }
        #overlay .overlay-atoms {
          display: flex;
          flex-direction: column;
        }
        #overlay .overlay-atoms > h3 {
          display: block;
          margin: 0 0 0.4rem;
          font-size: 0.95rem;
        }
        #overlay:not(.image-expanded) #ovAtoms {
          display: grid !important;
          grid-template-columns: 1fr !important;
          gap: 0.5rem !important;
        }
        #overlay:not(.image-expanded) #ovAtoms > .beast-cluster {
          grid-column: auto !important;
          grid-row: auto !important;
        }
      }
"""

path.write_text(text[:start] + new_css + text[end:], encoding="utf-8")
print("ok")
