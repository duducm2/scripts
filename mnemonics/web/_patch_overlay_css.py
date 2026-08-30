from pathlib import Path

path = Path(r"C:\Users\eduev\Meu Drive\17 - Projects\scripts\mnemonics\web\index.html")
text = path.read_text(encoding="utf-8")
start = text.find("      /* Palace overlay:")
if start < 0:
    raise SystemExit("start not found")
end = text.find("    </style>", start)
if end < 0:
    raise SystemExit("end not found")

new_css = r"""      /* Palace overlay: large snapshot on top + beast row below */
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
        display: flex;
        flex-direction: column;
        gap: 0.55rem;
        box-sizing: border-box;
        height: calc(100vh - 3.6rem);
        min-height: calc(100vh - 3.6rem);
        max-height: calc(100vh - 3.6rem);
        padding: 0.55rem 0.75rem 0.65rem;
        overflow: hidden;
        flex: 0 0 auto;
      }
      #overlay .overlay-image {
        flex: 0 0 auto;
        width: 100%;
        max-width: none;
        height: 38vh;
        max-height: 38vh;
        min-height: 180px;
        margin: 0;
        padding: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        border: 1px solid var(--line);
        border-radius: 10px;
        background: #07080b;
        cursor: zoom-in;
        overflow: hidden;
      }
      #overlay .overlay-image img {
        width: 100%;
        height: 100%;
        max-height: none;
        object-fit: contain;
        object-position: center center;
        display: block;
      }
      #overlay .overlay-image .missing-lg {
        padding: 1.25rem;
        font-size: 0.9rem;
      }
      #overlay .overlay-atoms {
        flex: 1 1 auto;
        min-height: 0;
        display: flex;
        flex-direction: column;
        overflow: auto;
        padding: 0;
      }
      #overlay .overlay-atoms > h3 {
        display: none;
      }
      #overlay:not(.image-expanded) #ovAtoms {
        flex: 0 0 auto;
        display: grid !important;
        grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)) !important;
        grid-auto-rows: auto !important;
        gap: 0.5rem !important;
        width: 100%;
        align-content: start;
        overflow: visible;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="1"] {
        grid-template-columns: minmax(0, 1fr) !important;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="2"] {
        grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="3"] {
        grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="4"] {
        grid-template-columns: repeat(4, minmax(0, 1fr)) !important;
      }
      #overlay:not(.image-expanded) #ovAtoms[data-beast-count="5"] {
        grid-template-columns: repeat(5, minmax(0, 1fr)) !important;
      }
      #overlay:not(.image-expanded) #ovAtoms > .beast-cluster {
        grid-column: auto !important;
        grid-row: auto !important;
        height: auto !important;
        min-height: 0;
        overflow: visible !important;
        gap: 0.35rem;
        padding: 0.45rem 0.55rem 0.5rem;
        border-radius: 8px;
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
      /* Concept-first: hide Quote/Story until Details (D) */
      #overlay:not(.show-atom-details) .atom-card .field-quote,
      #overlay:not(.show-atom-details) .atom-card .field-story {
        display: none !important;
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
        flex-direction: column;
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
        height: 100%;
        max-height: none;
        min-height: 0;
        border: none;
        border-radius: 0;
        cursor: zoom-out;
      }
      #overlay.image-expanded .overlay-image img {
        width: 100%;
        height: 100%;
        max-height: none;
        object-fit: contain;
      }
      @media (max-width: 900px) {
        #overlay .overlay-stage {
          height: auto;
          min-height: auto;
          max-height: none;
          overflow: auto;
        }
        #overlay .overlay-image {
          height: 32vh;
          max-height: 32vh;
          min-height: 140px;
        }
        #overlay:not(.image-expanded) #ovAtoms,
        #overlay:not(.image-expanded) #ovAtoms[data-beast-count="4"],
        #overlay:not(.image-expanded) #ovAtoms[data-beast-count="5"] {
          grid-template-columns: 1fr !important;
        }
        #overlay .overlay-atoms > h3 {
          display: block;
          margin: 0 0 0.4rem;
          font-size: 0.95rem;
        }
      }
"""

path.write_text(text[:start] + new_css + text[end:], encoding="utf-8")
print("ok", end - start, "->", len(new_css))
