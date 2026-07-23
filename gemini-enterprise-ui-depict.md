# UI Schema: Gemini Enterprise Main Interface

## 1. Global Dashboard View
*   **Header**
    *   `Title`: Gemini Enterprise
    *   `User_Profile`: "E" avatar (Top Right)
*   **Hero Section**
    *   `Greeting`: "Hello, Eduardo" (with AI sparkle icon)
    *   `Tagline`: "Let's get some work done!"
*   **Prompt Input Component**
    *   `Placeholder_Text`: "Ask anything, search your data, @mention or /tools"
    *   `Left_Action_Group`: 
        *   `Action_Add`: [+] (Attachment)
        *   `Action_Tools`: [Slider/Nodes Icon] (Triggers Tool Menu)
        *   `Action_Data`: [Database Icon] 
    *   `Right_Action_Group`:
        *   `Dropdown_Model_Selector`: [Currently "Auto v"] (Triggers Model Menu)
        *   `Action_Submit`: [Send/Arrow Icon]
*   **Secondary Actions**
    *   `Button`: "Improve your prompt"

---

## 2. Interaction State: Model Selection Menu (Image 1)
**Event Trigger:** `onClick(Dropdown_Model_Selector)`
**Component Type:** `Dropdown List`
**State:** Open
**List Items:**
1.  `Item_ID`: auto
    *   `Label`: Auto
    *   `Description`: Gemini chooses the best fit
    *   `Status`: Selected (Checked)
2.  `Item_ID`: 3.1_pro
    *   `Label`: 3.1 Pro
    *   `Description`: State-of-the-art reasoning
3.  `Item_ID`: 3.5_flash
    *   `Label`: 3.5 Flash
    *   `Description`: Frontier intelligence built for speed
4.  `Item_ID`: 2.5_pro
    *   `Label`: 2.5 Pro
    *   `Description`: Solves complex problems

---

## 3. Interaction State: Tools Menu (Image 2)
**Event Trigger:** `onClick(Action_Tools)`
**Component Type:** `Context Menu / Popover`
**State:** Open
**List Items:**
1.  `Action_Item`: Search company data
    *   `Icon`: Database + Search
2.  `Action_Item`: Search the web
    *   `Icon`: Globe
3.  `Action_Item`: Deep Research
    *   `Icon`: Search + Sparkles
4.  `Action_Item`: Create images
    *   `Icon`: Image + Sparkles
5.  `Action_Item`: Create videos (Veo 3.1)
    *   `Icon`: Clapperboard

---

## 4. Bottom Section: Announcements
**Component Type:** `Horizontal Scroll / Card Grid`
**Cards:**
*   **Card 1**: `Title`: Prompt Gallery | `Date`: 3 days ago | `Snippet`: "is back! Use your stored and shared prompts"
*   **Card 2**: `Title`: Learning Sessions | `Date`: July 9, 2026 | `Snippet`: "Boost your productivity: Join our training sessions!"
*   **Card 3**: `Title`: Feature Roadmap | `Date`: April 23, 2026 | `Snippet`: "Curious about new, upcoming features? Latest updates around gemini enterprise..."
*   **Card 4**: `Title`: Feedback Hub | `Date`: January 14, 2026 | `Snippet`: "Have feedback about AskBosch powered by Gemini Enterprise? Share it here!"
