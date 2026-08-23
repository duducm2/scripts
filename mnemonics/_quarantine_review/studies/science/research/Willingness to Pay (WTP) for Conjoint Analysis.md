# Willingness to Pay (WTP) for Conjoint Analysis

## Introduction to Willingness to Pay 
Since the inception of conjoint analysis, determining how much respondents are willing to pay for specific features has been a primary goal. However, achieving accurate WTP estimations has historically been difficult. Traditional methods have consistently overstated these values over the decades. 

## Traditional Approaches to WTP

### 1. The Algebraic Approach
This is the most common historical method for estimating WTP.
* **Mechanism:** It relies on simple algebra to infer value based on utility points.
* **Example:** In a conjoint study with "feature" and "price" attributes, if $100 is worth one utility point more than $200, and Feature A is worth two utility points compared to Feature B, the approach infers that Feature A is worth $200 more than Feature B.
* **Flaws:** The resulting WTP is typically too high and unbelievable. It fails to consider market competition, assuming the firm holds a monopoly on the feature and that buyers cannot walk away (ignoring the "none" option). Furthermore, it averages responses across all participants rather than isolating those on the cusp of choosing the product.

### 2. The Two Product Simulation Approach
This method utilizes a market simulator built from the conjoint analysis study.
* **Mechanism:** It begins by assuming only two identical products exist in the market, each holding a 50 percent share. One product is modified to include an enhanced feature, which increases its share. The price of this enhanced product is then incrementally raised until its share drops back to 50 percent.
* **Result:** The required price increase to return the share to 50 percent is considered the WTP. 
* **Flaws:** Similar to the algebraic approach, this method assumes a lack of competition and overstates WTP. 

## The Gilligan's Island Problem
To illustrate the flaws of assuming a monopoly in WTP calculations, Sawtooth Software uses the "Gilligan's Island problem" analogy. 
* **The Scenario:** Mr. Howell, a wealthy character stranded on an island, would gladly pay over a million dollars to a single boat offering rescue. His theoretical ability and willingness to pay are exceptionally high in a monopoly.
* **The Reality of Competition:** If a second boat arrives offering the same rescue for $500, Mr. Howell will choose the $500 option. The realistic market WTP is therefore $500, proving that competitive alternatives drastically reduce actual WTP.

## The Competitive Simulation Approach
To correct the overestimation issues, a more realistic approach integrates a competitive market framework.
* **Mechanism:** This approach evaluates a base product against a rich set of competitors (typically five to seven market offerings) and includes a "none" alternative, allowing buyers to opt out entirely. 
* **Process:** The simulation records the initial market share against these competitors. The target product is then enhanced, and the price differential required to drive the increased share back to the original baseline is calculated. 
* **Impact:** By focusing on respondents on the cusp of making a choice and accounting for competitor substitutes, this method yields WTP estimates that are an average of 20 percent lower than traditional approaches. In some datasets, it cuts the WTP estimate by half.

## Sampling of Scenarios
Clients often struggle to define a fixed, static base case scenario for their firm and all competitors. The "Sampling of Scenarios" approach resolves this difficulty.
* **Mechanism:** The software executes hundreds of random draws, creating diverse attribute specifications for the client's product and various competitor reactions at different price points and feature combinations. 
* **Calculation:** The system computes the WTP for each random scenario draw using the competitive simulation method and then takes the median WTP across all draws.
* **Advantages:** This eliminates the need to fix a static competitive landscape. It also allows researchers to impose logical rules on the simulation. 
* **Examples of Logical Rules:** * **Patents:** If a firm holds a patent, the simulator can restrict competitors from offering that specific feature level. In a provided dataset example, WTP with a patent was $109, but dropped to $58 if competitors were allowed to offer the same feature.
    * **Brand Exclusivity:** The simulator can restrict specific brand names. For example, WTP for a feature might be $66 under the Sony brand, but only $54 under the JVC brand.

***

## Exact Sentences for Story Reuse

* "This willingness to pay via the algebraic approach is just typically too high and just not believable because it doesn't consider competition."
* "Both of those previous two approaches just assumed that the firm has a monopoly and that there's no additional competitive offerings and it overstates willingness to pay."
* "A competitive market simulation approach would include a relevant rich set of offerings covering the majority of the market offerings in your space and potentially the none alternative."
* "In the sampling of scenarios we've programmed our software to make hundreds of random draws of attribute specifications for your clients product and also the potential reactions of myriad reactions of the competitors in the marketplace."
* "You can furthermore impose logical rules on these draws of your firm's product and competitors... for example your product has a patent on a particular feature and the alternatives cannot take on that feature."