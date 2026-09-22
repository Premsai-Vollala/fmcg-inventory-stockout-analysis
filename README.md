# Quick-Commerce FMCG Inventory & Stockout Analysis

## 📌 Executive Summary
In rapid fulfillment quick-commerce (10–15 minute grocery delivery), dark stores face extreme order spikes and inventory volatility. Ineffective replenishment logic leads directly to missed revenue from stockouts alongside capital erosion from high-spoilage perishables.

This project delivers an automated inventory health monitoring model and decision framework using Advanced Excel and Power Query. It evaluates **20,000 SKU distribution events** across dark stores to detect stockout vulnerabilities, simulate safety stock thresholds, and prioritize reorder queues by revenue impact.

---

## 🛠️ Technical Toolkit & Methodologies
* ETL Pipeline: Excel Power Query (automated schema normalization, null handling, type conversion, unpivoting fulfillment logs).
* Inventory Modeling: Dynamic Reorder Point (ROP) formulas, Safety Stock modeling under variable lead times, and SKU Stockout Frequency calculations.
* Excel Functions: XLOOKUP, INDEX-MATCH, SUMIFS, nested IFS, dynamic arrays (FILTER, UNIQUE, SORT), and LET functions for scalable calculation logic.
* Analytical Frameworks: ABC-XYZ Analysis (Value contribution vs. Demand predictability), Days of Inventory Outstanding (DIO), and Lost Revenue Imputation.

---

## 📐 Core Analytical Calculations & Formulas

### 1. Dynamic Reorder Point (ROP)

$$\text{Reorder Point} = (\text{Average Daily Demand} \times \text{Lead Time}) + \text{Safety Stock}$$

$$\text{Safety Stock} = Z \times \sigma_d \times \sqrt{L}$$

* Where $Z = 1.65$ (95% service level factor), $\sigma_d$ is the standard deviation of daily demand, and $L$ is vendor replenishment lead time in days.

```excel
=LET(
    AvgDemand, [@avg_daily_demand],
    LeadTime, [@lead_time_days],
    StdDevDemand, [@std_dev_demand],
    ZScore, 1.65,
    SafetyStock, ROUNDUP(ZScore * StdDevDemand * SQRT(LeadTime), 0),
    ROUNDUP((AvgDemand * LeadTime) + SafetyStock, 0)
)
```
### 2. Stockout Risk Classification
Categorizes SKUs into urgent operational action states based on current stock vs. lead time consumption:

Excel
=IFS(
    [@current_stock] = 0, "Critical - Stockout",
    [@current_stock] <= [@SafetyStock], "High Risk - Below Safety Buffer",
    [@current_stock] <= [@ReorderPoint], "Reorder Triggered",
    TRUE, "Optimal Stock"
)
### 3. Estimated Stockout Lost Revenue
Estimates the gross revenue impact during fulfillment downtime:

Excel
=[@stockout_hours_last_30d] * (([@avg_daily_demand] / 16) * [@unit_sp_inr])

### 📊 Business Insights & Recommendations
Perishable vs. Non-Perishable Divergence: Perishable lines (8,000 SKUs across Dairy and Bakery) averaged 6.84 stockout hours per SKU over the last 30 days, generating
an estimated gross revenue loss of ₹30.76 Lakhs. Dairy recorded the highest overall downtime at 6.90 hours per SKU, driven by stringent shelf-life constraints and compressed reorder cycles.

Lead-Time Volatility Impact: Over 81% of catalog volume operates within tight 1 to 3 day replenishment windows (16,237 SKUs), accounting for ₹74.26 Lakhs (79%) of 
total stockout losses. SKUs with 2 to 3 day lead times experienced the greatest operational exposure, highlighting the need for dynamic safety buffers tailored to
vendor replenishment cycles rather than static assumptions.

SKU Rationalization & Prioritization: Stockout downtime across all 20,000 SKUs created an estimated ₹94.01 Lakhs in gross revenue leakage over 30 days. Prioritizing automated reorder triggers on top-velocity SKUs recovers approximately ₹29.04 Lakhs in high-probability fulfillment losses.

### 📂 Repository Contents
├── quick_commerce_fmcg_dark_store_20000.csv  # 20,000-row dark store inventory & transaction log
├── fmcg_inventory_model.xlsx                 # Advanced Excel workbook with Power Query & data model
└── README.md                                 # Project documentation and analytical methods
