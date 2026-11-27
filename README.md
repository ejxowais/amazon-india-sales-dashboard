# Amazon India Sales Dashboard

**Student:** Owais Ejaz | **Tool used:** Microsoft Excel

## 1. Project overview
I built this workbook for Manoj to share with his manager, Ravi, who comes from a non-technical background and needs a quick, plain view of how the business is doing. Instead of digging through 10,000 raw order rows, Ravi can open the four dashboard tabs to see overall sales and profit trends, top categories and products, losses from returns and cancellations, and performance by payment method, fulfilment channel, and state.

## 2. Dataset
The source file (`data/Amazon_Sales_Data_India.xlsx`) has **10,000 rows and 13 columns** covering orders from **January 1, 2024 to August 30, 2026** (32 months across 5 categories, 25 products, 5 payment methods, 3 fulfilment channels, and 10 Indian states).

Each tab covers one section of the brief:
1. **Overall Sales & Profit:** Total Sales, Total Profit, Total Orders, Total Units Sold, Average Order Value (AOV), Profit Margin %, and Monthly Sales & Profit Trend.
2. **Category & Product Performance:** Category-wise Sales, Profit, Units Sold, and Top 10 Products by Sales.
3. **Order Status & Revenue Loss:** Delivered, Shipped, Returned, and Cancelled order counts, Return Rate %, Cancellation Rate %, and Category-wise Returns & Cancellations.
4. **Payment, Fulfillment & Geo:** Sales and Profit by Payment Method, Sales and Profit by Fulfilment Method, and State-wise Sales, Profit, and Orders.

## 3. Data preparation
- **Checks done:** I checked all 10,000 `Order_ID` entries (`IN-AMD-100000` to `IN-AMD-109999`) for duplicates and confirmed every order ID is unique. I also checked for missing values across sales, profit, quantity, category, and state columns, and verified that `Order_Date` runs cleanly from `2024-01-01` to `2026-08-30` without altering the original date column.
- **Fields I created:** In the hidden `Sales_Data` sheet, I added three helper columns so formulas can group by time cleanly without touching `Order_Date`: `Order_Year` (4-digit year), `Order_Month` (month number 1–12), and `Year_Month` (`YYYY-MM`). I also added hidden `Monthly_Trend` and `Category_Product_Data` sheets to drive the 12-month trend chart and dynamic Top 10 product ranking.
- **How I treated cancelled and returned orders:** I kept all 985 cancelled and returned rows in the dataset instead of deleting them, because removing them would hide the Return Rate % and Cancellation Rate %. In the source data, `Profit_INR` is already `₹0` for cancelled and returned orders, so they do not inflate net profit, while their `Total_Sales_INR` is tracked separately on Tab 3 to show the gross revenue lost.

## 4. KPI definitions
All KPI cards and tables use standard Excel formulas linked directly to `Sales_Data` (no typed or hardcoded numbers):
- **Total Sales** = `=SUMIFS(Sales_Data!$H$2:$H$10001, Sales_Data!$N$2:$N$10001, IF($C$4="All Years",">0",VALUE($C$4)), Sales_Data!$L$2:$L$10001, IF($F$4="All Statuses","*",$F$4))`
- **Total Profit** = `=SUMIFS(Sales_Data!$I$2:$I$10001, Sales_Data!$N$2:$N$10001, IF($C$4="All Years",">0",VALUE($C$4)), Sales_Data!$L$2:$L$10001, IF($F$4="All Statuses","*",$F$4))`
- **Total Orders** = `=COUNTIFS(Sales_Data!$N$2:$N$10001, IF($C$4="All Years",">0",VALUE($C$4)), Sales_Data!$L$2:$L$10001, IF($F$4="All Statuses","*",$F$4))`
- **Total Units Sold** = `=SUMIFS(Sales_Data!$E$2:$E$10001, Sales_Data!$N$2:$N$10001, IF($C$4="All Years",">0",VALUE($C$4)), Sales_Data!$L$2:$L$10001, IF($F$4="All Statuses","*",$F$4))`
- **Average Order Value (AOV)** = `=IF(Total_Orders=0, 0, Total_Sales / Total_Orders)` (`=IF(E7=0,0,A7/E7)`)
- **Profit Margin %** = `=IF(Total_Sales=0, 0, Total_Profit / Total_Sales)` (`=IF(A7=0,0,C7/A7)`)
- **Return Rate %** = `=IF(Total_Orders=0, 0, Returned_Orders / Total_Orders)` (`=IF(B42=0,0,B41/B42)`)
- **Cancellation Rate %** = `=IF(Total_Orders=0, 0, Cancelled_Orders / Total_Orders)` (`=IF(B42=0,0,B40/B42)`)

## 5. Dashboard preview


## 6. Key findings
1. **Steady revenue with stronger order values in 2026:** Total sales across Jan 2024 – Aug 2026 reached **₹15.58 Cr** (₹155,789,894) with **₹3.32 Cr** (₹33,167,008) in profit across 10,000 orders, holding a **21.3% profit margin**. Full-year sales edged up +0.7% from 2024 (₹5.70 Cr) to 2025 (₹5.74 Cr). In 2026 (Jan–Aug), average order value rose to **₹16,209** (up +5.7% from ₹15,338 in 2025), and July 2026 was the highest month on record at ₹64.44 Lakh in sales.
2. **Electronics & Mobiles brings in nearly 80% of sales:** Electronics & Mobiles generated **₹12.36 Cr (79.3% of total sales)** and **₹2.63 Cr in profit**, and accounts for all Top 5 products, led by Laptop Backpack (₹2.65 Cr) and Power Bank 20000mAh (₹2.55 Cr). At the other end, Pantry & Groceries makes up only 1.0% of sales (₹15.27 Lakh), and Matte Lipstick (19.8%) and Ghee 1L (20.2%) have the lowest margins among all 25 products.
3. **Returns and cancellations tie up ₹1.56 Cr in gross sales:** Out of 10,000 orders, 8,147 (81.5%) were Delivered and 868 (8.7%) are Shipped. However, **485 Returned orders (4.9%)** and **500 Cancelled orders (5.0%)** represent ₹79.61 Lakh and ₹76.88 Lakh in lost gross sales with ₹0 realized profit. Beauty & Personal Care has the highest combined loss rate (**10.6%**), while Electronics & Mobiles has the highest return rate (**5.4%**).
4. **UPI and Amazon (FBA) handle most of the volume:** **UPI** leads all payment methods with **₹7.85 Cr (50.4% of sales)**, followed by Credit/Debit Cards (19.9%) and Cash on Delivery (14.8%). On the fulfilment side, **Amazon (FBA)** drives **69.0% of sales** (₹10.75 Cr), while **Seller Flex** posts the highest profit margin at **21.8%** compared to 21.3% for FBA and 20.3% for Merchant (FBM).
5. **Three states generate almost half of total revenue:** **Maharashtra (₹2.96 Cr, 19.0%)**, **Karnataka (₹2.44 Cr, 15.7%)**, and **Delhi (₹2.24 Cr, 14.4%)** together contribute **49.0%** of all sales across the 10 active states.

## 7. Recommendations
1. **Focus return-reduction checks on Electronics & Mobiles:** Because Electronics & Mobiles accounts for 79.3% of sales and has the highest return rate (5.4%, 164 orders), adding pre-dispatch quality checks and clearer product compatibility details on the top 5 electronics items will protect the largest share of profit.
2. **Tighten order confirmation for Beauty and Apparel:** Beauty & Personal Care (5.6% cancellation rate) and Apparel & Fashion (5.7% cancellation rate) lose the highest share of orders before delivery. Faster dispatch timelines and quick confirmation checks on COD orders in these two categories can bring cancellations down.
3. **Shift eligible FBM items to Seller Flex and promote UPI checkout:** Seller Flex earns a 1.5 percentage-point higher margin than Merchant (FBM) (21.8% vs. 20.3%), and UPI already handles half of all sales. Moving high-volume FBM products onto Seller Flex and running UPI checkout offers in top states (Maharashtra, Karnataka, Delhi) can improve overall profitability.

## 8. How to use the dashboard
- Open `dashboard/Amazon_India_Dashboard.xlsx` to view the four tabs: `Overall Sales & Profit`, `Category & Product Performance`, `Order Status & Revenue Loss`, and `Payment, Fulfillment & Geo`. Supporting data and calculation sheets are hidden so the workbook stays clean.
- On every tab, use the **yellow dropdown cells (`C4` and `F4`)** at the top to filter by **Year** (`All Years`, `2024`, `2025`, `2026`) and by **Order Status** or **Category**. All KPI cards, tables, Top 10 rankings, and charts update automatically when you change a selection.

## 9. Limitations
- **Partial year for 2026:** The 2026 records only cover January through August (8 months), so 2026 totals should not be compared directly to full 12-month totals from 2024 and 2025.
- **No customer IDs:** The dataset tracks orders (`Order_ID`) without unique customer identifiers, so repeat purchase rates and customer retention cannot be measured.
- **No itemized cost breakdown:** `Profit_INR` is given as a single net figure per order, without separate columns for shipping charges, ad spend, or reverse-logistics costs on returns.
