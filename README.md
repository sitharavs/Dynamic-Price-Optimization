# 📈 Price Optimization Macro with Dashboard

Welcome to the **Price Optimization Macro** project! 🎉 This repository contains an Excel-based solution that leverages macros to optimize product prices based on demand, competitor prices, and business logic. Additionally, a dynamic dashboard has been created to visualize insights and trends effectively. 

---

## 📝 Project Overview

This project focuses on **price optimization** by applying business rules and logic to dynamically adjust prices. The macro processes data and generates optimized prices while providing actionable insights through visualizations.

### 🔧 Business Logic

The macro uses the following conditions to adjust prices:

| **Condition**                     | **Action**                | **Priority** |
|------------------------------------|---------------------------|--------------|
| If demand is high                 | Increase price by 5%      | 1            |
| If competitor's price is lower    | Decrease price by 3%      | 2            |
| If demand is low                  | Decrease price by 7%      | 3            |
| If profit margin is below 20%     | Increase price by 2%      | 4            |

---

## 📊 Features

1. **Macro-Based Price Optimization**:
   - Automatically adjusts product prices based on demand, competitor pricing, and profit margins.
   - Outputs optimized prices in a new Excel sheet.

2. **Interactive Dashboard**:
   - **Line Chart**: Visualizes old vs. new prices over time.
   - **Price vs. Demand Chart**: Highlights the relationship between pricing and demand.
   - **Demand vs. Market Season Chart**: Shows demand trends across different market conditions like holidays, summer sales, Black Friday, etc.

3. **Synthetic Data Generation**:
   - The dataset used for this project was generated synthetically using a Python script, ensuring flexibility for testing and analysis.

---

## 📂 Repository Contents

This repository includes the following files:

- **Macro File (`Dynamic_Price_Optimization.xlsm`)**: The main Excel file containing the VBA macro and dashboard.
- **Dataset (`target_synthetic_product_data.xlsx`)**: The synthetic dataset used for testing.
- **VBA Script (`Dynamic Pricing Optimization-VBA Code`)**: The standalone VBA code for the macro.
- **Python Script (`Product_Prices_Historical_Data.ipynb`)**: The Python script used to generate the synthetic dataset.

---

## 🚀 How to Use

### Prerequisites
- Microsoft Excel (with support for macros enabled).
- Python (optional, if you want to regenerate the dataset).

### Steps
1. Clone this repository:
   ```bash
   git clone https://github.com/sitharavs/Dynamic-Price-Optimization.git
   ```
2. Open the `Dynamic_Price_Optimization.xlsm` file in Microsoft Excel.
3. Enable macros if prompted.
4. Run the macro via the dashboard to generate optimized prices.
5. Explore insights through the interactive charts on the dashboard.

---

## 🛠️ Technologies Used

- **Excel VBA Macros**: For implementing price optimization logic.
- **Python (Pandas, NumPy)**: For generating synthetic datasets.
- **Excel Charts & Dashboards**: For visualizing key insights.

---

## 📈 Sample Visualizations

Here’s what you can expect from the dashboard:

1. **Old Price vs New Price (Line Chart)**:
   - Tracks how prices are adjusted over time.

2. **Price vs Demand (Scatter Plot)**:
   - Illustrates how demand correlates with pricing.

3. **Demand vs Market Seasons (Bar Chart)**:
   - Highlights seasonal demand trends during events like Black Friday, holidays, etc.

---

## 🤝 Contributing

Contributions are welcome! If you have suggestions or improvements, feel free to open an issue or submit a pull request.

---


## 📬 Contact

If you have any questions or feedback, feel free to reach out via email or open an issue in this repository.

Happy optimizing! 🎉✨
