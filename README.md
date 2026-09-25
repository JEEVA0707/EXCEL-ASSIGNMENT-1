📊 Excel Data Exploration & Analysis

📌 Project Overview

This project demonstrates my foundational Data Analysis skills using Microsoft Excel. The analysis was performed on a product dataset containing information such as Product ID, Product Name, Brand, Quantity, Category, and Price.

The objective of this project was to explore the dataset using Excel formulas and functions, perform basic statistical calculations, apply conditional logic, and extract useful information from text-based Product IDs.

This project is part of my journey toward becoming a Data Analyst and building a practical data analytics portfolio.

---

🗂️ Dataset

The dataset contains 34 product records with the following attributes:

Column| Description
Product ID| Unique identifier containing day, month, and country code
Product Name| Name of the product
Brand Name| Product brand
Quantity| Number of units
Category| Product category
Price ($)| Product price in USD

Example Product ID

"28-JAN-US"

The Product ID contains multiple pieces of information:

- "28" → Day
- "JAN" → Month
- "US" → Country Code

---

🎯 Objectives

The main objectives of this project were to:

- Explore and summarize product price data
- Calculate basic statistical measures
- Apply logical functions using "IF"
- Perform conditional aggregation using "SUMIF" and "COUNTIF"
- Extract information from Product IDs using "LEFT", "RIGHT", and "MID"
- Build a foundation in Excel-based data analysis

---

🔍 Analysis Performed

1. Basic Data Exploration

Total Price

The total price of all products was calculated using the "SUM" function.

=SUM(D2:D35)

Result: "$10,100"

Number of Products

The number of products was calculated using the "COUNT" function.

=COUNT(D2:D35)

Result: "34 products"

Average Price

The average product price was calculated using the "AVERAGE" function.

=AVERAGE(D2:D35)

Result: "$297.06"

---

2. Minimum and Maximum Price

The minimum and maximum product prices were identified using Excel's "MIN" and "MAX" functions.

Minimum Price

=MIN(D2:D35)

Result: "$30"

Maximum Price

=MAX(D2:D35)

Result: "$1,000"

---

3. Logical Function – IF

A new column called Price Range was created to categorize products based on their price.

Business Rule

Condition| Classification
Price >= $500| High Price
Price < $500| Standard Price

Excel Formula

=IF(D2>=500,"High Price","Standard Price")

The formula was then copied down for all products.

This demonstrates the use of conditional logic in Excel.

---

4. Conditional Functions – SUMIF and COUNTIF

Total Price of Electronics Products

The "SUMIF" function was used to calculate the total price of products belonging to the Electronics category.

=SUMIF(F2:F35,"Electronics",D2:D35)

Result: "$8,050"

Products Priced Below $100

The "COUNTIF" function was used to count products with a price below "$100".

=COUNTIF(D2:D35,"<100")

Result: "11 products"

---

🔤 5. Text Functions – LEFT, RIGHT and MID

The Product ID contains useful information that can be extracted using Excel text functions.

For example:

"28-JAN-US"

Day

The first two characters were extracted using the "LEFT" function.

=LEFT(A2,2)

Result: "28"

Country Code

The last two characters were extracted using the "RIGHT" function.

=RIGHT(A2,2)

Result: "US"

Month

The month was extracted from characters 4 to 6 using the "MID" function.

=MID(A2,4,3)

Result: "JAN"

Resulting Columns

Product ID| Day| Month| Country Code
28-JAN-US| 28| JAN| US
15-FEB-US| 15| FEB| US
03-MAR-US| 03| MAR| US
11-APR-US| 11| APR| US

---

📈 Key Results

Analysis| Result
Total Product Price| $10,100
Number of Products| 34
Average Price| $297.06
Minimum Price| $30
Maximum Price| $1,000
Electronics Total Price| $8,050
Products Below $100| 11

---

🛠️ Excel Functions Used

The following Excel functions were used in this project:

- "SUM"
- "COUNT"
- "AVERAGE"
- "MIN"
- "MAX"
- "IF"
- "SUMIF"
- "COUNTIF"
- "LEFT"
- "RIGHT"
- "MID"

---

💡 Skills Demonstrated

Through this project, I practiced the following data analysis skills:

Excel Data Analysis

- Data exploration
- Data summarization
- Descriptive statistics
- Conditional calculations

Logical Functions

- Applying business rules with "IF"
- Categorizing data based on conditions

Text Manipulation

- Extracting substrings
- Parsing structured IDs
- Creating new analytical columns

Data Interpretation

- Understanding product pricing
- Identifying category-level totals
- Finding products based on price conditions

---

📁 Project Structure

Excel-Data-Exploration/
│
├── Excel Assignment 1 - Data Exploration-1.xlsx
└── README.md

---

🚀 Learning Outcome

This project helped me strengthen my understanding of Excel as a data analysis tool.

I learned how to:

1. Summarize datasets using Excel functions.
2. Calculate important descriptive statistics.
3. Apply logical conditions to categorize data.
4. Perform conditional aggregation.
5. Extract meaningful information from structured text.
6. Convert raw dataset fields into useful analytical columns.

These are foundational skills that I plan to build upon as I continue learning SQL, Power BI, Python, and advanced data analytics techniques.

---

👨‍💻 About This Portfolio

I am building this repository as part of my journey toward becoming a Data Analyst.

My goal is to continuously work on practical projects and demonstrate my ability to:

Collect → Clean → Explore → Analyze → Visualize → Communicate Data

More data analytics projects will be added as I continue developing my skills.

---

⭐ Tools Used

- Microsoft Excel
- GitHub
- Markdown

---

📌 Project Status

Completed ✅
