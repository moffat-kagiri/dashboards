# Advanced Power BI Data Transformation Dashboard

This project showcases a high-performance Power BI dashboard built with a focus on efficient and reliable data processing. It demonstrates advanced M-language techniques in Power Query to solve complex data preparation challenges.

## 🚀 Technical Highlights

*   **Optimized Data Model:** Implements performant table buffering and type casting for faster refresh times.
*   **Advanced Multi-Key Lookups:** Replaces inefficient row-by-row searches with robust bulk merge operations for data validation and status flagging.
*   **Data Integrity:** Features automated handling of null values and default assignments to ensure consistent reporting outputs.

## 🛠️ Key Techniques Demonstrated

*   `Table.NestedJoin` & `Table.ExpandTableColumn` for SQL-like left joins.
*   `Table.Buffer` and `Table.TransformColumns` for in-memory performance optimization.
*   `Table.ReplaceValue` for robust data cleaning and default value logic.

## 📁 Repository Structure
```
├── Advanced_Dashboard.pbit # Power BI Template file (no data)
├── /SampleData
│ └── sample_data.xlsx # Anonymized sample dataset
├── /Screenshots
│ └── dashboard_preview.png # Dashboard overview
└── README.md # This file
```

## 🚦 Getting Started

1.  **Clone** this repository.
2.  Download the `Advanced_Dashboard.pbit` file.
3.  Open the template in Power BI Desktop and connect it to the `sample_data.xlsx` file when prompted.
4.  Refresh the report to see the data model and transformations in action.

## 📄 License

This project is provided for educational and portfolio purposes.
