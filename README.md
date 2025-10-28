# NOBO List – Investor Insights (Demo)

This project is a **demo version** of an interactive dashboard designed to visualize investor data in Canada, using anonymized or synthetic data inspired by real-world patterns.

The goal is to demonstrate how geographic, demographic, and market variables can be combined to support **strategic communication** and **investor engagement analysis**.

 **Live Demo:** [https://jaquelinefarah.github.io/nobo_demo/](https://jaquelinefarah.github.io/nobo_demo/)  
 **Repository:** [https://github.com/jaquelinefarah/nobo_demo](https://github.com/jaquelinefarah/nobo_demo)

---

##  Project Overview

The NOBO (Non-Objecting Beneficial Owners) List dashboard provides an exploratory view of investor data, combining:
- **Geographic segmentation** (province, city, FSA)
- **Socioeconomic indicators** (income, age)
- **Investor tiers and clusters**
- **Interactive map-based filtering**

Each map allows customized exploration by region, investor tier, or cluster, helping identify areas of potential engagement and concentration.

---

##  Technologies Used

- **HTML, CSS & JavaScript** – front-end structure and interactivity  
- **Google Maps API** – geolocation and marker visualization  
- **JSON** – lightweight data storage for demo investors  
- **Python (Pandas, GeoTools)** – data cleaning and preprocessing (offline)  
- **GitHub Pages** – hosting for the interactive demo

---

## Repository Structure
```
nobo_demo/
├── index.html # Main dashboard entry
├── investors_map_overview.html
├── investors_map_clusters.html
├── britishcolumbia_map.html
├── data/
│ └── nobo_demo_bc.json # Sample anonymized data
├── assets/
│ └── logo.png # (optional visual asset)
└── config.js # Google Maps API key loader
```
---

## Data Disclaimer

All datasets used in this demo are **synthetic or anonymized**.  
Some names or references may resemble real entities but are included **solely for illustrative purposes**.  
No real investor or private data is disclosed or represented.

---

## Author

**Jaqueline Farah**  
Data Analyst & Visualization Designer  
📍 Vancouver, Canada  
🔗 [LinkedIn](https://www.linkedin.com/in/jaquelinefarah) • [GitHub](https://github.com/jaquelinefarah)

---

## License

This repository is published for **educational and demonstrative purposes only**.  
Feel free to explore, fork, or reference it, with proper attribution.
