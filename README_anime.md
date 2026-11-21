# Top 300 Anime Database: Cleaning & Exploratory Analysis

A beginner-friendly data analysis project that takes a **raw, messy** spreadsheet of the top 300 anime, cleans it step by step, and explores it with Python, pandas and matplotlib.

> The raw file is intentionally messy (missing values, duplicates, inconsistent text and mixed formats), so you can see every cleaning decision in the notebook and practise the same skills.

## Project goal
Find out which studios, genres and titles stand out in a list of top-rated anime, and turn a disorganised spreadsheet into clean, analysis-ready data along the way.

## Questions answered
1. Which studios earn the largest share of the estimated revenue?
2. How do anime ratings change with release year?
3. Which are the 10 highest-rated anime, and which studios made them?
4. Which genres appear most often?

## Dataset
`Raw & Messy Top 300 Anime Database (Complete).xlsx` - 315 rows and 12 columns in the raw file (300 rows after cleaning).

| Column | Description |
|---|---|
| `Anime Name` | Title of the anime |
| `Studio` | Producing studio(s); several studios are separated by commas |
| `Director's Name` | Director(s) of the anime |
| `Manga Author` | Author of the source material |
| `Rating` | Score out of 10 |
| `Release Year` | Year of release |
| `Completed or Not` | Whether the series has finished airing |
| `Episodes` | Number of episodes |
| `Seasons` | Number of seasons |
| `Genre` | Genre tags separated by commas |
| `Revenue Collection` | Estimated revenue as free text, such as `~$350M+ (Franchise/Media Est.)` |
| `Streaming Platform` | Where the anime can be streamed |

**Problems in the raw data that the notebook fixes**
- Missing values: `Studio` (38), `Rating` (37), `Episodes` (32), `Genre` (25) and `Revenue Collection` (22).
- 15 duplicate anime rows.
- Inconsistent text, such as `bleach`, `MOBILE SUIT GUNDAM...` and `steins;gate`, plus extra spaces.
- Numbers stored as text: ratings like `9.1/10` or `TBD`, and episodes like `24 eps`.
- Release years entered in different formats (`2004`, `2004-05-01`, `04`).
- Revenue written as free text with `~`, `+`, `B`, `M` and `K`.

## Project structure
```
anime-content-analysis/
├── Raw & Messy Top 300 Anime Database (Complete).xlsx
├── Anime_content_organised.ipynb      (main notebook)
├── Anime_content_organised.py         (the same code exported as a script)
├── images/                            (charts saved from the notebook)
├── requirements.txt
└── README.md
```
The notebook looks for the Excel file in its own folder. If you move the file, change `DATA_PATH` in the first code cell.

## How to run
```bash
git clone https://github.com/<your-username>/anime-content-analysis.git
cd anime-content-analysis
pip install -r requirements.txt
jupyter notebook Anime_content_organised.ipynb
```
Then choose **Run All**. The cells depend on each other, so run them from top to bottom.

## Steps followed
1. Load and inspect the data
2. Clean the data
   - 2.1 Standardise text columns (spacing, Title Case, commas)
   - 2.2 Convert Rating and Episodes to numbers
   - 2.3 Extract a clean Release Year
   - 2.4 Remove duplicates and fill missing values
   - 2.5 Convert Revenue Collection to USD millions (`Revenue_USD_M`)
3. Calculate total estimated revenue per studio
4. Visualise: studio revenue share, release year vs. rating, top 10 anime, top genres

## Assumptions
- Missing **Rating**, **Episodes** and **Release Year** values are filled with the column **median**.
- Missing **Studio** values are labelled `Unknown`.
- When an anime appears more than once, the row that has a rating is kept.
- When an anime has several studios, its revenue is **split equally** between them.
- Revenue is converted to **millions of US dollars**.

## Key findings (from this dataset)
- Three studios dominate the estimated revenue: **Sunrise (25.1%)**, **Olm (21.6%)** and **Toei Animation (21.1%)** together make up about two thirds of the total. Ufotable is fourth at 5.4%.
- The highest-rated anime is **Frieren: Beyond Journey's End** (Madhouse, about 9.3), followed by **Fullmetal Alchemist: Brotherhood** (Bones, 9.1) and **Steins;Gate** (White Fox, 9.07).
- **Action** is the most common genre (about 137 anime), followed by **Drama** (about 106) and **Comedy** (about 93).
- Most anime in the list were released from 2000 onwards, and most ratings fall between about 7.5 and 8.5. There is no strong trend between release year and rating, but older titles (before 1990) are all rated 8.2 or lower.

## Charts
![Studio revenue share](images/01_studio_revenue_share.png)
![Release year vs. rating](images/02_release_year_vs_rating.png)
![Top 10 anime by rating](images/03_top10_anime_by_rating.png)
![Top genres](images/04_top_genres.png)

## Limitations
- Revenue figures are rough **franchise-level estimates**, not exact earnings, and splitting them equally between co-producing studios is a simplification.
- Anime with no studio listed are grouped as `Unknown`, which appears as its own slice (2.8%) in the revenue pie chart.
- Filling missing years with the median (2012) creates a vertical line of dots at 2012 in the release year scatter plot.
- The top 10 bar chart starts its axis at 8.5, so small rating differences look larger than they are.

## Tools
Python 3 · pandas · NumPy · matplotlib · openpyxl · Jupyter Notebook · Git and GitHub

## Author
Kapees Walia
2026B01010955
CSE 7
