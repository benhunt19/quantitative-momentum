# Quantitative-Momentum

This repo is for a quantitative momentum report generating program.

## Process
1. **User Input:** The user provides their required portfolio value and size.
2. **Data Import:** The current S&P 500 constituents are imported via an API.
3. **Data Collection:** For each stock, data is gathered over 1, 3, 6, and 12-month windows using yfinance.
4. **Performance Analysis:** The relative performance of stocks is weighted and ranked.
5. **Report Generation:** An Excel report is generated with a tab for each recommended stock, plots are provided to display their performance.
6. **Automated Emailing:** The report is emailed automatically to a predefined mailing list.

---

## Author

**Ben Hunt**  
[GitHub Profile](https://github.com/benhunt19)  
[LinkedIn](https://www.linkedin.com/in/benjaminrjhunt) 
[benhunt.click](https://benhunt.click/)  

If you have any questions, feedback, or suggestions, feel free to reach out or open an issue in the repository.
