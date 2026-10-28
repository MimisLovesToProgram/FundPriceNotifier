## 📈 Fund Price Notifier (Checker)

A project my dad requested so he could get notified when significant price movements happened in certain funds found in an excel file.
Scrapes some websites for the prices, tracks the maximum prices, looks for important changes, mails the results.

No more constant looking through many websites and manual price tracking!

# How it works (in more detail)

- Reads a list of funds from an Excel file
- **Searches for their prices** on numerous websites using Selenium
- Saves their max prices, and checks movements of their price relative to their max
- The price is split into 20 equal parts, representing 5% of the max price. **If a fund falls below one of those 5% zones, the program mails the users**. 

---
*More updates to come as the project expands in the future.*
