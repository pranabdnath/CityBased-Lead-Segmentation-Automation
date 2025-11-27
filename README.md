City-Based Lead Segmentation Automation

Overview

This project automates the process of extracting, filtering, and organizing lead records from an emailed CSV file. The script reads the latest email with a defined subject, processes the attached CSV, separates the data into two groups based on matched city conditions, formats date fields, and updates two sheets with clean results. It also applies a 30-day recurrence rule to mark each Buyer + Project combination with a Y/N flag.

Problem Statement

Manually filtering leads by city, status, and date, then splitting them into two regional sheets was time-consuming and error-prone. Repeated buyers reappeared in shorter periods, causing duplicate actions. The need was to automate filtering, formatting, sheet updates, and applying a 30-day logic.

Solution Summary

Reads the latest email containing the CSV file (subject replaced to "mention your mail subject").

Extracts data, trims it to required columns, and validates city + condition match.

Splits results into two sheets: North and South.

Adds a formatted date column.

Applies 30-day logic to mark entries as Y or N.

Clears old data and writes new records below the header.

Technologies Used

Google Apps Script (JavaScript runtime)
Google Sheets API
Gmail API
