# AI Job Finder

Automatically finds and scores jobs using Groq AI + JSearch API.

## What it does
- Searches for Data Analyst jobs daily
- Scores each job 1-10 against my resume using Groq Llama 3
- Outputs a ranked webpage with one-click apply links
- Saves results to Excel
- Runs automatically every morning via GitHub Actions

## Stack
Python · Groq API · JSearch API · GitHub Actions

## Setup
1. Add RAPIDAPI_KEY and GROQ_API_KEY to GitHub secrets
2. Run manually: python job_finder.py

Built by Nivas Verelli | MS Business Analytics, UT Dallas
