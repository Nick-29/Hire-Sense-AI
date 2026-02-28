# Hire-Sense-AI
Hire-Sense AI automates recruitment: resume parsing, skill matching, risk detection (hidden text, job hopping, gaps >2y). Generates ranked shortlists, triggers Twilio voice calls with IVR (Press 1/2), and sends automated emails. Streamlines hiring from screening to engagement.

Problem Statement:
Recruitment teams spend countless hours manually screening resumes, comparing skills with job requirements, and coordinating candidate communication. This leads to:

Inconsistent candidate evaluation across team members

Slow shortlisting and delayed outreach

Lack of transparency in why candidates are selected or rejected

Poor candidate experience due to delayed responses

These inefficiencies increase time‑to‑hire and risk losing top talent to faster competitors.

Solution Overview:
Hire‑Sense AI automates the entire recruitment workflow:

AI Resume Parser – Extracts name, email, phone, skills, and experience from .docx resumes.

Job Requisition Matcher – Compares candidate profiles with job descriptions using hybrid exact‑phrase detection + semantic similarity (Sentence‑Transformers).

Risk Detection – Flags resumes containing hidden text (white‑on‑white), job hopping (more than 3 jobs in the last 2 years), or employment gaps longer than 2 years.

Automated Calling – Uses Twilio Studio to call shortlisted candidates with a personalised IVR (Press 1 = Interested, Press 2 = Not Interested). Responses are captured in real time.

Email Notifications – Sends congratulations emails to shortlisted candidates and thank‑you emails to non‑shortlisted candidates via Gmail SMTP.

All steps are orchestrated by a master script, making the entire process fully autonomous.

Folder structure:

Hire-Sense-AI/
├── HiresenseAI.py                # Master orchestrator – run this to execute the whole pipeline
├── AIresumereader.py             # Resume parsing, matching, and ranking
├── risk_filter.py                # Hidden text, job hopping, and gap detection
├── calling.py                    # Twilio auto‑caller (uses Shortlisted_clean.xlsx)
├── emailautomation.py            # Sends emails to shortlisted/non‑shortlisted candidates
├── requirements.txt              # Python dependencies with pinned versions
├── .env.example                  # Template for environment variables (copy to .env)
├── README.md                     # This file
├── .gitignore                    # Excludes sensitive files from Git
├── Resumes/                      # **Place all candidate resumes (.docx) here**
└── (Excel files will be generated here after running)


