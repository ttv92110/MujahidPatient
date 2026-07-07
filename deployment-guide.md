# Deployment Guide for Ilm Ul Quran Platform

## Prerequisites
- Python 3.12+
- Google Cloud Project with Sheets API enabled
- Service account JSON key
- (Optional) Redis for caching

## 1. Environment Setup

```bash
# Clone repository
git clone https://github.com/yourorg/ilm-ul-quran.git
cd ilm-ul-quran

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt