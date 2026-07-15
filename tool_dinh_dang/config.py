import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Gemini API Configurations
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# JSLHU Rules and Standards
JSLHU_RULES = {
    "max_title_words": 20,
    "min_abstract_words": 150,
    "max_abstract_words": 250,
    "max_keywords": 5,
    "min_references": 10,
    "max_references": 15,
    "min_new_references": 3,
    "new_ref_years_limit": 5,  # Mới trong 5 năm gần nhất
    "required_email_domain": "@lhu.edu.vn"
}
