import os
from dotenv import load_dotenv

load_dotenv()

# OpenRouter API configuration
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
LLM_MODEL = "openai/gpt-4o-mini"  # You can also use "anthropic/claude-3.5-sonnet", "meta-llama/llama-3.1-8b-instruct", etc.
OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"