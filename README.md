# La Secrétaire — Voice-Activated AI Productivity Agent

**One-line:** La Secrétaire is a voice-first desktop assistant that summarizes emails, drafts replies, and manages Google Calendar via natural speech.

---

## Key features
- Wake-word activation (e.g., **"Hey Secretary"**) with a glowing listening orb
- Email summarization and multi-tone reply generation (AI-powered)
- Voice-driven Google Calendar event creation and timeline view
- Secure Gmail integration (OAuth) and local credential storage (opt-in)
- Desktop UI (Windows), packaged as an executable for easy distribution

---

## Demo / Screenshots
https://www.linkedin.com/feed/update/urn:li:activity:7340786985738846208/

---

## Requirements (recommended)
- Windows 10/11
- Python 3.10+ (3.12 recommended)
- Google account with Calendar & Gmail
- (Optional) OpenRouter or OpenAI account for LLM access
- Picovoice (Porcupine) AccessKey for wake-word detection

---

## Setup (developer / local)

1. Create a virtual environment
```bash
python -m venv venv
venv\Scripts\activate
