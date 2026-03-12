import json
import os
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

def parse_command(command):

    prompt = f"""
Convert this command into JSON.

Command: {command}

Output format:

{{
"app": "",
"action": "",
"parameters": {{}}
}}

Example:

delete first 5 rows in excel →

{{
"app": "excel",
"action": "delete_rows",
"parameters": {{"rows":5}}
}}
"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role":"user","content":prompt}]
    )

    text = response.choices[0].message.content

    return json.loads(text)