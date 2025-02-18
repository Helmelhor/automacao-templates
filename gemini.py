# gemini_script.py
from google import genai
client = genai.Client(api_key="AIzaSyAFdxMkokQUfbUvFbdxV30NDd3x9qR2Rk0")

response = client.models.generate_content(
    model="gemini-2.0-flash",
    contents="me ajude a planejar uma viagem",
)

print(response.text)