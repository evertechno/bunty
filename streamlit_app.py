import os
import av
import base64
import streamlit as st
from streamlit_webrtc import webrtc_streamer, VideoProcessorBase, RTCConfiguration
from google import genai
from google.genai import types

# --- Setup Gemini client ---
def get_gemini_client():
    return genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

# --- Bunty AI: Generate Response ---
def bunty_response(prompt: str):
    client = get_gemini_client()
    model = "gemini-2.5-flash"
    contents = [
        types.Content(
            role="user",
            parts=[types.Part.from_text(text=prompt)],
        ),
    ]
    generate_content_config = types.GenerateContentConfig(
        thinking_config=types.ThinkingConfig(thinking_budget=0),
    )

    output = ""
    for chunk in client.models.generate_content_stream(
        model=model, contents=contents, config=generate_content_config
    ):
        if chunk.text:
            output += chunk.text
    return output

# --- Streamlit App ---
st.set_page_config(page_title="Bunty AI Agent", layout="wide")

st.title("🤖 Bunty - Your AI Helper with Screen Share")

# Step 1: User asks a question
user_query = st.text_area("💬 Ask Bunty a Question:", placeholder="e.g. How do I debug this error?")

if st.button("Ask Bunty"):
    if user_query.strip():
        with st.spinner("Bunty is thinking..."):
            response = bunty_response(user_query)
        st.success("Bunty says:")
        st.write(response)

# Step 2: Screen Sharing (WebRTC)
st.markdown("### 📺 Share your screen with Bunty")
rtc_config = RTCConfiguration({"iceServers": [{"urls": ["stun:stun.l.google.com:19302"]}]})

class ScreenProcessor(VideoProcessorBase):
    def recv(self, frame):
        # Convert frame to image for AI (Optional: OCR, vision model)
        img = frame.to_ndarray(format="bgr24")
        # Currently, we just return frame as is
        return av.VideoFrame.from_ndarray(img, format="bgr24")

webrtc_streamer(
    key="screen-share",
    mode="recvonly",  # viewer mode
    rtc_configuration=rtc_config,
    video_processor_factory=ScreenProcessor,
    media_stream_constraints={"video": True, "audio": False},
)

st.info("👉 Start screen sharing, and Bunty will analyze what’s happening in real time!")

# Future scope: Capture frames, send to Gemini Vision model, and combine with text queries
