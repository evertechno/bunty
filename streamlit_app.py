# streamlit_video_to_gif.py
import streamlit as st
import cv2
import numpy as np
from PIL import Image
import io
import base64
import tempfile
import os
from typing import List, Tuple

st.set_page_config(page_title="Video → Email-compatible GIF", layout="centered")

st.title("🎥 ➜ GIF (email-ready) — Streamlit")
st.write(
    "Upload a short video and convert it to an optimized GIF suitable for embedding in emails. "
    "This app avoids external ffmpeg and uses OpenCV + Pillow."
)

# --- Helper functions ---
def save_temp_video_file(uploaded_file) -> str:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1])
    tmp.write(uploaded_file.getbuffer())
    tmp.flush()
    tmp.close()
    return tmp.name

def read_video_metadata(path: str) -> Tuple[float, int, int]:
    cap = cv2.VideoCapture(path)
    if not cap.isOpened():
        raise RuntimeError("Could not open video file.")
    fps = cap.get(cv2.CAP_PROP_FPS) or 25.0
    frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
    width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH) or 0)
    height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT) or 0)
    cap.release()
    duration = frame_count / fps if fps else 0
    return duration, fps, (width, height)

def video_to_frames(path: str, start: float, end: float, target_fps: float, max_width: int) -> List[Image.Image]:
    cap = cv2.VideoCapture(path)
    if not cap.isOpened():
        raise RuntimeError("Could not open video file.")
    orig_fps = cap.get(cv2.CAP_PROP_FPS) or 25.0
    frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
    duration = frame_count / orig_fps if orig_fps else 0

    # clamp
    start = max(0.0, min(start, duration))
    end = max(start, min(end, duration))
    if end <= start:
        end = min(start + 5.0, duration)  # fallback

    # compute frame indices to capture
    start_frame = int(start * orig_fps)
    end_frame = int(end * orig_fps)

    # step to achieve target_fps
    step = max(1, int(round(orig_fps / target_fps))) if target_fps > 0 else 1

    frames = []
    cap.set(cv2.CAP_PROP_POS_FRAMES, start_frame)
    frame_idx = start_frame
    while frame_idx <= end_frame:
        ret, frame = cap.read()
        if not ret:
            break
        if (frame_idx - start_frame) % step == 0:
            # convert BGR -> RGB
            rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            pil = Image.fromarray(rgb)
            # resize preserving aspect ratio if needed
            if max_width and pil.width > max_width:
                ratio = max_width / pil.width
                new_h = int(pil.height * ratio)
                pil = pil.resize((max_width, new_h), Image.LANCZOS)
            frames.append(pil.convert("RGBA"))
        frame_idx += 1

    cap.release()
    return frames

def frames_to_gif_bytes(frames: List[Image.Image], duration_ms: int, loop: int = 0, optimize: bool = True) -> bytes:
    if not frames:
        raise ValueError("No frames to encode.")
    # Convert frames to P mode (palette) for smaller GIFs, maintain transparency handling
    first, rest = frames[0].convert("RGBA"), [f.convert("RGBA") for f in frames[1:]]
    # Convert RGBA -> P using adaptive palette, keep preserve transparency by using a white background for quantization
    # To reduce artifacts, paste onto background then quantize
    bg = Image.new("RGBA", first.size, (255, 255, 255, 255))
    composed_frames = []
    for fr in [first] + rest:
        tmp = Image.alpha_composite(bg, fr).convert("RGB")  # flatten
        composed_frames.append(tmp)

    # Save to BytesIO via Pillow
    bio = io.BytesIO()
    # duration_ms per frame
    composed_frames[0].save(
        bio,
        format="GIF",
        save_all=True,
        append_images=composed_frames[1:],
        duration=duration_ms,
        loop=loop,
        optimize=optimize,
        disposal=2,
    )
    bio.seek(0)
    return bio.read()

def make_data_uri(gif_bytes: bytes) -> str:
    b64 = base64.b64encode(gif_bytes).decode("ascii")
    return f"data:image/gif;base64,{b64}"

# --- UI ---

uploaded = st.file_uploader("Upload video (mp4, mov, avi, webm...) — keep under 30 MB for best results", type=["mp4","mov","avi","webm","mkv"], accept_multiple_files=False)

if uploaded is None:
    st.info("Upload a short video to begin. Recommended: <= 10 seconds, <= 600×? width for email compatibility.")
    st.stop()

# Save to temp file (opencV requires a path)
with st.spinner("Saving upload..."):
    video_path = save_temp_video_file(uploaded)

try:
    duration, orig_fps, (w, h) = read_video_metadata(video_path)
except Exception as e:
    st.error(f"Couldn't read video metadata: {e}")
    os.unlink(video_path)
    st.stop()

st.write(f"**Filename:** {uploaded.name} — **Duration:** {duration:.2f}s — **Resolution:** {w}×{h} — **FPS:** {orig_fps:.2f}")

# Controls: start/end, fps, width, max size
col1, col2 = st.columns(2)
with col1:
    start_time = st.number_input("Start time (sec)", min_value=0.0, max_value=duration, value=0.0, step=0.5, format="%.2f")
    end_time = st.number_input("End time (sec)", min_value=0.0, max_value=duration, value=min(5.0, duration), step=0.5, format="%.2f")
with col2:
    target_fps = st.select_slider("Target FPS (lower → smaller file)", options=[1,2,3,5,6,8,10,12,15], value=6)
    max_width = st.selectbox("Max width (px) for GIF (email-friendly)", options=[320,400,480,600,800], index=0)

loop = st.checkbox("Loop GIF (set loop=0 for infinite)", value=True)
loop_val = 0 if loop else 1

duration_per_frame_ms = int(1000 / target_fps)

st.markdown("---")
st.write("Advanced / safety")
max_duration_allowed = st.slider("Maximum extracted duration (sec) — to keep GIF tiny", min_value=1, max_value=30, value=10)
if (end_time - start_time) > max_duration_allowed:
    st.warning(f"Selected clip is longer than {max_duration_allowed}s. It will be truncated to keep GIF size manageable.")
    end_time = start_time + max_duration_allowed

convert_button = st.button("Convert to GIF")

if convert_button:
    try:
        with st.spinner("Extracting frames (OpenCV)..."):
            frames = video_to_frames(video_path, start_time, end_time, target_fps, max_width)
        if not frames:
            st.error("No frames extracted. Try adjusting start/end or check video format.")
        else:
            st.success(f"Extracted {len(frames)} frames.")
            with st.spinner("Encoding GIF (Pillow)..."):
                gif_bytes = frames_to_gif_bytes(frames, duration_ms=duration_per_frame_ms, loop=loop_val, optimize=True)

            gif_size_kb = len(gif_bytes) / 1024
            st.write(f"### Result GIF — {gif_size_kb:.1f} KB")

            # Preview
            st.image(gif_bytes, format="GIF", caption="Preview GIF")

            # Provide download button
            st.download_button(
                label="Download GIF",
                data=gif_bytes,
                file_name=os.path.splitext(uploaded.name)[0] + ".gif",
                mime="image/gif"
            )

            # Data URI for embedding in HTML / iframe
            data_uri = make_data_uri(gif_bytes)

            st.markdown("#### Render inside HTML `img` (example)")
            html_img = f"""<div>
<img src="{data_uri}" alt="embedded gif" style="max-width:100%;height:auto;border:1px solid #ddd;border-radius:6px;" />
</div>"""
            st.components.v1.html(html_img, height=min(600, int(max_width * 0.8) + 80), scrolling=True)

            st.markdown("#### Render inside an `iframe` (example)")
            # iframe srcdoc with the same image
            iframe_srcdoc = f"""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>GIF in iframe</title>
</head>
<body style="margin:0;padding:10px;background:#f8f9fb;display:flex;align-items:center;justify-content:center;height:100%;">
<img src="{data_uri}" alt="gif" style="max-width:100%;height:auto;" />
</body>
</html>
"""
           # Escape double quotes safely before embedding in f-string
safe_srcdoc = iframe_srcdoc.replace('"', "&quot;")

iframe_html = (
    f'<iframe srcdoc="{safe_srcdoc}" '
    f'style="width:100%;height:420px;border:1px solid #ddd;border-radius:6px;"></iframe>'
)

st.components.v1.html(iframe_html, height=420)


    except Exception as e:
        st.error(f"Conversion failed: {e}")
    finally:
        try:
            os.unlink(video_path)
        except Exception:
            pass

# show example: small fallback auto-convert for very small uploads
st.caption("Tip: For email use, keep GIF short (3–7s), low fps (6–12), and ≤ 600px width.")

