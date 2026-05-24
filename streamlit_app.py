from __future__ import annotations

import base64
import datetime as dt
import hmac
import html
import time
from typing import Any, Dict, List, Tuple

import streamlit as st

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

try:
    from google import genai
    from google.genai import types as genai_types
except Exception:
    genai = None
    genai_types = None

try:
    from PyPDF2 import PdfReader
except Exception:
    PdfReader = None

try:
    import docx
except Exception:
    docx = None

try:
    import openpyxl
except Exception:
    openpyxl = None

try:
    from pptx import Presentation
except Exception:
    Presentation = None


APP_NAME = "ChatMD"
APP_VERSION = "V. 2026_05_24_4"
DEFAULT_PROVIDER = "Google Gemini"

FALLBACK_MODELS: Dict[str, List[str]] = {
    "Google Gemini": [
        "gemini-3-pro-preview",
        "gemini-2.5-pro",
        "gemini-2.5-flash",
        "gemini-2.0-flash",
    ],
    "OpenAI": [
        "gpt-4o",
        "gpt-4o-mini",
        "gpt-4.1",
        "gpt-4.1-mini",
    ],
}



IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp", ".gif"}
IMAGE_MIME_TYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
    ".gif": "image/gif",
}
MAX_INLINE_IMAGE_BYTES = 20 * 1024 * 1024

DEEP_RESEARCH_AGENT_FAST = "deep-research-preview-04-2026"
DEEP_RESEARCH_POLL_SECONDS = 10
DEEP_RESEARCH_MAX_POLLS = 30

DEFAULT_AGENTS = {
    "Bendras asistentas": {
        "icon": "💬",
        "system_prompt": "Tu esi naudingas, tikslus ir aiškiai lietuviškai atsakantis asistentas.",
        "examples": "",
    },
    "Darbo pagalbininkas": {
        "icon": "🧩",
        "system_prompt": "Padėk struktūruoti užduotis, rašyti aiškiai, trumpai ir profesionaliai lietuviškai.",
        "examples": "Atsakymo struktūra: trumpa išvada, tada konkretūs veiksmai, tada rizikos arba pastabos.",
    },
}


def get_secret(name: str) -> str:
    try:
        return str(st.secrets.get(name, "") or "").strip()
    except Exception:
        return ""


def get_api_key(provider: str) -> str:
    if provider == "Google Gemini":
        return get_secret("GOOGLE_API_KEY")
    if provider == "OpenAI":
        return get_secret("OPENAI_API_KEY")
    return ""


def render_css() -> None:
    st.markdown(
        """
        <style>
        html, body, [class*="css"] {
            font-size: 14px;
        }

        .stApp {
            background: #f7f9fc;
            color: #111827;
        }

        .main .block-container {
            max-width: 980px;
            padding-top: 1.2rem;
            padding-bottom: 7rem;
        }

        [data-testid="stSidebar"] {
            background: #eef3fa;
            border-right: 1px solid #d8e2f0;
        }

        [data-testid="stSidebar"] * {
            font-size: 0.92rem;
        }

        h1 {
            font-size: 1.55rem !important;
            margin-bottom: 0.25rem !important;
        }

        h2, h3 {
            font-size: 1.05rem !important;
        }

        .chat-title {
            display: flex;
            align-items: center;
            gap: 0.55rem;
            padding: 0.15rem 0 0.35rem 0;
        }

        .chat-logo {
            width: 38px;
            height: 38px;
            object-fit: contain;
        }

        .brand-main {
            font-size: 1.22rem;
            font-weight: 700;
            line-height: 1;
        }

        .brand-sub {
            font-size: 0.72rem;
            color: #526173;
            margin-top: 0.1rem;
        }

        .version-pill {
            display: inline-block;
            padding: 0.16rem 0.45rem;
            border-radius: 999px;
            background: #fff7ed;
            color: #9a3412;
            border: 1px solid #fed7aa;
            font-size: 0.72rem;
            margin-bottom: 0.6rem;
        }

        .status-pill {
            display: inline-block;
            padding: 0.17rem 0.5rem;
            border-radius: 999px;
            background: #dbeafe;
            color: #1d4ed8;
            font-size: 0.76rem;
            margin-right: 0.28rem;
            margin-bottom: 0.2rem;
        }

        .chat-area {
            margin-top: 1.1rem;
            padding-bottom: 1.5rem;
        }

        .msg-row {
            display: flex;
            width: 100%;
            margin: 0.45rem 0;
        }

        .msg-row.user {
            justify-content: flex-end;
        }

        .msg-bubble {
            max-width: 78%;
            padding: 0.62rem 0.78rem;
            border-radius: 16px;
            font-size: 0.96rem;
            line-height: 1.46;
            white-space: pre-wrap;
            word-wrap: break-word;
        }

        .msg-bubble.user {
            background: #e8f1ff;
            border: 1px solid #cfe0ff;
            color: #0f172a;
            border-bottom-right-radius: 6px;
        }

        .thinking-wrap {
            display: flex;
            justify-content: flex-start;
            width: 100%;
            margin: 0.45rem 0;
        }

        .thinking-card {
            display: inline-flex;
            align-items: center;
            gap: 0.65rem;
            padding: 0.62rem 0.78rem;
            border-radius: 16px;
            background: #ffffff;
            border: 1px solid #d8e2f0;
            box-shadow: 0 5px 18px rgba(15, 23, 42, 0.05);
            color: #334155;
            font-size: 0.92rem;
        }

        .thinking-spinner {
            width: 16px;
            height: 16px;
            border: 2px solid #cbd5e1;
            border-top-color: #2563eb;
            border-radius: 50%;
            animation: chatmd-spin 0.8s linear infinite;
        }

        .thinking-text {
            display: inline-flex;
            align-items: center;
            gap: 0.15rem;
        }

        .thinking-dots span {
            animation: chatmd-blink 1.2s infinite;
            opacity: 0.25;
        }

        .thinking-dots span:nth-child(2) {
            animation-delay: 0.2s;
        }

        .thinking-dots span:nth-child(3) {
            animation-delay: 0.4s;
        }

        .thinking-time {
            color: #64748b;
            font-size: 0.78rem;
            margin-top: 0.05rem;
        }

        @keyframes chatmd-spin {
            to {
                transform: rotate(360deg);
            }
        }

        @keyframes chatmd-blink {
            0%, 80%, 100% {
                opacity: 0.25;
            }
            40% {
                opacity: 1;
            }
        }

        [data-testid="stChatInput"] {
            max-width: 980px;
            margin: 0 auto;
        }

        [data-testid="stChatInput"] > div {
            background: #eef4fb !important;
            border: 1px solid #cbdcf0 !important;
            border-radius: 14px !important;
        }

        [data-testid="stChatInput"] button {
            border-radius: 10px !important;
        }

        [data-testid="stChatInput"] button:first-child {
            background: #dbeafe !important;
            border-right: 1px solid #b9cce8 !important;
            margin-right: 0.35rem !important;
        }

        .login-wrap {
            display: flex;
            justify-content: center;
            margin-top: 12vh;
        }

        .login-card {
            width: 100%;
            max-width: 420px;
            background: white;
            border: 1px solid #d8e2f0;
            border-radius: 20px;
            padding: 1.2rem;
            box-shadow: 0 10px 30px rgba(15, 23, 42, 0.08);
            text-align: center;
        }

        .login-title {
            font-size: 1.55rem;
            font-weight: 750;
            margin-bottom: 0.35rem;
        }

        .login-subtitle {
            font-size: 0.9rem;
            color: #64748b;
        }

        @media (max-width: 720px) {
            .main .block-container {
                padding-left: 0.75rem;
                padding-right: 0.75rem;
            }

            .msg-bubble {
                max-width: 92%;
                font-size: 0.95rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def check_password() -> bool:
    app_password = get_secret("APP_PASSWORD")

    if not app_password:
        st.error(
            "APP_PASSWORD nėra nustatytas Streamlit Secrets. "
            "Programa saugumo sumetimais nerodoma."
        )
        return False

    if st.session_state.get("authenticated", False):
        return True

    render_css()

    st.markdown(
        """
        <div class="login-wrap">
            <div class="login-card">
                <div class="login-title">🔐 ChatMD</div>
                <div class="login-subtitle">Įvesk slaptažodį, kad galėtum naudotis programa.</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    left, center, right = st.columns([1, 0.42, 1])

    with center:
        entered_password = st.text_input(
            "Slaptažodis",
            type="password",
            placeholder="Slaptažodis",
            label_visibility="collapsed",
        )

        if st.button("Prisijungti", use_container_width=True):
            if hmac.compare_digest(entered_password, app_password):
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Neteisingas slaptažodis.")

    return False


def init_state() -> None:
    st.session_state.setdefault("provider", DEFAULT_PROVIDER)
    st.session_state.setdefault("models_cache", {})
    st.session_state.setdefault("model", FALLBACK_MODELS[DEFAULT_PROVIDER][0])
    st.session_state.setdefault("web_search_enabled", True)
    st.session_state.setdefault("temperature", 0.7)
    st.session_state.setdefault("mode", "chat")
    st.session_state.setdefault("agent_name", "Bendras asistentas")
    st.session_state.setdefault("agents", DEFAULT_AGENTS.copy())
    st.session_state.setdefault("messages", [])
    st.session_state.setdefault("chat_counter", 1)
    st.session_state.setdefault("history", [])
    st.session_state.setdefault("authenticated", False)


def reset_chat() -> None:
    if st.session_state.messages:
        title = next(
            (m["content"] for m in st.session_state.messages if m["role"] == "user"),
            "Pokalbis",
        )
        st.session_state.history.insert(
            0,
            {
                "title": title[:60],
                "messages": list(st.session_state.messages),
                "created_at": dt.datetime.now().isoformat(timespec="minutes"),
                "mode": st.session_state.mode,
                "agent": st.session_state.agent_name,
            },
        )
    st.session_state.messages = []
    st.session_state.chat_counter += 1


def runtime_context() -> str:
    today = dt.datetime.now().strftime("%Y-%m-%d")
    return (
        f"Šiandienos data: {today}. "
        "Jei klausimas apie dabartinius įvykius ar faktus, naudok interneto paiešką, kai ji įjungta. "
        "Jei paieška nepavyksta, aiškiai tai pasakyk."
    )


def list_models(provider: str) -> Tuple[List[str], str]:
    api_key = get_api_key(provider)

    if not api_key:
        return FALLBACK_MODELS.get(provider, []), "fallback: nėra API rakto"

    if provider == "Google Gemini":
        if genai is None:
            return FALLBACK_MODELS[provider], "fallback: neįdiegta google-genai"

        try:
            client = genai.Client(api_key=api_key)
            names = []

            for model in client.models.list():
                name = getattr(model, "name", "") or ""

                if name.startswith("models/"):
                    name = name.split("/", 1)[1]

                if name and "embedding" not in name.lower():
                    names.append(name)

            return sorted(set(names)) or FALLBACK_MODELS[provider], "live"
        except Exception as exc:
            return FALLBACK_MODELS[provider], f"fallback: {exc}"

    if provider == "OpenAI":
        if OpenAI is None:
            return FALLBACK_MODELS[provider], "fallback: neįdiegta openai"

        try:
            client = OpenAI(api_key=api_key)
            models = client.models.list()
            names = sorted({m.id for m in models.data if getattr(m, "id", None)})
            names = [m for m in names if m.startswith(("gpt-", "o", "chatgpt-"))]
            return names or FALLBACK_MODELS[provider], "live"
        except Exception as exc:
            return FALLBACK_MODELS[provider], f"fallback: {exc}"

    return [], "fallback"


class PathLike:
    def __init__(self, name: str) -> None:
        self.name = name

    @property
    def suffix(self) -> str:
        if "." not in self.name:
            return ""
        return "." + self.name.rsplit(".", 1)[1].lower()


def is_image_file(uploaded_file) -> bool:
    return PathLike(uploaded_file.name).suffix in IMAGE_SUFFIXES


def get_image_mime_type(uploaded_file) -> str:
    suffix = PathLike(uploaded_file.name).suffix
    return IMAGE_MIME_TYPES.get(suffix, "application/octet-stream")


def make_image_note(uploaded_files) -> str:
    if not uploaded_files:
        return ""

    image_names = [f.name for f in uploaded_files if is_image_file(f)]

    if not image_names:
        return ""

    return "Prisegti vaizdai, perduoti modeliui kaip tikri multimodal vaizdo duomenys:\n" + "\n".join(
        f"- {name}" for name in image_names
    )


def make_google_contents(messages: List[Dict[str, str]], files_context: str, uploaded_files) -> List[Any]:
    text_parts = [f"Sisteminė instrukcija:\n{build_system_prompt()}"]

    for m in messages[-20:]:
        role = "Vartotojas" if m["role"] == "user" else "Asistentas"
        text_parts.append(f"{role}: {m['content']}")

    if files_context:
        text_parts.append("Prisegtų failų tekstinis turinys:\n" + files_context)

    image_note = make_image_note(uploaded_files)

    if image_note:
        text_parts.append(image_note)

    contents: List[Any] = ["\n\n".join(text_parts)]

    if genai_types is not None and uploaded_files:
        for uploaded_file in uploaded_files:
            if not is_image_file(uploaded_file):
                continue

            image_bytes = uploaded_file.getvalue()

            if len(image_bytes) > MAX_INLINE_IMAGE_BYTES:
                contents.append(
                    f"[Vaizdas {uploaded_file.name} per didelis inline siuntimui: "
                    f"{len(image_bytes) / (1024 * 1024):.1f} MB. Didžiausia riba: 20 MB.]"
                )
                continue

            contents.append(
                genai_types.Part.from_bytes(
                    data=image_bytes,
                    mime_type=get_image_mime_type(uploaded_file),
                )
            )

    return contents


def make_openai_input_messages(
    messages: List[Dict[str, str]],
    files_context: str,
    uploaded_files,
) -> List[Dict[str, Any]]:
    input_messages: List[Dict[str, Any]] = [
        {"role": "system", "content": build_system_prompt()}
    ]
    input_messages.extend(messages[-20:])

    content_parts: List[Dict[str, str]] = []

    if files_context:
        content_parts.append(
            {
                "type": "input_text",
                "text": "Prisegtų failų tekstinis turinys:\n" + files_context,
            }
        )

    image_note = make_image_note(uploaded_files)

    if image_note:
        content_parts.append({"type": "input_text", "text": image_note})

    if uploaded_files:
        for uploaded_file in uploaded_files:
            if not is_image_file(uploaded_file):
                continue

            image_bytes = uploaded_file.getvalue()

            if len(image_bytes) > MAX_INLINE_IMAGE_BYTES:
                content_parts.append(
                    {
                        "type": "input_text",
                        "text": (
                            f"[Vaizdas {uploaded_file.name} per didelis inline siuntimui: "
                            f"{len(image_bytes) / (1024 * 1024):.1f} MB. Didžiausia riba: 20 MB.]"
                        ),
                    }
                )
                continue

            encoded = base64.b64encode(image_bytes).decode("utf-8")
            mime_type = get_image_mime_type(uploaded_file)
            data_url = f"data:{mime_type};base64,{encoded}"
            content_parts.append({"type": "input_image", "image_url": data_url})

    if content_parts:
        input_messages.append({"role": "user", "content": content_parts})

    return input_messages


def read_uploaded_file(uploaded_file) -> str:
    name = uploaded_file.name
    suffix = PathLike(name).suffix
    data = uploaded_file.getvalue()

    try:
        if suffix in IMAGE_SUFFIXES:
            return f"[Vaizdas {name}] Vaizdas bus perduotas modeliui kaip multimodal image input, ne kaip tekstinė base64 ištrauka."

        if suffix == ".pdf":
            if PdfReader is None:
                return f"[{name}] PDF nuskaitymui reikia PyPDF2."

            import io

            reader = PdfReader(io.BytesIO(data))
            pages = [(p.extract_text() or "") for p in reader.pages[:20]]
            return f"[{name}]\n" + "\n".join(pages)[:15000]

        if suffix == ".docx":
            if docx is None:
                return f"[{name}] DOCX nuskaitymui reikia python-docx."

            import io

            document = docx.Document(io.BytesIO(data))
            text = "\n".join(p.text for p in document.paragraphs if p.text.strip())
            return f"[{name}]\n{text[:15000]}"

        if suffix == ".xlsx":
            if openpyxl is None:
                return f"[{name}] XLSX nuskaitymui reikia openpyxl."

            import io

            wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
            rows = []

            for ws in wb.worksheets[:3]:
                rows.append(f"Lapas: {ws.title}")

                for row in ws.iter_rows(max_row=80, values_only=True):
                    vals = [str(v) for v in row if v is not None]

                    if vals:
                        rows.append(" | ".join(vals))

            return f"[{name}]\n" + "\n".join(rows)[:15000]

        if suffix == ".pptx":
            if Presentation is None:
                return f"[{name}] PPTX nuskaitymui reikia python-pptx."

            import io

            prs = Presentation(io.BytesIO(data))
            slides = []

            for i, slide in enumerate(prs.slides[:30], 1):
                texts = []

                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text:
                        texts.append(shape.text)

                if texts:
                    slides.append(f"Skaidrė {i}:\n" + "\n".join(texts))

            return f"[{name}]\n" + "\n\n".join(slides)[:15000]

        return f"[{name}]\n" + data.decode("utf-8", errors="ignore")[:15000]
    except Exception as exc:
        return f"[{name}] Nepavyko nuskaityti: {exc}"


def build_system_prompt() -> str:
    base = runtime_context()

    if st.session_state.mode == "agent":
        agent = st.session_state.agents.get(
            st.session_state.agent_name,
            DEFAULT_AGENTS["Bendras asistentas"],
        )
        prompt = agent.get("system_prompt", "")
        examples = agent.get("examples", "")

        if examples.strip():
            prompt += "\n\nAgento pavyzdžiai / žinios:\n" + examples.strip()

        return base + "\n\n" + prompt

    if st.session_state.mode == "thinking":
        return (
            base
            + "\n\nTu esi ChatMD Mąstymo režime. Atsakyk giliau nei paprastame pokalbyje: "
            "pirmiausia įvertink klausimo esmę, tada pateik struktūruotą, aiškų ir praktišką atsakymą. "
            "Jei reikia, naudok interneto paiešką, kai ji įjungta. Neužtęsk be reikalo, bet parodyk svarbiausią logiką ir išvadas. "
            "Atsakyk lietuviškai."
        )

    if st.session_state.mode == "deep":
        return (
            base
            + "\n\nTu esi ChatMD Gilios analizės režime. Atlik išsamesnį tyrimą lietuviškai: "
            "pateik aiškią struktūrą, svarbiausias išvadas, argumentus ir šaltinius, jei jie prieinami."
        )

    return base + "\n\nTu esi naudingas, tikslus ir aiškiai lietuviškai atsakantis asistentas."


def make_files_context(uploaded_files) -> str:
    if not uploaded_files:
        return ""

    chunks = [read_uploaded_file(f) for f in uploaded_files]
    return "\n\n".join(chunks)



def make_research_prompt(messages: List[Dict[str, str]], files_context: str, uploaded_files) -> str:
    user_question = ""

    for m in reversed(messages):
        if m.get("role") == "user":
            user_question = str(m.get("content", ""))
            break

    research_prompt = [
        build_system_prompt(),
        "",
        "Atlik gilią analizę lietuviškai. Pateik aiškią struktūrą, svarbiausias išvadas ir šaltinius, jei agentas juos turi.",
        "",
        "Vartotojo klausimas:",
        user_question,
    ]

    if files_context:
        research_prompt.extend(["", "Prisegtų failų tekstinis turinys:", files_context])

    image_names = [f.name for f in uploaded_files or [] if is_image_file(f)]

    if image_names:
        research_prompt.extend(
            [
                "",
                "Pastaba: ši pirmoji Gilios analizės versija vaizdų tiesiogiai Deep Research agentui dar neperduoda.",
                "Prisegti vaizdai: " + ", ".join(image_names),
            ]
        )

    return "\n".join(research_prompt).strip()


def extract_interaction_text(interaction) -> str:
    output_text = getattr(interaction, "output_text", "") or ""

    if output_text:
        return output_text

    outputs = getattr(interaction, "outputs", None) or []

    if outputs:
        last_output = outputs[-1]
        text = getattr(last_output, "text", "") or ""

        if text:
            return text

    steps = getattr(interaction, "steps", None) or []

    for step in reversed(steps):
        text = getattr(step, "text", "") or ""

        if text:
            return text

        output = getattr(step, "output", None)
        text = getattr(output, "text", "") or ""

        if text:
            return text

    return ""


def stream_gemini_deep_research(files_context: str, uploaded_files=None):
    api_key = get_api_key("Google Gemini")

    if not api_key:
        yield "🔴 Google Gemini API raktas neįvestas Streamlit Secrets."
        return

    if genai is None:
        yield "🔴 Deep Research reikia google-genai bibliotekos."
        return

    if st.session_state.provider != "Google Gemini":
        yield "🔴 Gili analizė šiuo metu veikia tik pasirinkus Google Gemini teikėją."
        return

    agent = DEEP_RESEARCH_AGENT_FAST
    mode_label = "Gili analizė"

    client = genai.Client(api_key=api_key)
    research_input = make_research_prompt(st.session_state.messages, files_context, uploaded_files or [])

    try:
        interaction = client.interactions.create(
            input=research_input,
            agent=agent,
            background=True,
        )
    except Exception as exc:
        yield f"⚠️ Nepavyko paleisti {mode_label} užduoties: {exc}"
        return

    interaction_id = getattr(interaction, "id", "") or ""

    yield (
        f"🔎 **{mode_label} pradėta.**\n\n"
        "Gemini Deep Research renka informaciją ir ruošia atsakymą. "
        "Tai gali trukti ilgiau nei paprastas pokalbis.\n\n"
    )

    for _ in range(DEEP_RESEARCH_MAX_POLLS):
        time.sleep(DEEP_RESEARCH_POLL_SECONDS)

        try:
            interaction = client.interactions.get(interaction_id)
        except Exception as exc:
            yield f"\n⚠️ Nepavyko patikrinti tyrimo būsenos: {exc}"
            return

        status = str(getattr(interaction, "status", "") or "").lower()

        if status == "completed":
            result = extract_interaction_text(interaction)
            yield "\n✅ **Tyrimas baigtas.**\n\n"
            yield result or "Gautas tuščias Deep Research atsakymas."
            return

        if status == "failed":
            error = getattr(interaction, "error", "") or "Nežinoma klaida."
            yield f"\n⚠️ Tyrimas nepavyko: {error}"
            return

        yield "⏳ "

    yield (
        "\n⚠️ Tyrimas dar nebaigtas per numatytą laiką. "
        "Pabandyk trumpesnį klausimą arba naudok Mąstymas režimą, kuris yra greitesnis ir neveikia per Deep Research."
    )

def answer_google(model: str, messages: List[Dict[str, str]], files_context: str, uploaded_files=None) -> str:
    api_key = get_api_key("Google Gemini")

    if not api_key:
        return "🔴 Google Gemini API raktas neįvestas Streamlit Secrets."

    client = genai.Client(api_key=api_key)
    contents = make_google_contents(messages, files_context, uploaded_files or [])

    config = {
        "temperature": st.session_state.temperature,
    }

    if st.session_state.web_search_enabled and genai_types is not None:
        try:
            config = genai_types.GenerateContentConfig(
                temperature=st.session_state.temperature,
                tools=[genai_types.Tool(google_search=genai_types.GoogleSearch())],
            )
        except Exception:
            pass

    response = client.models.generate_content(
        model=model,
        contents=contents,
        config=config,
    )

    return getattr(response, "text", "") or "Gautas tuščias atsakymas."


def stream_google(model: str, messages: List[Dict[str, str]], files_context: str, uploaded_files=None):
    api_key = get_api_key("Google Gemini")

    if not api_key:
        yield "🔴 Google Gemini API raktas neįvestas Streamlit Secrets."
        return

    client = genai.Client(api_key=api_key)
    contents = make_google_contents(messages, files_context, uploaded_files or [])

    config = {
        "temperature": st.session_state.temperature,
    }

    if st.session_state.web_search_enabled and genai_types is not None:
        try:
            config = genai_types.GenerateContentConfig(
                temperature=st.session_state.temperature,
                tools=[genai_types.Tool(google_search=genai_types.GoogleSearch())],
            )
        except Exception:
            pass

    try:
        stream = client.models.generate_content_stream(
            model=model,
            contents=contents,
            config=config,
        )

        for chunk in stream:
            text = getattr(chunk, "text", "") or ""
            if text:
                yield text
    except Exception:
        yield answer_google(model, messages, files_context, uploaded_files)


def answer_openai(model: str, messages: List[Dict[str, str]], files_context: str, uploaded_files=None) -> str:
    api_key = get_api_key("OpenAI")

    if not api_key:
        return "🔴 OpenAI API raktas neįvestas Streamlit Secrets."

    client = OpenAI(api_key=api_key)
    input_messages = make_openai_input_messages(messages, files_context, uploaded_files or [])

    kwargs = {
        "model": model,
        "input": input_messages,
    }

    if not model.startswith(("gpt-5", "o1", "o3", "o4")):
        kwargs["temperature"] = st.session_state.temperature

    if st.session_state.web_search_enabled:
        kwargs["tools"] = [{"type": "web_search"}]

    try:
        response = client.responses.create(**kwargs)
    except Exception as exc:
        msg = str(exc)

        if "temperature" in msg and "Unsupported parameter" in msg:
            kwargs.pop("temperature", None)
            response = client.responses.create(**kwargs)
        elif "web_search" in msg or "tool" in msg.lower():
            kwargs.pop("tools", None)
            response = client.responses.create(**kwargs)
        else:
            raise

    return getattr(response, "output_text", "") or "Gautas tuščias atsakymas."


def stream_openai(model: str, messages: List[Dict[str, str]], files_context: str, uploaded_files=None):
    api_key = get_api_key("OpenAI")

    if not api_key:
        yield "🔴 OpenAI API raktas neįvestas Streamlit Secrets."
        return

    client = OpenAI(api_key=api_key)
    input_messages = make_openai_input_messages(messages, files_context, uploaded_files or [])

    kwargs = {
        "model": model,
        "input": input_messages,
        "stream": True,
    }

    if not model.startswith(("gpt-5", "o1", "o3", "o4")):
        kwargs["temperature"] = st.session_state.temperature

    if st.session_state.web_search_enabled:
        kwargs["tools"] = [{"type": "web_search"}]

    try:
        stream = client.responses.create(**kwargs)

        for event in stream:
            event_type = getattr(event, "type", "")

            if event_type == "response.output_text.delta":
                delta = getattr(event, "delta", "") or ""
                if delta:
                    yield delta

    except Exception as exc:
        msg = str(exc)

        try:
            if "temperature" in msg and "Unsupported parameter" in msg:
                kwargs.pop("temperature", None)
                stream = client.responses.create(**kwargs)

                for event in stream:
                    event_type = getattr(event, "type", "")

                    if event_type == "response.output_text.delta":
                        delta = getattr(event, "delta", "") or ""
                        if delta:
                            yield delta

            elif "web_search" in msg or "tool" in msg.lower():
                kwargs.pop("tools", None)
                stream = client.responses.create(**kwargs)

                for event in stream:
                    event_type = getattr(event, "type", "")

                    if event_type == "response.output_text.delta":
                        delta = getattr(event, "delta", "") or ""
                        if delta:
                            yield delta
            else:
                raise
        except Exception:
            kwargs.pop("stream", None)
            yield answer_openai(model, messages, files_context, uploaded_files)


def stream_answer(uploaded_files):
    files_context = make_files_context(uploaded_files)
    provider = st.session_state.provider
    model = st.session_state.model

    try:
        if st.session_state.mode == "deep":
            yield from stream_gemini_deep_research(files_context, uploaded_files)
            return

        if provider == "Google Gemini":
            yield from stream_google(model, st.session_state.messages, files_context, uploaded_files)
            return

        if provider == "OpenAI":
            yield from stream_openai(model, st.session_state.messages, files_context, uploaded_files)
            return

        yield "Nežinomas teikėjas."
    except Exception as exc:
        yield f"⚠️ Klaida: {exc}"


def render_brand() -> None:
    logo_path = "chatmd_logo.png"
    logo_html = ""

    try:
        with open(logo_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")
        logo_html = f'<img class="chat-logo" src="data:image/png;base64,{b64}" />'
    except Exception:
        logo_html = ""

    st.sidebar.markdown(
        f"""
        <div class="chat-title">
            {logo_html}
            <div>
                <div class="brand-main">ChatMD</div>
                <div class="brand-sub">AI pokalbiai ir agentai</div>
            </div>
        </div>
        <div class="version-pill">{APP_VERSION}</div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar() -> None:
    render_brand()

    if st.sidebar.button("＋ Naujas pokalbis", use_container_width=True):
        reset_chat()
        st.rerun()

    st.sidebar.divider()

    st.sidebar.subheader("Režimas")
    mode = st.sidebar.radio(
        "Pasirink režimą",
        ["chat", "thinking", "deep"],
        format_func=lambda x: {
            "chat": "💬 Pokalbis",
            "thinking": "🧠 Mąstymas",
            "deep": "🔎 Gili analizė",
        }.get(x, x),
        horizontal=False,
        label_visibility="collapsed",
    )

    if mode != st.session_state.mode:
        reset_chat()
        st.session_state.mode = mode
        st.rerun()

    if st.session_state.mode == "thinking":
        st.sidebar.info("Mąstymas atsako giliau, bet veikia greitai ir nenaudoja Deep Research.")

    if st.session_state.mode == "deep":
        st.sidebar.info("Gili analizė naudoja Gemini Deep Research. Ji gali trukti kelias minutes. Pasirink Google Gemini teikėją.")

    st.sidebar.divider()

    st.sidebar.subheader("AI nustatymai")

    provider = st.sidebar.selectbox(
        "Teikėjas",
        ["Google Gemini", "OpenAI"],
        index=["Google Gemini", "OpenAI"].index(st.session_state.provider),
    )

    if provider != st.session_state.provider:
        st.session_state.provider = provider
        models, _ = list_models(provider)
        st.session_state.model = models[0] if models else ""
        reset_chat()
        st.rerun()

    key = get_api_key(st.session_state.provider)
    st.sidebar.markdown("🟢 API raktas aktyvus" if key else "🔴 API raktas neįvestas")

    if st.sidebar.button("Atnaujinti modelių sąrašą", use_container_width=True):
        models, source = list_models(st.session_state.provider)
        st.session_state.models_cache[st.session_state.provider] = models

        if st.session_state.model not in models:
            st.session_state.model = models[0] if models else ""

        st.sidebar.info(f"Modeliai: {source}")

    models = st.session_state.models_cache.get(st.session_state.provider)

    if not models:
        models, _ = list_models(st.session_state.provider)
        st.session_state.models_cache[st.session_state.provider] = models

    if st.session_state.model not in models and models:
        st.session_state.model = models[0]

    st.session_state.model = st.sidebar.selectbox(
        "Modelis",
        models,
        index=models.index(st.session_state.model) if st.session_state.model in models else 0,
    )

    st.session_state.web_search_enabled = st.sidebar.toggle(
        "Interneto paieška",
        value=st.session_state.web_search_enabled,
    )

    st.session_state.temperature = st.sidebar.slider(
        "Kūrybiškumas",
        0.0,
        2.0,
        float(st.session_state.temperature),
        0.1,
    )

    st.sidebar.divider()
    st.sidebar.subheader("Istorija šioje sesijoje")

    if not st.session_state.history:
        st.sidebar.caption("Istorija atsiras pradėjus naują pokalbį.")

    for i, chat in enumerate(st.session_state.history[:15]):
        col1, col2 = st.sidebar.columns([0.82, 0.18])

        with col1:
            if st.button(chat["title"], key=f"hist_{i}", use_container_width=True):
                st.session_state.messages = list(chat["messages"])
                st.session_state.mode = chat.get("mode", "chat")
                if st.session_state.mode == "agent":
                    st.session_state.mode = "chat"
                st.session_state.agent_name = chat.get("agent", "Bendras asistentas")
                st.rerun()

        with col2:
            if st.button("🗑", key=f"del_{i}"):
                st.session_state.history.pop(i)
                st.rerun()


def parse_chat_input(chat_value):
    if not chat_value:
        return "", []

    if isinstance(chat_value, str):
        return chat_value.strip(), []

    text = getattr(chat_value, "text", "")
    files = getattr(chat_value, "files", [])

    if not text and isinstance(chat_value, dict):
        text = chat_value.get("text", "")
        files = chat_value.get("files", [])

    return str(text or "").strip(), list(files or [])


def render_user_message(content: str) -> None:
    safe_content = html.escape(content).replace("\n", "<br>")

    st.markdown(
        f"""
        <div class="msg-row user">
            <div class="msg-bubble user">{safe_content}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_chat_message(role: str, content: str) -> None:
    if role == "user":
        render_user_message(content)
    else:
        st.markdown(content)


def render_thinking(placeholder) -> None:
    placeholder.markdown(
        """
        <div class="thinking-wrap">
            <div class="thinking-card">
                <div class="thinking-spinner"></div>
                <div>
                    <div class="thinking-text">
                        Galvoju<span class="thinking-dots"><span>.</span><span>.</span><span>.</span></span>
                    </div>
                    <div class="thinking-time">Ruošiu atsakymą</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="💬",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    init_state()

    if not check_password():
        return

    if st.session_state.mode == "agent":
        st.session_state.mode = "chat"

    render_css()
    render_sidebar()

    mode_label = {
        "chat": "Pokalbis",
        "thinking": "Mąstymas",
        "deep": "Gili analizė",
    }.get(st.session_state.mode, "Pokalbis")

    st.markdown("### ChatMD")
    st.markdown(
        f"""
        <span class="version-pill">{APP_VERSION}</span>
        <br>
        <span class="status-pill">{st.session_state.provider}</span>
        <span class="status-pill">{st.session_state.model}</span>
        <span class="status-pill">{mode_label}</span>
        <span class="status-pill">Internetas: {"ON" if st.session_state.web_search_enabled else "OFF"}</span>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="chat-area">', unsafe_allow_html=True)

    for m in st.session_state.messages:
        render_chat_message(m["role"], m["content"])

    st.markdown("</div>", unsafe_allow_html=True)

    chat_value = st.chat_input(
        "Rašykite žinutę...",
        accept_file="multiple",
        file_type=[
            "png",
            "jpg",
            "jpeg",
            "webp",
            "gif",
            "pdf",
            "docx",
            "xlsx",
            "pptx",
            "txt",
            "csv",
            "md",
            "py",
            "json",
        ],
    )

    prompt, uploaded_files = parse_chat_input(chat_value)

    if prompt or uploaded_files:
        if not prompt and uploaded_files:
            prompt = "Peržiūrėk prisegtus failus."

        visible_prompt = prompt

        if uploaded_files:
            file_names = ", ".join([f.name for f in uploaded_files])
            visible_prompt = f"{prompt}\n\n📎 Prisegti failai: {file_names}"

        st.session_state.messages.append({"role": "user", "content": visible_prompt})

        render_user_message(visible_prompt)

        assistant_placeholder = st.empty()
        render_thinking(assistant_placeholder)

        streamed_answer = ""

        for chunk in stream_answer(uploaded_files):
            streamed_answer += chunk
            assistant_placeholder.markdown(streamed_answer + "▌")

        final_answer = streamed_answer.strip() or "Gautas tuščias atsakymas."
        assistant_placeholder.markdown(final_answer)

        st.session_state.messages.append({"role": "assistant", "content": final_answer})


if __name__ == "__main__":
    main()